import type { ISharePointClient, ISharePointFileInfo } from '../src/core/services/SharePointClient';

/**
 * Bản giả lập SharePoint trong bộ nhớ (thư viện tài liệu + list) cho khung chạy thử cục bộ (dev/).
 * Khi đóng gói khung chạy thử, module `core/services/SharePointClient` được thay bằng file này nên code module
 * chạy nguyên bản mà không cần SharePoint thật. KHÔNG thuộc bản đóng gói web part.
 */

interface MemFile {
  name: string;
  bytes: Uint8Array;
  modified: string;
  etag: number;
}

const files: Record<string, MemFile> = {};
const folders: Record<string, true> = {};
const lists: Record<string, Array<Record<string, unknown>>> = {};

const norm = (p: string): string => p.replace(/\/+$/, '');
const parentOf = (p: string): string => p.slice(0, p.lastIndexOf('/'));
const nameOf = (p: string): string => p.slice(p.lastIndexOf('/') + 1);

function toBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function ensure(folder: string): void {
  let cur = norm(folder);
  while (cur && !folders[cur]) {
    folders[cur] = true;
    cur = parentOf(cur);
  }
}

/** Nạp file có sẵn (base64) vào thư viện giả lập. */
export function seedFile(path: string, base64: string): void {
  const bin = window.atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  ensure(parentOf(path));
  files[path] = { name: nameOf(path), bytes, modified: new Date().toISOString(), etag: 1 };
}

export function seedList(name: string, rows: Array<Record<string, unknown>>): void {
  lists[name] = rows.map((r, i) => ({ Id: i + 1, ...r }));
}

/** Danh sách file hiện có (để kiểm tra sau khi lưu trong khung chạy thử). */
export function dumpFiles(): Array<{ path: string; size: number }> {
  return Object.keys(files).sort().map((p) => ({ path: p, size: files[p].bytes.length }));
}

export function readText(path: string): string | undefined {
  const f = files[path];
  return f ? new TextDecoder('utf-8').decode(f.bytes) : undefined;
}

export function dumpList(name: string): Array<Record<string, unknown>> {
  return (lists[name] || []).slice();
}

/* ------------------------------ OData tối giản ------------------------------ */

type Tok = { t: 'id' | 'str' | 'num' | 'op' | 'par'; v: string };

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === '(' || c === ')') { out.push({ t: 'par', v: c }); i++; continue; }
    if (c === "'") {
      let j = i + 1;
      let s = '';
      while (j < src.length) {
        if (src[j] === "'" && src[j + 1] === "'") { s += "'"; j += 2; continue; }
        if (src[j] === "'") break;
        s += src[j++];
      }
      out.push({ t: 'str', v: s });
      i = j + 1;
      continue;
    }
    const m = /^[A-Za-z0-9_.\-]+/.exec(src.slice(i));
    if (!m) throw new Error(`OData: ký tự lạ "${c}"`);
    const w = m[0];
    out.push({ t: /^-?\d+(\.\d+)?$/.test(w) ? 'num' : /^(eq|ne|gt|ge|lt|le|and|or)$/i.test(w) ? 'op' : 'id', v: w });
    i += w.length;
  }
  return out;
}

function evalFilter(filter: string, row: Record<string, unknown>): boolean {
  const toks = tokenize(filter);
  let pos = 0;
  const peek = (): Tok | undefined => toks[pos];
  const operand = (): string | number => {
    const t = toks[pos++];
    if (t.t === 'str') return t.v;
    if (t.t === 'num') return parseFloat(t.v);
    const v = row[t.v];
    return v === undefined || v === null ? '' : (v as string | number);
  };
  const cmp = (): boolean => {
    const t = peek();
    if (t && t.t === 'par' && t.v === '(') {
      pos++;
      const v = or();
      pos++; // ')'
      return v;
    }
    const a = operand();
    const op = toks[pos++].v.toLowerCase();
    const b = operand();
    const x = typeof a === 'number' || typeof b === 'number' ? Number(a) : String(a);
    const y = typeof a === 'number' || typeof b === 'number' ? Number(b) : String(b);
    return op === 'eq' ? x === y : op === 'ne' ? x !== y : op === 'gt' ? x > y : op === 'ge' ? x >= y : op === 'lt' ? x < y : x <= y;
  };
  const and = (): boolean => {
    let v = cmp();
    while (peek() && peek()!.t === 'op' && peek()!.v.toLowerCase() === 'and') { pos++; const r = cmp(); v = v && r; }
    return v;
  };
  const or = (): boolean => {
    let v = and();
    while (peek() && peek()!.t === 'op' && peek()!.v.toLowerCase() === 'or') { pos++; const r = and(); v = v || r; }
    return v;
  };
  return or();
}

/** Cột chuẩn của các list; "Hiển thị=Nội bộ" khi tên nội bộ khác tên hiển thị (vd DisplayName là cột Title). */
const FIELD_MAPS: Record<string, string[]> = {
  ReportRole: ['DisplayName=Title', 'Email', 'Unit', 'DocID', 'DocTitle', 'Active'],
  ReportRegistry: ['Title', 'Unit', 'DocID', 'ReportDate', 'ReportNo', 'Reporter', 'Status', 'SavedAt', 'SignedAt', 'Counters', 'SignedUploadBy', 'SignedRemovedAt', 'SignedRemovedBy'],
  CatalogItems: ['Title', 'Unit', 'Catalog', 'Code', 'Link']
};

export class SharePointClient implements ISharePointClient {
  // Tham số (SPHttpClient thật) bị bỏ qua: mọi thứ nằm trong bộ nhớ
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  constructor(_spHttpClient?: unknown) {}

  public getItems<T>(_siteUrl: string, listTitle: string): Promise<T[]> {
    if (!lists[listTitle]) return Promise.reject(new Error(`SharePoint API Error (404): list '${listTitle}' does not exist`));
    return Promise.resolve(lists[listTitle].slice() as unknown as T[]);
  }

  public queryItems<T>(_siteUrl: string, listTitle: string, query: string): Promise<T[]> {
    const params: Record<string, string> = {};
    query.split(/&(?=\$)/).forEach((p) => {
      const i = p.indexOf('=');
      if (i > 0) params[p.slice(0, i)] = p.slice(i + 1);
    });
    if (!lists[listTitle]) return Promise.reject(new Error(`SharePoint API Error (404): list '${listTitle}' does not exist`));
    let rows = lists[listTitle].slice();
    if (params.$filter) rows = rows.filter((r) => evalFilter(params.$filter, r));
    if (params.$orderby) {
      const [col, dir] = params.$orderby.trim().split(/\s+/);
      rows.sort((a, b) => {
        const x = Number(a[col]) || 0;
        const y = Number(b[col]) || 0;
        return dir && dir.toLowerCase() === 'desc' ? y - x : x - y;
      });
    }
    if (params.$top) rows = rows.slice(0, parseInt(params.$top, 10));
    return Promise.resolve(rows as unknown as T[]);
  }

  public getListFieldMap(_siteUrl: string, listTitle: string): Promise<Record<string, string>> {
    const map: Record<string, string> = {};
    (FIELD_MAPS[listTitle] || []).forEach((c) => {
      const [display, internal] = c.split('=');
      map[display.toLowerCase()] = internal || display;
    });
    return Promise.resolve(map);
  }

  public createItem<T>(_siteUrl: string, listTitle: string, data: Partial<T>): Promise<void> {
    const list = lists[listTitle] || (lists[listTitle] = []);
    const key = (data as Record<string, unknown>).Title;
    if (key !== undefined && list.some((r) => r.Title === key)) return Promise.reject(new Error('Duplicate value (409) unique'));
    list.push({ Id: list.length + 1, ...(data as Record<string, unknown>) });
    return Promise.resolve();
  }

  public updateItem<T>(_siteUrl: string, listTitle: string, id: number, data: Partial<T>): Promise<void> {
    const row = (lists[listTitle] || []).filter((r) => r.Id === id)[0];
    if (!row) return Promise.reject(new Error('Item not found (404)'));
    Object.assign(row, data);
    return Promise.resolve();
  }

  public getFileWithMeta(_siteUrl: string, path: string): Promise<{ buffer: ArrayBuffer; eTag: string } | undefined> {
    const f = files[norm(path)];
    return Promise.resolve(f ? { buffer: toBuffer(f.bytes), eTag: `"${f.etag}"` } : undefined);
  }

  public getFileBuffer(_siteUrl: string, path: string): Promise<ArrayBuffer | undefined> {
    const f = files[norm(path)];
    return Promise.resolve(f ? toBuffer(f.bytes) : undefined);
  }

  public uploadFile(
    _siteUrl: string,
    folder: string,
    fileName: string,
    buffer: ArrayBuffer,
    overwrite = true,
    eTag?: string
  ): Promise<{ serverRelativeUrl: string; eTag?: string }> {
    const path = `${norm(folder)}/${fileName}`;
    const existing = files[path];
    if (existing && !overwrite) return Promise.reject(new Error('File exists (409)'));
    if (existing && eTag && eTag !== `"${existing.etag}"`) return Promise.reject(new Error('Precondition failed (412): file đã bị người khác sửa'));
    ensure(folder);
    const etag = existing ? existing.etag + 1 : 1;
    files[path] = { name: fileName, bytes: new Uint8Array(buffer.slice(0)), modified: new Date().toISOString(), etag };
    return Promise.resolve({ serverRelativeUrl: path, eTag: `"${etag}"` });
  }

  public deleteFile(_siteUrl: string, path: string): Promise<void> {
    delete files[norm(path)];
    return Promise.resolve();
  }

  public folderExists(_siteUrl: string, folder: string): Promise<boolean> {
    return Promise.resolve(Boolean(folders[norm(folder)]));
  }

  public ensureFolder(_siteUrl: string, folder: string): Promise<void> {
    ensure(folder);
    return Promise.resolve();
  }

  public getFolderFiles(_siteUrl: string, folder: string): Promise<ISharePointFileInfo[]> {
    const base = norm(folder);
    return Promise.resolve(
      Object.keys(files)
        .filter((p) => parentOf(p) === base)
        .map((p) => ({ name: files[p].name, serverRelativeUrl: p, timeLastModified: files[p].modified, length: files[p].bytes.length, eTag: `"${files[p].etag}"` }))
    );
  }

  public getFolderSubfolderNames(_siteUrl: string, folder: string): Promise<string[]> {
    const base = norm(folder);
    return Promise.resolve(Object.keys(folders).filter((p) => parentOf(p) === base).map(nameOf));
  }
}
