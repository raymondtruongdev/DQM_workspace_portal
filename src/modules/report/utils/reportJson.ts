import {
  CellValue,
  FormField,
  PhotoGroup,
  ReportDoc,
  ReportForm,
  ReportMetadata,
  ReportPhotoSection,
  ReportSection,
  ReportTable,
  TableColumn,
  TableRow
} from '../types/report.types';
import {
  JsonCellValue,
  LabelJson,
  REPORT_JSON_SCHEMA,
  ReportJson,
  ReportJsonColumn,
  ReportJsonField,
  ReportJsonPhotoRow,
  ReportJsonRow,
  ReportJsonSection,
  ReportTemplateJson,
  TemplateColumnJson,
  TemplateFieldJson,
  TemplateSectionJson
} from '../types/reportJson.types';
import { applyFormDefaults } from './reportForm';
import { formatReportNo, syncMetadata } from './reportMeta';
import { carryForwardTable, columnTotal, defaultCellValue, recalcTable, splitLabel } from './reportTable';

/**
 * Chuyển đổi giữa ba dạng:
 *   - file báo cáo `.data.json` (ReportJson): nhãn + giá trị thật, gọn, đọc được một mình;
 *   - template `.json` (ReportTemplateJson): cấu trúc + cách nhập (kiểu ô, danh mục, công thức, mặc định…);
 *   - mô hình đang sửa trên web (ReportDoc): file báo cáo ghép với template theo `id` mục / `id` cột / `id` trường.
 * Cột/mục/trường có trong file báo cáo mà template không có (hoặc chưa tải được template) vẫn hiển thị như ô văn bản thường.
 */

/* ------------------------------- Nhãn ------------------------------- */

function labelText(l: LabelJson): string {
  return l.en ? `${l.vi} | ${l.en}` : l.vi;
}

function labelJson(text: string): LabelJson {
  const { vi, en } = splitLabel(text);
  return en ? { vi, en } : { vi };
}

/* ------------------------- Template -> mô hình ------------------------- */

/** Cột của template thành cột của mô hình (chưa có nhãn của báo cáo). */
export function templateColumn(tc: TemplateColumnJson): TableColumn {
  const col: TableColumn = {
    code: tc.id,
    label: labelText(tc.label),
    type: tc.type || (tc.catalog || tc.options ? 'select' : 'text')
  };
  if (tc.catalog) {
    col.optionsRef = tc.catalog.toLowerCase();
    if (tc.catalogBy === 'link') col.optionsBy = 'link';
  }
  if (tc.options) col.options = tc.options.slice();
  if (tc.noNone) col.noNone = true;
  if (tc.from) col.from = tc.from;
  if (tc.formula) col.formula = tc.formula;
  if (tc.readonly) col.readonly = true;
  if (tc.key) col.key = true;
  if (tc.total) col.total = true;
  if (tc.align) col.align = tc.align;
  if (tc.width) col.width = tc.width;
  if (tc.default !== undefined) col.defaultValue = String(tc.default);
  if (tc.reset) col.reset = true;
  if (tc.carry) col.carry = tc.carry;
  if (tc.tones) col.tones = { ...tc.tones };
  return col;
}

function templateField(tf: TemplateFieldJson, value: CellValue): FormField {
  const f: FormField = { key: tf.id, label: labelText(tf.label), type: tf.type || (tf.options ? 'select' : 'text'), value };
  if (tf.options) f.options = tf.options.slice();
  if (tf.full) f.full = true;
  if (tf.span) f.span = tf.span;
  if (tf.rows) f.rows = tf.rows;
  if (tf.lines) f.lines = tf.lines;
  if (tf.placeholder) f.placeholder = tf.placeholder;
  if (tf.default !== undefined) f.defaultValue = String(tf.default);
  return f;
}

function jsonRowToRow(r: ReportJsonRow): TableRow {
  return { id: r.id, isLocked: r.locked ? true : undefined, values: { ...r.cells } };
}

const sectionTitle = (t: { vi: string; en?: string; short?: string }): { title: string; titleEn: string; shortTitle: string } => ({
  title: t.vi,
  titleEn: t.en || '',
  shortTitle: t.short || ''
});

/** Mục mới tạo từ template (dòng mặc định + trường mặc định). */
export function templateSection(ts: TemplateSectionJson): ReportSection {
  const head = { key: ts.id, ...sectionTitle(ts.title), summary: '', showSummary: ts.showSummary !== false };
  if (ts.kind === 'photo') {
    return { kind: 'photo', ...head, columns: ts.perRow || 2, groups: [] };
  }
  if (ts.columns && ts.columns.length > 0) {
    const table: ReportTable = {
      kind: 'table',
      ...head,
      single: ts.single === true,
      columns: ts.columns.map(templateColumn),
      rows: (ts.rows || []).map(jsonRowToRow)
    };
    if (ts.fields && ts.fields.length > 0) {
      table.fields = ts.fields.map((tf) => templateField(tf, ''));
      table.fieldColumns = ts.layout ? ts.layout.columns : 1;
    }
    return recalcTable(table);
  }
  const form: ReportForm = {
    kind: 'form',
    ...head,
    columns: ts.layout ? ts.layout.columns : 1,
    sign: Boolean(ts.layout && ts.layout.sign),
    fields: (ts.fields || []).map((tf) => templateField(tf, ''))
  };
  return form;
}

export interface NewDocContext {
  docId: string;
  unit: string;
  dateStr: string;
  email: string;
  siteManager?: string;
  /** Tên template (vd template_DRP_ST01_20261114). */
  version: string;
}

/** Báo cáo mới hoàn toàn từ template (đơn vị chưa có báo cáo nào). */
export function docFromTemplate(tpl: ReportTemplateJson, ctx: NewDocContext): ReportDoc {
  const metadata: ReportMetadata = syncMetadata(
    {
      docTitle: tpl.title.vi,
      docTitleEn: tpl.title.en,
      subTitle: tpl.title.sub,
      docCode: '',
      docCodePattern: tpl.docCodePattern,
      templateVersion: ctx.version,
      location: ctx.unit,
      reportNo: formatReportNo(1),
      subcontractors: [],
      email: ctx.email,
      updateTime: new Date().toISOString()
    },
    ctx.unit
  );
  return applyFormDefaults(
    {
      docId: ctx.docId,
      unit: ctx.unit,
      dateStr: ctx.dateStr,
      metadata,
      sections: tpl.sections.map(templateSection),
      tasks: [],
      templateUpdateRequired: false
    },
    { unit: ctx.unit, email: ctx.email, siteManager: ctx.siteManager }
  );
}

/* ------------------------- File báo cáo -> mô hình ------------------------- */

/** Cột chỉ có số hoặc trống thì coi là cột số (khi không có template). */
function inferredType(id: string, rows: ReportJsonRow[]): 'number' | 'text' {
  const vals = rows.map((r) => r.cells[id]).filter((v) => v !== undefined && v !== '');
  return vals.length > 0 && vals.every((v) => typeof v === 'number') ? 'number' : 'text';
}

function jsonColumn(jc: ReportJsonColumn, tc: TemplateColumnJson | undefined, rows: ReportJsonRow[]): TableColumn {
  const base: TableColumn = tc ? templateColumn(tc) : { code: jc.id, label: '', type: inferredType(jc.id, rows) };
  return { ...base, code: jc.id, label: labelText(jc.label) };
}

function jsonField(jf: ReportJsonField, tf: TemplateFieldJson | undefined): FormField {
  const base: FormField = tf
    ? templateField(tf, jf.value)
    : { key: jf.id, label: '', type: typeof jf.value === 'number' ? 'number' : 'text', value: jf.value };
  return { ...base, key: jf.id, label: labelText(jf.label), value: jf.value };
}

function isPhotoRow(r: ReportJsonRow | ReportJsonPhotoRow): r is ReportJsonPhotoRow {
  return Array.isArray((r as ReportJsonPhotoRow).files);
}

function jsonSection(js: ReportJsonSection, ts: TemplateSectionJson | undefined): ReportSection {
  const head = {
    key: js.id,
    ...sectionTitle(js.title),
    summary: js.summary || '',
    // Có template: theo template. Chưa có: chỉ hiện ô Tóm tắt khi đã có nội dung
    showSummary: ts ? ts.showSummary !== false : (js.summary || '') !== ''
  };
  const jsonRows = (js.rows || []).filter((r): r is ReportJsonRow => !isPhotoRow(r));
  if (js.type === 'photo') {
    const groups: PhotoGroup[] = (js.rows || []).filter(isPhotoRow).map((r) => ({
      id: r.id,
      department: r.department,
      summary: r.summary || '',
      files: r.files.slice()
    }));
    const photo: ReportPhotoSection = { kind: 'photo', ...head, columns: (ts && ts.perRow) || 2, groups };
    return photo;
  }
  const fields = (js.fields || []).map((jf) => jsonField(jf, ts && ts.fields ? ts.fields.filter((f) => f.id === jf.id)[0] : undefined));
  if (js.columns && js.columns.length > 0) {
    const table: ReportTable = {
      kind: 'table',
      ...head,
      single: Boolean(ts && ts.single),
      columns: js.columns.map((jc) => jsonColumn(jc, ts && ts.columns ? ts.columns.filter((c) => c.id === jc.id)[0] : undefined, jsonRows)),
      rows: jsonRows.map(jsonRowToRow)
    };
    if (fields.length > 0) {
      table.fields = fields;
      table.fieldColumns = ts && ts.layout ? ts.layout.columns : 1;
    }
    // Ô công thức luôn tính lại, không tin giá trị đã lưu
    return recalcTable(table);
  }
  const form: ReportForm = {
    kind: 'form',
    ...head,
    columns: ts && ts.layout ? ts.layout.columns : 1,
    sign: Boolean(ts && ts.layout && ts.layout.sign),
    fields
  };
  return form;
}

/** Ghép file báo cáo với template (có thể chưa có) thành mô hình đang sửa. */
export function jsonToDoc(json: ReportJson, tpl?: ReportTemplateJson): ReportDoc {
  const metadata: ReportMetadata = syncMetadata(
    {
      docTitle: json.title.vi,
      docTitleEn: json.title.en,
      subTitle: json.title.sub,
      docCode: json.docCode,
      docCodePattern: tpl ? tpl.docCodePattern : undefined,
      templateVersion: json.template,
      location: json.unit,
      reportNo: formatReportNo(json.reportNo),
      subcontractors: (json.subcontractors || []).slice(),
      email: json.updated.by,
      updateTime: json.updated.at,
      created: { ...json.created }
    },
    json.unit
  );
  // Mã tài liệu đã lưu là giá trị phát hành: giữ nguyên, không tính lại theo mẫu của template mới
  metadata.docCode = json.docCode;
  return {
    docId: json.docId,
    unit: json.unit,
    dateStr: json.date,
    metadata,
    sections: json.sections.map((js) => jsonSection(js, tpl ? tpl.sections.filter((s) => s.id === js.id)[0] : undefined)),
    tasks: (json.tasks || []).map((t) => ({ ...t }))
  };
}

/* ------------------------- Mô hình -> file báo cáo ------------------------- */

function cellOf(v: CellValue): JsonCellValue | undefined {
  return v === undefined ? undefined : v;
}

function fieldToJson(f: FormField): ReportJsonField {
  return { id: f.key, label: labelJson(f.label), value: f.value === undefined ? '' : f.value };
}

function sectionToJson(s: ReportSection): ReportJsonSection {
  const title = { vi: s.title, ...(s.titleEn ? { en: s.titleEn } : {}), ...(s.shortTitle ? { short: s.shortTitle } : {}) };
  if (s.kind === 'photo') {
    return {
      id: s.key,
      title,
      type: 'photo',
      summary: s.summary || '',
      rows: s.groups.map((g) => ({ id: g.id, department: g.department, summary: g.summary || '', files: g.files.slice() }))
    };
  }
  if (s.kind === 'form') {
    return { id: s.key, title, summary: s.summary || '', fields: s.fields.map(fieldToJson) };
  }
  const out: ReportJsonSection = {
    id: s.key,
    title,
    summary: s.summary || '',
    columns: s.columns.map((c) => ({ id: c.code, label: labelJson(c.label) })),
    rows: s.rows.map((r) => {
      const cells: Record<string, JsonCellValue> = {};
      s.columns.forEach((c) => {
        const v = cellOf(r.values[c.code]);
        if (v !== undefined) cells[c.code] = v;
      });
      return r.isLocked ? { id: r.id, locked: true, cells } : { id: r.id, cells };
    })
  };
  const totalCols = s.columns.filter((c) => c.total);
  if (totalCols.length > 0) {
    const totals: Record<string, number> = {};
    totalCols.forEach((c) => {
      totals[c.code] = Math.round(columnTotal(s, c) * 100) / 100;
    });
    out.totals = totals;
  }
  if (s.fields && s.fields.length > 0) out.fields = s.fields.map(fieldToJson);
  return out;
}

/** Mô hình đang sửa thành nội dung file `.data.json` (gọn, chỉ nhãn + giá trị thật). */
export function docToJson(doc: ReportDoc): ReportJson {
  const m = doc.metadata;
  const reportNo = parseInt(m.reportNo, 10);
  const json: ReportJson = {
    schema: REPORT_JSON_SCHEMA,
    unit: doc.unit,
    docId: doc.docId,
    date: doc.dateStr,
    reportNo: isNaN(reportNo) ? 1 : reportNo,
    docCode: m.docCode,
    title: { vi: m.docTitle, en: m.docTitleEn, sub: m.subTitle },
    template: m.templateVersion || '',
    created: m.created ? { ...m.created } : { by: m.email, at: m.updateTime },
    updated: { by: m.email, at: m.updateTime },
    subcontractors: m.subcontractors.slice(),
    sections: doc.sections.map(sectionToJson)
  };
  if (doc.tasks && doc.tasks.length > 0) json.tasks = doc.tasks.map((t) => ({ ...t }));
  return json;
}

/** Nội dung file `.data.json` dạng văn bản (xuống dòng + thụt lề để người mở file đọc được). */
export function serializeReportJson(json: ReportJson): string {
  return `${JSON.stringify(json, null, 2)}\n`;
}

/* ------------------------------ Đọc / kiểm tra ------------------------------ */

function fail(what: string, why: string): never {
  throw new Error(`${what} không hợp lệ: ${why}`);
}

/** Đọc file báo cáo; sai cấu trúc cơ bản thì ném lỗi dễ hiểu. */
export function parseReportJson(text: string): ReportJson {
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    return fail('File báo cáo', 'không phải JSON');
  }
  const j = raw as Partial<ReportJson>;
  if (!j || typeof j !== 'object') return fail('File báo cáo', 'nội dung rỗng');
  if (typeof j.schema !== 'number' || j.schema > REPORT_JSON_SCHEMA) {
    return fail('File báo cáo', `phiên bản định dạng (schema) ${String(j.schema)} chưa được hỗ trợ`);
  }
  if (typeof j.unit !== 'string' || typeof j.docId !== 'string' || typeof j.date !== 'string') {
    return fail('File báo cáo', 'thiếu unit / docId / date');
  }
  if (!Array.isArray(j.sections)) return fail('File báo cáo', 'thiếu danh sách mục (sections)');
  const t = (j.title || {}) as Partial<ReportJson['title']>;
  const stamp = (s: ReportJson['created'] | undefined): ReportJson['created'] => ({ by: (s && s.by) || '', at: (s && s.at) || '' });
  return {
    ...(j as ReportJson),
    reportNo: Number(j.reportNo) || 1,
    docCode: j.docCode || '',
    title: { vi: t.vi || '', en: t.en || '', sub: t.sub || '' },
    template: j.template || '',
    created: stamp(j.created),
    updated: stamp(j.updated),
    subcontractors: Array.isArray(j.subcontractors) ? j.subcontractors : []
  };
}

/** Đọc template JSON; sai cấu trúc cơ bản thì ném lỗi dễ hiểu. */
export function parseTemplateJson(text: string): ReportTemplateJson {
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    return fail('Template', 'không phải JSON');
  }
  const t = raw as Partial<ReportTemplateJson>;
  if (!t || typeof t !== 'object' || !Array.isArray(t.sections)) return fail('Template', 'thiếu danh sách mục (sections)');
  if (typeof t.schema !== 'number' || t.schema > REPORT_JSON_SCHEMA) return fail('Template', 'phiên bản định dạng (schema) chưa được hỗ trợ');
  const title = (t.title || {}) as Partial<ReportTemplateJson['title']>;
  return {
    ...(t as ReportTemplateJson),
    docId: t.docId || 'DailyReport',
    docCodePattern: t.docCodePattern || '',
    title: { vi: title.vi || '', en: title.en || '', sub: title.sub || '' }
  };
}

/* ------------------------------ Apply template ------------------------------ */

function mergeFields(existing: FormField[] | undefined, defs: TemplateFieldJson[] | undefined): FormField[] {
  return (defs || []).map((tf) => {
    const cur = existing ? existing.filter((f) => f.key === tf.id)[0] : undefined;
    return templateField(tf, cur ? cur.value : '');
  });
}

function mergeSection(cur: ReportSection, ts: TemplateSectionJson): ReportSection {
  const fresh = templateSection(ts);
  if (cur.kind !== fresh.kind) return fresh;
  const keep = { summary: cur.summary };
  if (fresh.kind === 'photo' && cur.kind === 'photo') return { ...fresh, ...keep, groups: cur.groups };
  if (fresh.kind === 'form' && cur.kind === 'form') return { ...fresh, ...keep, fields: mergeFields(cur.fields, ts.fields) };
  if (fresh.kind === 'table' && cur.kind === 'table') {
    const templateRowIds = (ts.rows || []).filter((r) => r.locked).map((r) => r.id);
    const rows: TableRow[] = cur.rows.map((r) => {
      const values: Record<string, CellValue> = {};
      fresh.columns.forEach((c) => {
        values[c.code] = r.values[c.code] !== undefined ? r.values[c.code] : defaultCellValue(c);
      });
      // Dòng từng là dòng mẫu nhưng template mới không còn thì mở khoá để người dùng xoá được
      return { ...r, isLocked: r.isLocked && templateRowIds.indexOf(r.id) === -1 ? undefined : r.isLocked, values };
    });
    // Dòng mẫu mới của template mà báo cáo chưa có
    fresh.rows.forEach((fr) => {
      if (fr.isLocked && !rows.some((r) => r.id === fr.id)) rows.push(fr);
    });
    return recalcTable({ ...fresh, ...keep, rows, fields: fresh.fields ? mergeFields(cur.fields, ts.fields) : undefined });
  }
  return fresh;
}

/**
 * Áp dụng template mới nhất lên báo cáo đang có, ghép theo `id`: giữ dữ liệu cũ, thêm mục/cột/trường mới (ô trống hoặc mặc định),
 * bỏ mục/cột đã bị xoá khỏi template, cập nhật nhãn và cách nhập theo template.
 */
export function applyTemplate(doc: ReportDoc, tpl: ReportTemplateJson, version: string, ctx: { email: string; siteManager?: string }): ReportDoc {
  const sections = tpl.sections.map((ts) => {
    const cur = doc.sections.filter((s) => s.key === ts.id)[0];
    return cur ? mergeSection(cur, ts) : templateSection(ts);
  });
  const metadata = syncMetadata(
    {
      ...doc.metadata,
      docTitle: tpl.title.vi,
      docTitleEn: tpl.title.en,
      subTitle: tpl.title.sub,
      docCodePattern: tpl.docCodePattern,
      templateVersion: version
    },
    doc.unit
  );
  return applyFormDefaults(
    { ...doc, metadata, sections, templateUpdateRequired: false },
    { unit: doc.unit, email: ctx.email, siteManager: ctx.siteManager }
  );
}

/* --------------------------- Tạo từ báo cáo ngày trước --------------------------- */

/**
 * Báo cáo của `dateStr` khởi tạo từ báo cáo ngày trước: luỹ kế (cột carry) và reset (cột reset) theo template,
 * ảnh không sao chép; các trường văn bản giữ nguyên. Số báo cáo/mã tài liệu do nơi gọi cấp lại.
 */
export function carryForwardDoc(prev: ReportDoc, dateStr: string, email: string): ReportDoc {
  const now = new Date().toISOString();
  return {
    ...prev,
    dateStr,
    isExisting: false,
    eTag: undefined,
    copiedFromDate: prev.dateStr,
    sections: prev.sections.map((s) =>
      s.kind === 'photo' ? { ...s, groups: [] } : s.kind === 'table' ? carryForwardTable(s) : s
    ),
    metadata: { ...prev.metadata, email, updateTime: now, created: undefined }
  };
}
