#!/usr/bin/env node
/**
 * Tự động cài đặt SharePoint cho Workspace Portal (module Report).
 *
 * Làm gì (idempotent: chạy lại nhiều lần an toàn, chỉ bổ sung phần còn thiếu):
 *   1. Tạo 4 list: UserProfiles, ReportRole, ReportRegistry, CatalogItems (cột, index, unique, versioning).
 *   2. Tạo thư viện DailyReportLibrary + thư mục _template, resource, data, data/<Unit>.
 *   3. Upload ActivityP6.xlsx, template JSON, logo (mặc định KHÔNG ghi đè file đã có).
 *   4. Nạp dữ liệu: CatalogItems (mẫu), admin đầu tiên vào UserProfiles, (tuỳ chọn) ReportRole mẫu.
 *
 * KHÔNG làm (phải làm tay trên SharePoint, xem README.md): phân quyền list/thư viện, build + deploy .sppkg.
 *
 * Chạy:  node docs/sharepoint_setup/setup-sharepoint.mjs --admin you@company.com [--dry-run]
 * Xem README.md cùng thư mục để biết cách đăng nhập và các tuỳ chọn.
 * Không cần cài thêm gói npm (Node >= 22, dùng fetch có sẵn; đọc xlsx bằng gói `xlsx` của repo).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..', '..');
const require = createRequire(import.meta.url);

// ───────────────────────── tham số ─────────────────────────
function parseArgs(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith('--')) { out._.push(a); continue; }
    const key = a.slice(2);
    const next = argv[i + 1];
    if (next === undefined || next.startsWith('--')) out[key] = true;
    else { out[key] = next; i++; }
  }
  return out;
}

const args = parseArgs(process.argv.slice(2));
if (args.help || args.h) {
  console.log(`Tuỳ chọn:
  --config <file>      File cấu hình (mặc định setup.config.json cạnh script)
  --site <url>         Ghi đè siteUrl trong config
  --client-id <guid>   Client ID app Entra (public client) để đăng nhập device code
  --tenant <id|domain> Mặc định "organizations"
  --token <jwt>        Dùng sẵn access token SharePoint (hoặc biến môi trường SP_TOKEN) thay cho đăng nhập
  --admin <email>      Email admin đầu tiên ghi vào UserProfiles (bắt buộc khi tạo mới UserProfiles)
  --seed-sample        Nạp thêm dữ liệu mẫu cho ReportRole (CatalogItems luôn nạp mẫu; --no-seed-catalog để tắt)
  --no-seed-catalog    Không nạp mẫu CatalogItems
  --skip-lists         Bỏ qua phần list
  --skip-files         Bỏ qua phần thư viện + upload file
  --force              Ghi đè file đã có trong thư viện (mặc định bỏ qua file đã có)
  --dry-run            Chỉ in kế hoạch, không gọi SharePoint`);
  process.exit(0);
}

const configPath = path.resolve(args.config || path.join(HERE, 'setup.config.json'));
const cfg = JSON.parse(fs.readFileSync(configPath, 'utf8'));
const DRY = !!args['dry-run'];
const SITE = String(args.site || cfg.siteUrl || '').replace(/\/+$/, '');
if (!/^https:\/\/[^/]+\/(sites|teams)\/[^/]+/i.test(SITE) && !/^https:\/\/[^/]+$/i.test(SITE)) {
  fail(`siteUrl không hợp lệ: "${SITE}"`);
}
const SITE_PATH = new URL(SITE).pathname.replace(/\/+$/, ''); // "/sites/metro-checkin-dev"
const HOST = new URL(SITE).origin;

function fail(msg) { console.error(`✖ ${msg}`); process.exit(1); }
const log = (m) => console.log(m);
const ok = (m) => console.log(`  ✔ ${m}`);
const skip = (m) => console.log(`  • ${m}`);
const plan = (m) => console.log(`  ⋯ [dry-run] ${m}`);
const warn = (m) => console.warn(`  ⚠ ${m}`);

// ───────────────────────── đăng nhập ─────────────────────────
async function getToken() {
  const direct = args.token || process.env.SP_TOKEN;
  if (direct) return String(direct);
  const clientId = args['client-id'] || cfg.clientId;
  if (!clientId) fail('Thiếu client ID. Truyền --client-id <guid> (xem README mục 2) hoặc --token / SP_TOKEN.');
  const tenant = args.tenant || cfg.tenant || 'organizations';
  const scope = `${HOST}/AllSites.FullControl offline_access`;
  const base = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0`;
  const dc = await (await fetch(`${base}/devicecode`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: clientId, scope })
  })).json();
  if (!dc.device_code) fail(`Không lấy được device code: ${JSON.stringify(dc)}`);
  log(`\n${dc.message}\n`);
  const deadline = Date.now() + (dc.expires_in || 900) * 1000;
  let interval = (dc.interval || 5) * 1000;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, interval));
    const t = await (await fetch(`${base}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: clientId,
        device_code: dc.device_code
      })
    })).json();
    if (t.access_token) return t.access_token;
    if (t.error === 'authorization_pending') continue;
    if (t.error === 'slow_down') { interval += 5000; continue; }
    fail(`Đăng nhập thất bại: ${t.error_description || t.error}`);
  }
  return fail('Hết thời gian đăng nhập.');
}

// ───────────────────────── REST helper ─────────────────────────
let TOKEN = '';
const JSON_V = 'application/json;odata=verbose';

async function sp(method, urlPath, body, extra = {}) {
  const url = urlPath.startsWith('http') ? urlPath : `${SITE}/_api/${urlPath}`;
  const headers = {
    Authorization: `Bearer ${TOKEN}`,
    Accept: JSON_V,
    ...extra.headers
  };
  let payload;
  if (body instanceof Uint8Array || Buffer.isBuffer(body)) {
    payload = body;
    headers['Content-Type'] = 'application/octet-stream';
  } else if (body !== undefined) {
    payload = JSON.stringify(body);
    headers['Content-Type'] = `${JSON_V};charset=utf-8`;
  }
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, { method, headers, body: payload });
    if ((res.status === 429 || res.status === 503) && attempt < 4) {
      const wait = Number(res.headers.get('Retry-After')) || 2 ** attempt * 2;
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    if (res.status === 404 && extra.allow404) return null;
    const text = await res.text();
    if (!res.ok) {
      let detail = text;
      try { detail = JSON.parse(text).error.message.value; } catch { /* giữ nguyên */ }
      throw new Error(`${method} ${url.replace(SITE, '')} → ${res.status}: ${String(detail).slice(0, 400)}`);
    }
    return text ? JSON.parse(text) : {};
  }
}
const q = (s) => String(s).replace(/'/g, "''"); // escape trong OData literal
const enc = (s) => encodeURIComponent(s).replace(/'/g, '%27');
const merge = (urlPath, body) => sp('POST', urlPath, body, { headers: { 'X-HTTP-Method': 'MERGE', 'If-Match': '*' } });

// ───────────────────────── đọc xlsx mẫu ─────────────────────────
function readSheet(file, sheet) {
  let XLSX;
  try { XLSX = require('xlsx'); } catch { fail('Thiếu gói xlsx: chạy `npm install` ở gốc repo trước.'); }
  const wb = XLSX.readFile(path.resolve(ROOT, file));
  const ws = wb.Sheets[sheet] || wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

// ───────────────────────── list ─────────────────────────
function fieldXml(col) {
  const n = col.name;
  const attrs = `DisplayName="${n}" Name="${n}" StaticName="${n}"`;
  switch (col.type) {
    case 'Number': return `<Field Type="Number" ${attrs} Decimals="${col.decimals ?? 0}" />`;
    case 'Boolean': return `<Field Type="Boolean" ${attrs}><Default>${col.default === false ? 0 : 1}</Default></Field>`;
    default: return `<Field Type="Text" ${attrs} MaxLength="255" />`;
  }
}

async function ensureList(def, baseTemplate = 100) {
  const key = `lists/getbytitle('${q(def.name)}')`;
  if (DRY) { plan(`đảm bảo ${baseTemplate === 101 ? 'thư viện' : 'list'} "${def.name}"`); return { created: true, key }; }
  const existing = await sp('GET', `${key}?$select=Id`, undefined, { allow404: true });
  if (existing) { skip(`"${def.name}" đã tồn tại`); return { created: false, key }; }
  await sp('POST', 'lists', {
    __metadata: { type: 'SP.List' },
    BaseTemplate: baseTemplate,
    Title: def.name,
    Description: def.description || '',
    AllowContentTypes: true
  });
  ok(`đã tạo "${def.name}"`);
  return { created: true, key };
}

async function setupList(def) {
  log(`\n▶ List ${def.name}`);
  const { key } = await ensureList(def);
  const wantVersioning = def.versioning !== false;

  // 1) cột còn thiếu
  let have = new Set();
  if (!DRY) {
    const r = await sp('GET', `${key}/fields?$select=Title,InternalName&$filter=Hidden eq false`);
    for (const f of r.d.results) { have.add(f.Title.toLowerCase()); have.add(f.InternalName.toLowerCase()); }
  }
  for (const col of def.columns) {
    if (have.has(col.name.toLowerCase())) { skip(`cột ${col.name} đã có`); continue; }
    if (DRY) { plan(`thêm cột ${col.name} (${col.type})`); continue; }
    await sp('POST', `${key}/fields/createfieldasxmlinput`, {
      parameters: {
        __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
        SchemaXml: fieldXml(col),
        Options: 8 // AddFieldInternalNameHint: tên nội bộ = tên cột, không bị "field_N"
      }
    });
    ok(`thêm cột ${col.name} (${col.type})`);
  }

  // 2) Title: bắt buộc hay không, index + duy nhất
  const titleBody = { __metadata: { type: 'SP.Field' }, Required: def.titleRequired !== false };
  if (def.titleUnique) { titleBody.Indexed = true; titleBody.EnforceUniqueValues = true; }
  if (DRY) plan(`Title: Required=${titleBody.Required}${def.titleUnique ? ', Indexed + Enforce unique values' : ''}`);
  else { await merge(`${key}/fields/getbytitle('Title')`, titleBody); ok('cấu hình cột Title'); }

  // 3) index các cột khác
  for (const col of def.columns.filter((c) => c.indexed)) {
    if (DRY) { plan(`index cột ${col.name}`); continue; }
    try {
      await merge(`${key}/fields/getbyinternalnameortitle('${q(col.name)}')`, { __metadata: { type: 'SP.Field' }, Indexed: true });
      ok(`index cột ${col.name}`);
    } catch (e) { warn(`không index được ${col.name}: ${e.message}`); }
  }

  // 4) versioning
  if (wantVersioning) {
    if (DRY) plan('bật versioning');
    else { await merge(key, { __metadata: { type: 'SP.List' }, EnableVersioning: true }); ok('bật versioning'); }
  }

  await seedList(def, key);
}

async function seedList(def, key) {
  const seed = def.seed || { mode: 'none' };
  let rows = [];
  if (seed.mode === 'admin') {
    const email = args.admin;
    if (!email) {
      warn(`UserProfiles chưa có admin: truyền --admin <email> để ghi admin đầu tiên (không có admin thì không ai vào được Admin Config).`);
      return;
    }
    rows = [{ Title: email, Email: email, DisplayName: email.split('@')[0], Department: 'ADMIN', DepartmentName: 'Admin site', Role: 'ADMIN', Active: true }];
  } else if (seed.mode === 'sample') {
    const enabled = seed.default ? !args['no-seed-catalog'] : !!args['seed-sample'];
    if (!enabled) { skip('bỏ qua dữ liệu mẫu'); return; }
    rows = readSheet(seed.file, seed.sheet).filter((r) => Object.values(r).some((v) => String(v).trim() !== ''));
    if (def.name === 'ReportRole') rows = rows.map((r) => ({ ...r, Title: r.DisplayName || r.Email }));
  } else return;

  if (DRY) { plan(`nạp ${rows.length} dòng vào ${def.name}`); return; }

  // khoá nhận diện dòng đã có: Title (CatalogItems) hoặc Email (UserProfiles) hoặc cả bộ (ReportRole)
  const idOf = (r) => def.name === 'UserProfiles' ? String(r.Email).toLowerCase()
    : def.name === 'ReportRole' ? [r.Email, r.Unit, r.DocID].join('|').toLowerCase()
      : String(r.Title).toLowerCase();
  const exist = new Set();
  const selectCols = def.name === 'UserProfiles' ? 'Title,Email' : def.name === 'ReportRole' ? 'Title,Email,Unit,DocID' : 'Title';
  let next = `${SITE}/_api/${key}/items?$select=${selectCols}&$top=5000`;
  while (next) {
    const r = await sp('GET', next);
    for (const it of r.d.results) exist.add(idOf(it));
    next = r.d.__next || '';
  }
  const meta = await sp('GET', `${key}?$select=ListItemEntityTypeFullName`);
  const type = meta.d.ListItemEntityTypeFullName;
  const cols = new Set(['Title', ...def.columns.map((c) => c.name)]);
  let added = 0, existed = 0;
  for (const r of rows) {
    if (exist.has(idOf(r))) { existed++; continue; }
    const body = { __metadata: { type } };
    for (const k of Object.keys(r)) if (cols.has(k) && r[k] !== '') body[k] = r[k];
    await sp('POST', `${key}/items`, body);
    added++;
  }
  ok(`dữ liệu ${def.name}: thêm ${added}, đã có sẵn ${existed}`);
}

// ───────────────────────── thư viện + file ─────────────────────────
function globFiles(pattern) {
  const abs = path.resolve(ROOT, pattern);
  const dir = path.dirname(abs);
  const re = new RegExp('^' + path.basename(abs).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*') + '$');
  return fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => re.test(f)).map((f) => path.join(dir, f)) : [];
}

async function ensureFolder(rel) {
  const full = `${SITE_PATH}/${cfg.library.name}/${rel}`;
  if (DRY) { plan(`thư mục ${cfg.library.name}/${rel}`); return; }
  await sp('POST', `web/folders/add('${q(full)}')`); // trả về thư mục sẵn có nếu đã tồn tại
  ok(`thư mục ${cfg.library.name}/${rel}`);
}

async function setupLibrary() {
  const lib = cfg.library;
  log(`\n▶ Thư viện ${lib.name}`);
  const { key } = await ensureList({ name: lib.name, description: lib.description }, 101);
  if (lib.versioning !== false) {
    if (DRY) plan('bật versioning');
    else { await merge(key, { __metadata: { type: 'SP.List' }, EnableVersioning: true }); ok('bật versioning'); }
  }
  for (const f of lib.folders) await ensureFolder(f);

  // danh sách file cần upload
  const uploads = [];
  for (const f of lib.files) {
    if (f.from) uploads.push({ local: path.resolve(ROOT, f.from), folder: path.posix.dirname(f.to) === '.' ? '' : path.posix.dirname(f.to), name: path.posix.basename(f.to) });
    else for (const local of globFiles(f.fromGlob)) uploads.push({ local, folder: f.toFolder, name: path.basename(local) });
  }

  if (lib.unitFoldersUnderData) {
    const units = new Set();
    for (const u of uploads) { const m = /^template_DRP_([^_]+)_\d{8}\.json$/i.exec(u.name); if (m) units.add(m[1]); }
    for (const unit of [...units].sort()) await ensureFolder(`data/${unit}`);
  }

  for (const u of uploads) {
    const label = `${u.folder ? u.folder + '/' : ''}${u.name}`;
    if (!fs.existsSync(u.local)) { warn(`không thấy file nguồn ${path.relative(ROOT, u.local)}`); continue; }
    if (DRY) { plan(`upload ${label}`); continue; }
    const folderUrl = `${SITE_PATH}/${lib.name}${u.folder ? '/' + u.folder : ''}`;
    const fileUrl = `${folderUrl}/${u.name}`;
    if (!args.force) {
      const ex = await sp('GET', `web/GetFileByServerRelativeUrl('${q(fileUrl)}')?$select=Name`, undefined, { allow404: true });
      if (ex) { skip(`${label} đã có (dùng --force để ghi đè)`); continue; }
    }
    await sp('POST', `web/GetFolderByServerRelativeUrl('${q(folderUrl)}')/Files/add(url='${q(u.name)}',overwrite=true)`, fs.readFileSync(u.local));
    ok(`upload ${label}`);
  }
}

// ───────────────────────── chạy ─────────────────────────
async function main() {
  log(`Site: ${SITE}${DRY ? '   (DRY-RUN: không gọi SharePoint)' : ''}`);
  if (!DRY) {
    TOKEN = await getToken();
    const web = await sp('GET', 'web?$select=Title,Url');
    ok(`đã đăng nhập, kết nối site "${web.d.Title}"`);
  }
  if (!args['skip-lists']) for (const def of cfg.lists) await setupList(def);
  if (!args['skip-files']) await setupLibrary();

  log(`\n${DRY ? 'Kế hoạch xong.' : 'Hoàn tất.'} Việc còn lại phải làm tay (xem README.md mục 5):
  • Phân quyền list/thư viện (người thường chỉ đọc UserProfiles, ReportRole, _template, resource; chỉ thêm CatalogItems...)
  • Build + deploy workspace_portal.sppkg, thêm web part lên trang
  • Chạy bảng kiểm tra mục 8 của docs/SHAREPOINT_SETUP.md`);
}

main().catch((e) => { console.error(`\n✖ ${e.message}`); process.exit(1); });
