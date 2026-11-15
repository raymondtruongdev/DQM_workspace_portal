/**
 * Test chuyển đổi file báo cáo JSON <-> mô hình <-> template. Chạy: node dev/build.js && node --test dev/tests
 * (dùng bản biên dịch trong temp/dev/js nên không đụng tới dist/ của `heft start`).
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const js = path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report', 'utils');
const J = require(path.join(js, 'reportJson'));
const T = require(path.join(js, 'reportTable'));

const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');
// Báo cáo mẫu đã lưu = template ST01 (cùng định dạng) + phần đầu của một báo cáo thật (ngày, người lập, đơn vị thi công…)
const sampleText = (() => {
  const tplFile = JSON.parse(
    read('docs/sharepoint_setup/DailyReportLibrary/_template/template_DRP_ST01_20261114.json')
      .replace(/\{user\}/g, 'user@example.com')
      .replace(/\{siteManager\}/g, 'Trưởng ga cũ')
  );
  tplFile.sections.forEach((sec) => (sec.rows || []).forEach((r) => { if (r.locked === false) delete r.locked; }));
  tplFile.sections.find((sec) => sec.type === 'photo').rows = [
    { id: 'g1', department: 'DCEM', summary: 'Hố đào tấm P-12', files: ['img1.jpeg', 'img2.jpeg', 'img3.jpeg'] },
    { id: 'g2', department: 'HSE', summary: 'Kiểm tra an toàn hiện trường', files: ['img4.jpeg', 'img5.jpeg', 'img6.jpeg'] }
  ];
  return JSON.stringify({
    ...tplFile,
    date: '2026-11-10',
    reportNo: 1,
    template: 'template_DRP_ST01_20261114',
    created: { by: 'user@example.com', at: '2026-11-10T01:00:00.000Z' },
    updated: { by: 'user@example.com', at: '2026-11-14T02:47:46.958Z' },
    subcontractors: ['THAC_CRSG', 'THAC_THAN']
  });
})();
const tpl = () => J.parseTemplateJson(read('docs/sharepoint_setup/DailyReportLibrary/_template/template_DRP_ST01_20261114.json'));
const clone = (o) => JSON.parse(JSON.stringify(o));

test('mở rồi lưu lại file mẫu không mất/đổi dữ liệu (round-trip)', () => {
  const json = J.parseReportJson(sampleText);
  const back = J.docToJson(J.jsonToDoc(json, tpl()));
  assert.deepEqual(back, json);
});

test('round-trip khi chưa có template: mọi cột vẫn giữ nguyên giá trị', () => {
  const json = J.parseReportJson(sampleText);
  const back = J.docToJson(J.jsonToDoc(json, undefined));
  assert.deepEqual(back.sections.map((s) => s.id), json.sections.map((s) => s.id));
  assert.deepEqual(back.sections.find((s) => s.id === 'events').rows, json.sections.find((s) => s.id === 'events').rows);
  assert.deepEqual(back.sections.find((s) => s.id === 'hse').fields, json.sections.find((s) => s.id === 'hse').fields);
});

test('ô công thức luôn tính lại, không tin giá trị đã lưu', () => {
  const json = J.parseReportJson(sampleText);
  const wq = json.sections.find((s) => s.id === 'workQty');
  wq.rows[0].cells.progress = 999; // giá trị bị sửa tay
  const mp = json.sections.find((s) => s.id === 'manpower');
  mp.totals = { daily: 1, prevCum: 1, cumTotal: 1 };
  const back = J.docToJson(J.jsonToDoc(json, tpl()));
  assert.equal(back.sections.find((s) => s.id === 'workQty').rows[0].cells.progress, 60);
  assert.deepEqual(back.sections.find((s) => s.id === 'manpower').totals, { daily: 114, prevCum: 100, cumTotal: 214 });
});

test('mã tài liệu đã lưu được giữ nguyên dù template đổi mẫu', () => {
  const json = J.parseReportJson(sampleText);
  const t = tpl();
  t.docCodePattern = 'XXX-{unit}-{reportNo}';
  const doc = J.jsonToDoc(json, t);
  assert.equal(doc.metadata.docCode, json.docCode);
});

test('báo cáo mới từ template: đủ mục, số báo cáo 00001, mặc định {user}', () => {
  const t = J.parseTemplateJson(read('docs/sharepoint_setup/DailyReportLibrary/_template/template_DRP_ST01_20261114.json'));
  const doc = J.docFromTemplate(t, { docId: 'DailyReport', unit: 'ST01', dateStr: '2026-11-15', email: 'a@b.c', siteManager: 'Sếp', version: 'template_DRP_ST01_20261114' });
  assert.equal(doc.sections.length, 11);
  assert.equal(doc.metadata.reportNo, '00001');
  assert.match(doc.metadata.docCode, /-ST01-CWS-DRP-00001$/);
  const sign = doc.sections.find((s) => s.key === 'signOff');
  assert.equal(sign.kind, 'form');
  assert.equal(sign.fields.find((f) => f.key === 'preparedBy').value, 'a@b.c');
  assert.equal(sign.fields.find((f) => f.key === 'siteManager').value, 'Sếp');
});

test('tạo từ báo cáo ngày trước: luỹ kế cộng dồn, ô hôm nay về 0, bỏ ảnh', () => {
  const doc = J.jsonToDoc(J.parseReportJson(sampleText), tpl());
  const next = J.carryForwardDoc(doc, '2026-11-11', 'x@y.z');
  assert.equal(next.dateStr, '2026-11-11');
  assert.equal(next.copiedFromDate, '2026-11-10');
  const wq = next.sections.find((s) => s.key === 'workQty');
  assert.equal(wq.rows[0].values.daily, 0);
  assert.equal(wq.rows[0].values.prevCum, 600);
  assert.equal(wq.rows[0].values.progress, 60);
  assert.equal(next.sections.find((s) => s.key === 'photos').groups.length, 0);
  assert.equal(next.metadata.created, undefined);
  const w = next.sections.find((s) => s.key === 'weather');
  assert.equal(w.rows[0].values.am, 'Nắng'); // reset về mặc định
});

test('Apply template mới: thêm cột/mục mới, giữ dữ liệu cũ, bỏ cột đã xoá', () => {
  const doc = J.jsonToDoc(J.parseReportJson(sampleText), tpl());
  const t = tpl();
  const events = t.sections.find((s) => s.id === 'events');
  events.columns.push({ id: 'owner', label: { vi: 'Phụ trách', en: 'Owner' }, default: 'Site' });
  events.columns = events.columns.filter((c) => c.id !== 'remark');
  t.sections.push({ id: 'extra', title: { vi: 'MỤC MỚI', en: 'NEW' }, fields: [{ id: 'note', label: { vi: 'Ghi chú' }, default: 'x' }], layout: { columns: 1 } });
  const next = J.applyTemplate(doc, t, 'template_DRP_ST01_29991231', { email: 'a@b.c' });
  const ev = next.sections.find((s) => s.key === 'events');
  assert.ok(ev.columns.some((c) => c.code === 'owner'));
  assert.ok(!ev.columns.some((c) => c.code === 'remark'));
  assert.equal(ev.rows.length, 4);
  assert.equal(ev.rows[0].values.owner, 'Site');
  assert.equal(ev.rows[0].values.eventId, 'EVT-ST01-00001'); // dữ liệu cũ còn nguyên
  assert.equal(next.sections[next.sections.length - 1].key, 'extra');
  assert.equal(next.metadata.templateVersion, 'template_DRP_ST01_29991231');
  assert.equal(next.sections.find((s) => s.key === 'photos').groups.length, 2);
});

test('template chỉ chứa dữ liệu: không có thuộc tính UI, UI do code (reportUiSchema) mở rộng khi đọc', () => {
  const file = JSON.parse(read('docs/sharepoint_setup/DailyReportLibrary/_template/template_DRP_ST01_20261114.json'));
  const uiKeys = ['type', 'catalog', 'catalogBy', 'from', 'options', 'noNone', 'formula', 'readonly', 'key', 'total', 'align', 'width', 'default', 'reset', 'carry', 'tones',
    'kind', 'single', 'showSummary', 'layout', 'perRow', 'full', 'span', 'lines', 'placeholder'];
  const bad = [];
  // Template cùng định dạng file báo cáo: cùng bộ khoá cấp cao nhất và cấp mục
  assert.deepEqual(Object.keys(file), Object.keys(JSON.parse(sampleText)));
  file.sections.forEach((s) => {
    uiKeys.filter((k) => k !== 'type').forEach((k) => { if (k in s) bad.push(`${s.id}.${k}`); });
    (s.columns || []).concat(s.fields || []).forEach((c) => uiKeys.forEach((k) => { if (k in c) bad.push(`${s.id}.${c.id}.${k}`); }));
  });
  assert.deepEqual(bad, []);
  const t = tpl();
  const wq = t.sections.find((s) => s.id === 'workQty');
  assert.equal(wq.columns.find((c) => c.id === 'progress').formula, 'round((daily+prevCum)/planned*100)');
  assert.equal(wq.columns.find((c) => c.id === 'activityId').catalog, 'ActivityP6_ID');
  assert.equal(t.sections.find((s) => s.id === 'photos').kind, 'photo');
  assert.equal(t.docCodePattern, file.docCodeTemplate); // mẫu mã tài liệu có {unit}/{reportNo}
  assert.match(t.docCodePattern, /\{unit\}.*\{reportNo\}/);
  assert.equal(t.sections.find((s) => s.id === 'signOff').fields.find((f) => f.id === 'preparedBy').auto, '{user}');
  // Mục/cột không khai báo UI trong code vẫn dùng được như văn bản thường
  const extra = J.parseTemplateJson(JSON.stringify({ schema: 1, docId: 'DailyReport', title: {}, docCodePattern: '', sections: [{ id: 'zz', title: { vi: 'Z' }, columns: [{ id: 'a', label: { vi: 'A' } }] }] }));
  assert.equal(extra.sections[0].columns[0].type, undefined);
});

test('báo cáo đã lưu dùng thẳng làm template: mã tài liệu tính lại theo đơn vị đang mở, người lập/trưởng ga theo người dùng hiện tại', () => {
  const t = J.parseTemplateJson(sampleText); // sampleText = báo cáo đã lưu (docCodeTemplate + docCode thật, preparedBy/siteManager là tên người cũ)
  assert.equal(t.docCodePattern, 'HCM2_1-23_CS-THAC_DCEM-CPZZ-{unit}-CWS-DRP-{reportNo}');
  const doc = J.docFromTemplate(t, { docId: 'DailyReport', unit: 'ST02', dateStr: '2026-11-15', email: 'moi@b.c', siteManager: 'Sếp mới', version: 'x' });
  assert.equal(doc.metadata.docCode, 'HCM2_1-23_CS-THAC_DCEM-CPZZ-ST02-CWS-DRP-00001');
  // File lưu ghi lại docCodeTemplate để tự dùng làm template được
  assert.equal(J.docToJson(doc).docCodeTemplate, t.docCodePattern);
  const sign = doc.sections.find((s) => s.key === 'signOff');
  assert.equal(sign.fields.find((f) => f.key === 'preparedBy').value, 'moi@b.c');
  assert.equal(sign.fields.find((f) => f.key === 'siteManager').value, 'Sếp mới');
  // Kế thừa từ báo cáo ngày trước cũng điền lại người lập (không giữ tên người cũ); mở lại báo cáo đã lưu thì giữ giá trị đã lưu
  const prev = J.jsonToDoc(J.parseReportJson(sampleText), tpl());
  assert.equal(prev.sections.find((s) => s.key === 'signOff').fields.find((f) => f.key === 'preparedBy').value, 'user@example.com');
  const next = J.carryForwardDoc(prev, '2026-11-11', 'moi@b.c');
  assert.equal(next.sections.find((s) => s.key === 'signOff').fields.find((f) => f.key === 'preparedBy').value, '');
});

test('không có template riêng của đơn vị thì dùng template_DRP_default.json; template riêng luôn ưu tiên', async () => {
  const { ReportFileService } = require(path.join(js, '..', 'services', 'ReportFileService'));
  const N = require(path.join(js, 'reportFileNaming'));
  assert.equal(N.matchDefaultTemplateFileName('template_DRP_default.json', 'DailyReport').version, 'template_DRP_default');
  assert.equal(N.matchDefaultTemplateFileName('template_DRP_default_20261120.json', 'DailyReport').versionDate, '20261120');
  assert.equal(N.matchDefaultTemplateFileName('template_WRP_default.json', 'DailyReport'), undefined);
  assert.equal(N.matchDefaultTemplateFileName('template_DRP_ST01_20261114.json', 'DailyReport'), undefined);
  const make = (names) => new ReportFileService({
    getFolderFiles: async () => names.map((n) => ({ name: n, serverRelativeUrl: '/' + n, timeLastModified: '2026-11-15T00:00:00Z' })),
    getFileBuffer: async (_site, url) => Buffer.from(url.slice(1))
  }, 'https://x.sharepoint.com/sites/s');
  const pick = async (names, unit) => { const r = await make(names).findLatestTemplateForUnit(unit, 'DailyReport'); return r.info && r.info.fileName; };
  assert.equal(await pick(['template_DRP_ST01_20261114.json', 'template_DRP_default.json'], 'ST02'), 'template_DRP_default.json');
  assert.equal(await pick(['template_DRP_ST01_20261114.json', 'template_DRP_default_20271231.json'], 'ST01'), 'template_DRP_ST01_20261114.json'); // riêng thắng dù chung mới hơn
  assert.equal(await pick(['template_DRP_ST01_20261114.json'], 'ST02'), undefined);
});

test('template_DRP_default.json: sạch dữ liệu riêng của ST01, tạo báo cáo cho đơn vị khác dùng được', () => {
  const text = read('docs/sharepoint_setup/DailyReportLibrary/_template/template_DRP_default.json');
  const t = J.parseTemplateJson(text);
  const doc = J.docFromTemplate(t, { docId: 'DailyReport', unit: 'ST02', dateStr: '2026-11-15', email: 'a@b.c', siteManager: 'Sếp', version: 'template_DRP_default' });
  assert.equal(doc.metadata.docCode, 'HCM2_1-23_CS-THAC_DCEM-CPZZ-ST02-CWS-DRP-00001');
  const rows = doc.sections.filter((s) => s.kind === 'table').reduce((a, s) => a.concat(s.rows), []);
  assert.ok(!rows.some((r) => r.isLocked), 'không có dòng khoá');
  assert.ok(!rows.some((r) => Object.values(r.values).some((v) => /ST01/.test(String(v)))), 'không còn mã riêng ST01');
  assert.equal(doc.sections.find((s) => s.key === 'weather').rows[0].values.am, 'Nắng');
  assert.equal(doc.sections.find((s) => s.key === 'manpower').rows.length, 4);
  assert.equal(doc.sections.find((s) => s.key === 'signOff').fields.find((f) => f.key === 'preparedBy').value, 'a@b.c');
  assert.equal(doc.sections.find((s) => s.key === 'hse').fields.find((f) => f.key === 'engineerInstructions').value, '');
});

test('cột Luỹ kế hôm qua chỉ đọc ở mọi dòng (Khối lượng và Nhân lực) nhưng vẫn cộng dồn khi kế thừa', () => {
  const t = tpl();
  ['workQty', 'manpower'].forEach((id) => {
    const col = t.sections.find((s) => s.id === id).columns.find((c) => c.id === 'prevCum');
    assert.equal(col.readonly, true, `${id}.prevCum readonly`);
    assert.equal(col.carry, 'daily+prevCum');
  });
  const doc = J.jsonToDoc(J.parseReportJson(sampleText), t);
  const mp = doc.sections.find((s) => s.key === 'manpower');
  assert.equal(mp.columns.find((c) => c.code === 'prevCum').readonly, true);
  assert.equal(J.carryForwardDoc(doc, '2026-11-11', 'x@y.z').sections.find((s) => s.key === 'manpower').rows[0].values.prevCum, 108);
});

test('mã dòng mới = số lớn nhất + 1, không trùng', () => {
  const doc = J.jsonToDoc(J.parseReportJson(sampleText), tpl());
  const ev = doc.sections.find((s) => s.key === 'events');
  ev.rows = ev.rows.filter((r) => r.id !== 'r2'); // xoá r2: r1, r3, r4
  assert.equal(T.nextRowId(ev), 'r5');
});

test('file hỏng hoặc sai phiên bản thì báo lỗi dễ hiểu', () => {
  assert.throws(() => J.parseReportJson('{không phải json'), /không phải JSON/);
  assert.throws(() => J.parseReportJson('{"schema":1,"unit":"ST01"}'), /thiếu/);
  assert.throws(() => J.parseReportJson(JSON.stringify({ ...clone(JSON.parse(sampleText)), schema: 99 })), /chưa được hỗ trợ/);
  assert.throws(() => J.parseTemplateJson('{"schema":1}'), /thiếu danh sách mục/);
});
