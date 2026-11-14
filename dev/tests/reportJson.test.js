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
const sampleText = read('docs/sample/DailyReport.sample.data.json');
const tpl = () => J.parseTemplateJson(read('data_test/sharepoint_sample/_template/template_DRP_ST01_20261114.json'));
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
  const t = J.parseTemplateJson(read('data_test/sharepoint_sample/_template/template_DRP_ST02_20261114.json'));
  const doc = J.docFromTemplate(t, { docId: 'DailyReport', unit: 'ST02', dateStr: '2026-11-15', email: 'a@b.c', siteManager: 'Sếp', version: 'template_DRP_ST02_20261114' });
  assert.equal(doc.sections.length, 11);
  assert.equal(doc.metadata.reportNo, '00001');
  assert.match(doc.metadata.docCode, /-ST02-CWS-DRP-00001$/);
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
