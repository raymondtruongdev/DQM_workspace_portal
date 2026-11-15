/** Test đọc danh mục Activity từ ActivityP6.xlsx (một sheet: ActivityID, Unit, Description). Chạy: npm run test:report */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');
const XLSX = require('xlsx');

const root = path.resolve(__dirname, '..', '..');
const C = require(path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report', 'utils', 'reportCatalog'));

const book = (aoa, name = 'ActivityP6') => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
  return wb;
};
const ids = (cat) => cat.tables['activityp6_id'].items.map((i) => i.id);

test('file mẫu: lọc đúng Activity của từng đơn vị, kèm mô tả', () => {
  const wb = XLSX.readFile(path.join(root, 'docs', 'sharepoint_setup', 'DailyReportLibrary', 'ActivityP6.xlsx'));
  const st01 = C.parseCatalogFromWorkbook(XLSX, wb, 'ST01');
  assert.deepEqual(ids(st01), ['CV-ST01-00001', 'CV-ST01-00002', 'CV-ST01-00003']);
  assert.match(st01.tables['activityp6_id'].items[0].description, /tường vây/);
  assert.deepEqual(st01.warnings, []);
  assert.deepEqual(ids(C.parseCatalogFromWorkbook(XLSX, wb, 'st02')), ['CV-ST02-00001', 'CV-ST02-00002', 'CV-ST02-00003']);
});

test('Unit = ALL dùng chung mọi đơn vị; Description không bắt buộc; thứ tự cột tuỳ ý', () => {
  const wb = book([['Đơn vị', 'Activity ID'], ['ST01', 'A1'], ['ALL', 'A2'], ['ST09', 'A3']]);
  const cat = C.parseCatalogFromWorkbook(XLSX, wb, 'ST01');
  assert.deepEqual(ids(cat), ['A1', 'A2']);
  assert.equal(cat.tables['activityp6_id'].items[0].description, undefined);
});

test('cảnh báo: trùng mã, đơn vị không có Activity, thiếu cột, bỏ qua dòng trống và "---"', () => {
  const dup = C.parseCatalogFromWorkbook(XLSX, book([['ActivityID', 'Unit'], ['A1', 'ST01'], ['a1', 'ST01'], ['---', 'ST01'], ['', 'ST01']]), 'ST01');
  assert.deepEqual(ids(dup), ['A1']);
  assert.equal(dup.warnings.length, 1);
  assert.match(dup.warnings[0], /Trùng mã/);
  const none = C.parseCatalogFromWorkbook(XLSX, book([['ActivityID', 'Unit'], ['A1', 'ST01']]), 'ST07');
  assert.deepEqual(ids(none), []);
  assert.match(none.warnings[0], /Chưa có Activity nào của đơn vị ST07/);
  const bad = C.parseCatalogFromWorkbook(XLSX, book([['Mã', 'Tên']]), 'ST01');
  assert.match(bad.warnings[0], /cần có cột ActivityID và Unit/);
});

test('nhãn dropdown "mã – mô tả" chỉ cho mục có mô tả', () => {
  const wb = book([['ActivityID', 'Unit', 'Description'], ['A1', 'ST01', 'Đào đất'], ['A2', 'ST01', '']]);
  const cat = C.parseCatalogFromWorkbook(XLSX, wb, 'ST01');
  assert.deepEqual(C.catalogOptionLabels(cat, 'ActivityP6_ID'), { A1: 'A1 – Đào đất' });
  assert.equal(C.catalogOptionLabels(cat, 'element_id'), undefined);
  // gộp thêm mục từ list giữ mô tả
  const merged = C.mergeCatalogRows(cat, [{ table: 'activityp6_id', code: 'A9', link: '', description: 'Mới' }]);
  assert.equal(merged.tables['activityp6_id'].items.length, 3);
  assert.equal(C.catalogOptionLabels(merged, 'activityp6_id').A9, 'A9 – Mới');
});

test('chọn theo tên: nhãn "Tên Việt (mã) — ghi chú" cho mục có NameVN/NameEN/Note', () => {
  const cat = C.mergeCatalogRows({ unit: 'ST01', tables: {}, warnings: [] }, [
    { table: 'resequip_id', code: 'E1', link: 'Xe nâng người', nameVn: 'Xe nâng người', nameEn: 'Aerial work platform', note: 'Cao tối đa 12 m' },
    { table: 'resequip_id', code: 'E2', link: 'Máy khoan', note: 'Thuê ngoài' },
    { table: 'resequip_id', code: 'E3', link: 'Máy ủi' }
  ]);
  assert.deepEqual(C.catalogOptionLabels(cat, 'resequip_id', true), {
    'Xe nâng người': 'Xe nâng người (E1) — Cao tối đa 12 m',
    'Máy khoan': 'Máy khoan (E2) — Thuê ngoài'
  });
});

test('chọn theo mã: nhãn "Code — Note", không có Note thì chỉ hiện mã', () => {
  const cat = C.mergeCatalogRows({ unit: 'ST01', tables: {}, warnings: [] }, [
    { table: 'event_type', code: 'KIEM_TRA_ATLD', link: '', nameVn: 'Kiểm tra ATLĐ hiện trường', nameEn: 'HSE inspection', note: 'Kiểm tra ATLĐ' },
    { table: 'event_type', code: 'NGHIEM_THU', link: '', nameVn: 'Nghiệm thu', nameEn: 'Inspection request' }
  ]);
  assert.deepEqual(C.catalogOptionLabels(cat, 'event_type'), { KIEM_TRA_ATLD: 'KIEM_TRA_ATLD — Kiểm tra ATLĐ' });
  assert.deepEqual(C.catalogOptions(cat, 'event_type', false, true), ['KIEM_TRA_ATLD', 'NGHIEM_THU']);
});

test('fromNote: cột Loại tự điền theo Note của mã (không dùng Link)', () => {
  const cat = C.mergeCatalogRows({ unit: 'ST01', tables: {}, warnings: [] }, [
    { table: 'reshuman_id', code: 'SEG', link: 'Tên cũ', note: 'Bảo vệ' },
    { table: 'reshuman_id', code: 'XXX', link: '' }
  ]);
  const columns = [
    { code: 'resId', label: 'Mã', type: 'select', optionsRef: 'reshuman_id' },
    { code: 'category', label: 'Loại', type: 'text', from: 'resId', fromNote: true }
  ];
  const values = { resId: 'SEG', category: '' };
  C.applyCatalogLinks(columns, values, 'resId', cat);
  assert.equal(values.category, 'Bảo vệ');
  assert.equal(C.linkedValue(columns, columns[1], { resId: 'SEG' }, cat), 'Bảo vệ');
  const v2 = { resId: 'XXX', category: 'x' };
  C.applyCatalogLinks(columns, v2, 'resId', cat);
  assert.equal(v2.category, '');
  assert.equal(C.linkedValue(columns, columns[1], { resId: 'XXX' }, cat), undefined);
});
