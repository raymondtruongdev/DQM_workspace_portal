/** Test đọc phân quyền báo cáo từ list Report_Role. Chạy: npm run test:report */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const R = require(path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report', 'utils', 'reportRole'));
const { RoleResolutionService } = require(path.join(root, 'temp', 'dev', 'js', 'src', 'core', 'services', 'RoleResolutionService'));

const names = { displayName: 'Title', email: 'Email', unit: 'Unit', docId: 'DocID', docTitle: 'DocTitle', active: 'Active' };
const row = (o) => ({ Id: 1, Title: 'Nguyễn A', Email: 'a@thadico.com', Unit: 'ST01', DocID: 'DailyReport', DocTitle: 'Báo cáo ngày', Active: true, ...o });

test('dòng hợp lệ thành phân quyền; DisplayName (cột Title) thành tên trưởng ga', () => {
  const [e] = R.parseRoleRows([row({})], names);
  assert.deepEqual(e, { id: '1', docId: 'DailyReport', docTitle: 'Báo cáo ngày', unit: 'ST01', siteManager: 'Nguyễn A', email: 'a@thadico.com', status: true });
});

test('một ô Email nhiều địa chỉ -> mỗi email một phân quyền', () => {
  const out = R.parseRoleRows([row({ Email: 'a@thadico.com; b@thadico.com,\nc@thadico.com' })], names);
  assert.deepEqual(out.map((e) => e.email), ['a@thadico.com', 'b@thadico.com', 'c@thadico.com']);
  assert.deepEqual(out.map((e) => e.id), ['1_1', '1_2', '1_3']);
});

test('chuẩn hoá DocID và đọc cột Active (Yes/No, chữ, để trống)', () => {
  assert.equal(R.normalizeDocId('drp'), 'DailyReport');
  assert.equal(R.normalizeDocId('Daily Report'), 'DailyReport');
  assert.equal(R.normalizeDocId('wrp'), 'WeeklyReport');
  assert.equal(R.normalizeDocId('Khác'), 'Khác');
  assert.equal(R.parseActive(true), true);
  assert.equal(R.parseActive(false), false);
  assert.equal(R.parseActive('TRUE'), true);
  assert.equal(R.parseActive('No'), false);
  assert.equal(R.parseActive(''), true);
  assert.equal(R.parseActive(undefined), true);
});

test('bỏ qua dòng thiếu Unit, DocID hoặc Email', () => {
  const out = R.parseRoleRows([row({ Unit: '' }), row({ DocID: '' }), row({ Email: '' }), row({ Id: 9 })], names);
  assert.equal(out.length, 1);
  assert.equal(out[0].id, '9');
});

test('người thường chỉ thấy phân quyền đang hiệu lực của mình; admin thấy tất cả đang hiệu lực', () => {
  const all = R.parseRoleRows([
    row({ Id: 1, Email: 'a@thadico.com', Unit: 'ST01' }),
    row({ Id: 2, Email: 'a@thadico.com', Unit: 'ST02', Active: false }),
    row({ Id: 3, Email: 'b@thadico.com', Unit: 'ST04' })
  ], names);
  const user = RoleResolutionService.resolveRoles(all, 'A@Thadico.com', false);
  assert.deepEqual(user.assignedRoles.map((e) => e.unit), ['ST01']);
  const admin = RoleResolutionService.resolveRoles(all, 'x@thadico.com', true);
  assert.deepEqual(admin.assignedRoles.map((e) => e.unit), ['ST01', 'ST04']);
  assert.equal(admin.allRoles.length, 3);
});
