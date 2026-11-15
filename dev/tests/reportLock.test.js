/** Test khoá báo cáo (cột Locked của sổ DRPRegistry). Chạy: npm run test:report */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const R = require(path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report', 'services', 'ReportRegistryService'));

/** Client giả: một list DRPRegistry trong bộ nhớ; bộ lọc chỉ tìm theo khoá Title. */
function fakeClient(rows, columns) {
  const map = {};
  columns.forEach((c) => {
    map[c.toLowerCase()] = c;
  });
  return {
    rows,
    updates: [],
    getListFieldMap: async () => map,
    queryItems: async (_site, _list, query) => {
      const m = /Title eq '([^']+)'/.exec(query);
      return m ? rows.filter((r) => r.Title === m[1]) : rows.slice();
    },
    updateItem: async function (_site, _list, id, data) {
      this.updates.push({ id, data });
      Object.assign(rows.filter((r) => r.Id === id)[0], data);
    },
    createItem: async (_site, _list, data) => {
      rows.push({ Id: rows.length + 1, ...data });
    }
  };
}
const BASE_COLS = ['Title', 'Unit', 'DocID', 'ReportDate', 'ReportNo', 'Status', 'SavedAt', 'SignedAt'];
const row = (extra) => ({ Id: 1, Title: 'DRP_20261114_ST01', Unit: 'ST01', DocID: 'DRP', ReportDate: '2026-11-14', ReportNo: 1, Status: 'unsigned', ...extra });
const service = (client) => new R.ReportRegistryService(client, 'https://x', () => 'lib');

test('isRegistryRowLocked: cột Locked trống thì theo trạng thái (đã ký = khoá), có giá trị thì theo giá trị', () => {
  assert.equal(R.isRegistryRowLocked(undefined, 'done'), true);
  assert.equal(R.isRegistryRowLocked(undefined, 'unsigned'), false);
  assert.equal(R.isRegistryRowLocked(true, 'unsigned'), true); // admin khoá dù chưa ký
  assert.equal(R.isRegistryRowLocked(false, 'done'), false); // admin đã mở khoá
});

test('getLock: báo cáo đã ký từ trước (Locked trống) coi là khoá; chưa có dòng thì không khoá', async () => {
  const c = fakeClient([row({ Status: 'done' })], BASE_COLS.concat('Locked'));
  const s = service(c);
  assert.deepEqual(await s.getLock('ST01', 'DRP', '2026-11-14'), { exists: true, locked: true, canToggle: true });
  assert.deepEqual(await s.getLock('ST01', 'DRP', '2026-11-13'), { exists: false, locked: false, canToggle: true });
});

test('admin mở khoá rồi khoá lại: ghi Locked = 0 / 1 (Number)', async () => {
  const c = fakeClient([row({ Status: 'done' })], BASE_COLS.concat('Locked'));
  const s = service(c);
  await s.setLocked('ST01', 'DRP', '2026-11-14', false);
  assert.equal(c.updates[0].data.Locked, 0);
  assert.equal((await s.getLock('ST01', 'DRP', '2026-11-14')).locked, false);
  await s.setLocked('ST01', 'DRP', '2026-11-14', true);
  assert.equal(c.updates[1].data.Locked, 1);
  assert.equal((await s.getLock('ST01', 'DRP', '2026-11-14')).locked, true);
});

test('upload bản ký tự khoá lại; xoá bản ký mở khoá', async () => {
  const c = fakeClient([row({ Locked: 0 })], BASE_COLS.concat('Locked'));
  const s = service(c);
  await s.recordSigned('ST01', 'DRP', '2026-11-14', '00001', 'a@b.c', 'a@b.c');
  assert.equal(c.rows[0].Locked, 1);
  assert.equal((await s.getLock('ST01', 'DRP', '2026-11-14')).locked, true);
  await s.recordUnsigned('ST01', 'DRP', '2026-11-14', 'admin@b.c');
  assert.equal(c.rows[0].Locked, 0);
  assert.equal((await s.getLock('ST01', 'DRP', '2026-11-14')).locked, false);
});

test('lịch/Tổng quan: ReportDayInfo.locked lấy từ sổ', async () => {
  const c = fakeClient(
    [row({ Status: 'done', Locked: 0 }), row({ Id: 2, Title: 'DRP_20261113_ST01', ReportDate: '2026-11-13', Status: 'done' }), row({ Id: 3, Title: 'DRP_20261112_ST01', ReportDate: '2026-11-12', Locked: 1 })],
    BASE_COLS.concat('Locked')
  );
  const all = await service(c).getDayInfosSince('DRP', '2026-11-01');
  const infos = all[Object.keys(all)[0]];
  assert.equal(infos['2026-11-14'].locked, false); // admin đã mở khoá
  assert.equal(infos['2026-11-13'].locked, true); // đã ký, Locked trống
  assert.equal(infos['2026-11-12'].locked, true); // admin khoá dù chưa ký
});

test('list chưa có cột Locked: báo cáo đã ký vẫn khoá nhưng không bật/tắt được', async () => {
  const c = fakeClient([row({ Status: 'done' })], BASE_COLS);
  const s = service(c);
  assert.deepEqual(await s.getLock('ST01', 'DRP', '2026-11-14'), { exists: true, locked: true, canToggle: false });
  await assert.rejects(() => s.setLocked('ST01', 'DRP', '2026-11-14', false), /Locked/);
});

test('admin khoá trước ngày chưa nhập: tạo dòng giữ chỗ, vẫn tính "Chưa nhập" nhưng đã khoá', async () => {
  const c = fakeClient([], BASE_COLS.concat('Locked'));
  const s = service(c);
  await s.setLocked('ST01', 'DRP', '2026-11-14', true);
  assert.equal(c.rows.length, 1);
  assert.equal(c.rows[0].Locked, 1);
  assert.equal(c.rows[0].ReportNo, 0);
  assert.deepEqual(await s.getLock('ST01', 'DRP', '2026-11-14'), { exists: true, locked: true, canToggle: true });
  const all = await s.getDayInfosSince('DRP', '2026-11-01');
  const info = all[Object.keys(all)[0]]['2026-11-14'];
  assert.equal(info.status, 'missing');
  assert.equal(info.locked, true);
  assert.equal(info.savedAt, undefined);
});

test('mở khoá ngày giữ chỗ rồi lưu báo cáo: dòng được điền số báo cáo và thành "Chưa upload bản ký"', async () => {
  const c = fakeClient([], BASE_COLS.concat('Locked', 'Reporter'));
  const s = service(c);
  await s.setLocked('ST01', 'DRP', '2026-11-14', true);
  await s.setLocked('ST01', 'DRP', '2026-11-14', false);
  assert.equal(c.rows.length, 1); // không tạo thêm dòng
  assert.equal((await s.getLock('ST01', 'DRP', '2026-11-14')).locked, false);
  await s.recordSave('ST01', 'DRP', '2026-11-14', '00007', 'a@b.c');
  assert.equal(c.rows.length, 1);
  assert.equal(c.rows[0].ReportNo, 7);
  const all = await s.getDayInfosSince('DRP', '2026-11-01');
  assert.equal(all[Object.keys(all)[0]]['2026-11-14'].status, 'unsigned');
});

test('mở khoá ngày chưa có dòng thì không tạo dòng', async () => {
  const c = fakeClient([], BASE_COLS.concat('Locked'));
  await service(c).setLocked('ST01', 'DRP', '2026-11-14', false);
  assert.equal(c.rows.length, 0);
});

test('isPlaceholderRow: chỉ dòng chưa lưu lần nào (không SavedAt, số 0, chưa ký)', () => {
  assert.equal(R.isPlaceholderRow('', 0, 'unsigned'), true);
  assert.equal(R.isPlaceholderRow('2026-11-14T01:00:00Z', 1, 'unsigned'), false);
  assert.equal(R.isPlaceholderRow('', 0, 'done'), false);
});

test('ReportLockedError có thông điệp hướng dẫn liên hệ admin', () => {
  const e = new R.ReportLockedError();
  assert.ok(e instanceof Error);
  assert.match(e.message, /admin/);
});
