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

// ---- Chế độ "chỉ nhập ngày hiện tại" (Setup DailyReport) ----
const S = require(path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report', 'utils', 'reportSettings'));

test('parseReportSettings: mặc định, ép kiểu và kẹp số ngày', () => {
  assert.deepEqual(S.parseReportSettings(undefined), { todayOnly: false, windowDays: 14, unlockHours: 3 });
  assert.deepEqual(S.parseReportSettings({ todayOnly: true, windowDays: 7 }), { todayOnly: true, windowDays: 7, unlockHours: 3 });
  assert.equal(S.parseReportSettings({ windowDays: 999 }).windowDays, 60);
  assert.equal(S.parseReportSettings({ windowDays: 0 }).windowDays, 1);
  assert.equal(S.parseReportSettings({ windowDays: 'abc', todayOnly: 'yes' }).windowDays, 14);
  assert.equal(S.parseReportSettings({ todayOnly: 'yes' }).todayOnly, false);
});

test('isLockedByTodayOnly: chỉ khoá ngày cũ khi bật', () => {
  const today = S.todayDateStr();
  assert.equal(S.isLockedByTodayOnly({ todayOnly: true, windowDays: 14, unlockHours: 3 }, today), false);
  assert.equal(S.isLockedByTodayOnly({ todayOnly: true, windowDays: 14, unlockHours: 3 }, '2000-01-01'), true);
  assert.equal(S.isLockedByTodayOnly({ todayOnly: false, windowDays: 14, unlockHours: 3 }, '2000-01-01'), false);
});

test('isRegistryRowLocked: ngày cũ khoá tự động nhưng admin mở khoá riêng (Locked = 0) vẫn được', () => {
  const R = require(path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report', 'services', 'ReportRegistryService'));
  assert.equal(R.isRegistryRowLocked(undefined, 'unsigned', true), true);
  assert.equal(R.isRegistryRowLocked(false, 'unsigned', true), false);
  assert.equal(R.isRegistryRowLocked(false, 'done', true), false);
  assert.equal(R.isRegistryRowLocked(true, 'unsigned', false), true);
  assert.equal(R.isRegistryRowLocked(undefined, 'unsigned', false), false);
  assert.equal(R.isRegistryRowLocked(undefined, 'done', false), true);
});

test('mở khoá có hạn: hết hạn thì quay về quy tắc mặc định', () => {
  const R = require(path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report', 'services', 'ReportRegistryService'));
  const now = Date.parse('2026-11-18T12:00:00Z');
  const future = '2026-11-18T13:00:00Z';
  const past = '2026-11-18T11:00:00Z';
  assert.equal(R.isRegistryRowLocked(false, 'done', false, future, now), false); // còn hạn: mở
  assert.equal(R.isRegistryRowLocked(false, 'done', false, past, now), true); // hết hạn, đã ký: khoá lại
  assert.equal(R.isRegistryRowLocked(false, 'unsigned', true, past, now), true); // hết hạn, ngày cũ: khoá lại
  assert.equal(R.isRegistryRowLocked(false, 'unsigned', false, past, now), false); // hết hạn, chưa ký, hôm nay: vẫn mở
  assert.equal(R.isRegistryRowLocked(false, 'done', false, '', now), false); // không có hạn: mở mãi
  assert.equal(R.isUnlockExpired('rác', now), false);
});

// ---- Giờ máy chủ + ngày theo giờ Việt Nam ----
const CLK = require(path.join(root, 'temp', 'dev', 'js', 'src', 'core', 'services', 'ServerClock'));

test('vnDateStr: ngày theo giờ Việt Nam (UTC+7), không phụ thuộc múi giờ máy', () => {
  assert.equal(S.vnDateStr(Date.parse('2026-11-17T16:59:59Z')), '2026-11-17');
  assert.equal(S.vnDateStr(Date.parse('2026-11-17T17:00:00Z')), '2026-11-18'); // 0h giờ VN
  assert.equal(S.todayDateStr(Date.parse('2026-11-18T01:00:00Z')), '2026-11-18');
});

test('ServerClock: dùng giờ máy chủ, không theo giờ máy người dùng; lỗi thì dùng giờ máy', async () => {
  const farServer = Date.now() + 10 * 86400000; // máy chủ chạy trước máy người dùng 10 ngày
  const clock = new CLK.ServerClock({ getServerTime: async () => farServer }, 'https://x');
  await clock.ensureFresh(0);
  assert.ok(Math.abs(clock.now() - farServer) < 2000);
  assert.equal(clock.isSynced, true);
  const broken = new CLK.ServerClock({ getServerTime: async () => { throw new Error('mạng'); } }, 'https://x');
  await broken.ensureFresh(0);
  assert.equal(broken.isSynced, false);
  assert.ok(Math.abs(broken.now() - Date.now()) < 2000);
  const none = new CLK.ServerClock({}, 'https://x'); // client không có getServerTime
  await none.ensureFresh(0);
  assert.equal(none.isSynced, false);
});

test('getLock dùng giờ máy chủ: ngày cũ (theo giờ VN) khoá, hết hạn mở khoá tính theo máy chủ', async () => {
  const c = fakeClient([row({ Locked: 0, UnlockUntil: '2026-11-18T03:00:00.000Z' })], BASE_COLS.concat('Locked', 'UnlockUntil'));
  const mk = (iso) => new R.ReportRegistryService(c, 'https://x', () => 'lib', async () => ({ todayOnly: true, windowDays: 14, unlockHours: 3 }), { now: () => Date.parse(iso), ensureFresh: async () => {} });
  // 02:00Z = 09:00 giờ VN ngày 18: ngày 14 là ngày cũ nhưng admin mở khoá còn hạn (đến 03:00Z)
  const open = await mk('2026-11-18T02:00:00Z').getLock('ST01', 'DRP', '2026-11-14');
  assert.equal(open.locked, false);
  assert.equal(open.unlockUntil, '2026-11-18T03:00:00.000Z');
  // 04:00Z: hết hạn, ngày cũ bị khoá lại
  const closed = await mk('2026-11-18T04:00:00Z').getLock('ST01', 'DRP', '2026-11-14');
  assert.equal(closed.locked, true);
  // ngày chưa có dòng: 16:59Z ngày 17 (23:59 VN) thì ngày 17 còn là hôm nay, 17:00Z thì đã là ngày cũ
  assert.equal((await mk('2026-11-17T16:59:00Z').getLock('ST01', 'DRP', '2026-11-17')).locked, false);
  assert.equal((await mk('2026-11-17T17:00:00Z').getLock('ST01', 'DRP', '2026-11-17')).locked, true);
});
