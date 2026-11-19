/** Test bộ đếm mã tự tăng (list DRPCounters): nhập báo cáo lệch ngày không trùng mã, xung đột ghi (412) được gộp. Chạy: npm run test:report */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const base = path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report');
const { ReportCounterService, mergeCounters, parseCounters } = require(path.join(base, 'services', 'ReportCounterService'));
const T = require(path.join(base, 'utils', 'reportTable'));

/** List giả có ETag: `beforeUpdate(row)` cho phép giả lập người khác ghi chen vào giữa lúc đọc và lúc ghi. */
function fakeClient(opts = {}) {
  const rows = [];
  const calls = { create: 0, update: 0, conflict412: 0 };
  return {
    rows,
    calls,
    getListFieldMap: async () => ({ title: 'Title', counters: 'Counters0' }),
    queryItemsWithETag: async (_s, _l, query) => {
      const m = /Title eq '([^']*)'/.exec(query);
      return rows.filter((r) => r.Title === m[1].replace(/''/g, "'")).map((r) => ({ item: { ...r }, eTag: `"${r.v}"` }));
    },
    createItem: async (_s, _l, data) => {
      calls.create++;
      if (opts.beforeCreate) opts.beforeCreate(rows, data);
      if (rows.some((r) => r.Title === data.Title)) throw new Error('Duplicate value (409) unique');
      rows.push({ Id: rows.length + 1, v: 1, ...data });
    },
    updateItem: async (_s, _l, id, data, eTag) => {
      calls.update++;
      const row = rows.find((r) => r.Id === id);
      if (opts.beforeUpdate) opts.beforeUpdate(row);
      if (eTag !== `"${row.v}"`) {
        calls.conflict412++;
        throw new Error('Failed to update item (412): Precondition Failed');
      }
      Object.assign(row, data);
      row.v++;
    }
  };
}

const read = (client, unit = 'ST01') => JSON.parse(client.rows.find((r) => r.Title === `${unit}|DRP`).Counters0);

test('parseCounters / mergeCounters: lấy số lớn hơn, bỏ giá trị hỏng', () => {
  assert.deepEqual(parseCounters('{"a":5,"b":"x","c":-1,"d":2.9}'), { a: 5, d: 2 });
  assert.deepEqual(parseCounters('không phải json'), {});
  assert.deepEqual(mergeCounters({ a: 5, b: 1 }, { a: 3, c: 7 }), { a: 5, b: 1, c: 7 });
});

test('nhập báo cáo KHÔNG theo thứ tự ngày vẫn không trùng mã (ngày 10 → 12 → nhập bù 11 → 13)', async () => {
  const c = fakeClient();
  const svc = new ReportCounterService(c, 'https://x');
  const issued = [];
  const issue = async (n) => {
    const floor = await svc.get('ST01', 'DRP');
    const start = floor['events.eventid'] || 0;
    for (let i = 1; i <= n; i++) issued.push(start + i);
    await svc.raise('ST01', 'DRP', { 'events.eventid': start + n }); // = bộ đếm ghi khi Lưu
  };
  await issue(5); // ngày 10: 1..5
  await issue(3); // ngày 12: 6..8
  await issue(2); // ngày 11 (nhập bù): 9..10
  await issue(2); // ngày 13: phải là 11..12, KHÔNG phải 9..10
  assert.deepEqual(issued, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
  assert.equal(new Set(issued).size, issued.length);
});

test('raise: không bao giờ giảm bộ đếm (lưu báo cáo cũ có số nhỏ hơn)', async () => {
  const c = fakeClient();
  const svc = new ReportCounterService(c, 'https://x');
  await svc.raise('ST01', 'DRP', { 'events.eventid': 10, 'issues.issueid': 4 });
  await svc.raise('ST01', 'DRP', { 'events.eventid': 6 });
  assert.deepEqual(read(c), { 'events.eventid': 10, 'issues.issueid': 4 });
  assert.equal(c.calls.update, 0); // không có gì nâng thì không ghi
});

test('raise: mỗi đơn vị một dòng riêng', async () => {
  const c = fakeClient();
  const svc = new ReportCounterService(c, 'https://x');
  await svc.raise('ST01', 'DRP', { 'events.eventid': 3 });
  await svc.raise('ST02', 'DRP', { 'events.eventid': 9 });
  assert.deepEqual(await svc.get('ST01', 'DRP'), { 'events.eventid': 3 });
  assert.deepEqual(await svc.get('ST02', 'DRP'), { 'events.eventid': 9 });
  assert.deepEqual(await svc.get('ST03', 'DRP'), {});
});

test('raise: người khác ghi chen giữa (412) -> đọc lại, gộp max, không mất số của họ', async () => {
  let injected = false;
  const c = fakeClient({
    beforeUpdate: (row) => {
      if (injected) return;
      injected = true;
      row.Counters0 = JSON.stringify({ 'events.eventid': 20, 'issues.issueid': 5 }); // người khác vừa lưu
      row.v++;
    }
  });
  const svc = new ReportCounterService(c, 'https://x');
  await svc.raise('ST01', 'DRP', { 'events.eventid': 2 }); // tạo dòng
  await svc.raise('ST01', 'DRP', { 'events.eventid': 8, 'issues.issueid': 1 });
  assert.equal(c.calls.conflict412, 1);
  assert.deepEqual(read(c), { 'events.eventid': 20, 'issues.issueid': 5 });
  await svc.raise('ST01', 'DRP', { 'events.eventid': 25 });
  assert.equal(read(c)['events.eventid'], 25);
});

test('raise: hai người cùng tạo dòng đầu tiên (trùng khoá) -> người sau đọc lại và cập nhật', async () => {
  let injected = false;
  const c = fakeClient({
    beforeCreate: (rows) => {
      if (injected) return;
      injected = true;
      rows.push({ Id: 1, v: 1, Title: 'ST01|DRP', Counters0: JSON.stringify({ 'events.eventid': 7 }) });
    }
  });
  const svc = new ReportCounterService(c, 'https://x');
  await svc.raise('ST01', 'DRP', { 'events.eventid': 3, 'issues.issueid': 2 });
  assert.deepEqual(read(c), { 'events.eventid': 7, 'issues.issueid': 2 });
});

test('raise: lỗi khác (quyền, mạng) báo ngay, không thử lặp', async () => {
  const c = fakeClient();
  c.createItem = async () => {
    throw new Error('Failed to create item (403): Access denied');
  };
  const svc = new ReportCounterService(c, 'https://x');
  await assert.rejects(svc.raise('ST01', 'DRP', { 'events.eventid': 1 }), /403/);
});

test('get: list chưa tạo -> ném lỗi (nơi gọi cảnh báo)', async () => {
  const c = fakeClient();
  c.getListFieldMap = async () => {
    throw new Error('SharePoint API Error (404): list does not exist');
  };
  await assert.rejects(new ReportCounterService(c, 'https://x').get('ST01', 'DRP'), /404/);
});

test('mã mới của dòng thêm tay dùng bộ đếm đơn vị làm nền (createRow/fillAutoIds)', () => {
  const table = {
    kind: 'table',
    key: 'events',
    columns: [{ code: 'eventId', type: 'text', defaultValue: 'EVT-{unit}-{n5}' }],
    rows: []
  };
  const row = T.createRow(table, 'ST01', { 'events.eventid': 10 });
  assert.equal(row.values.eventId, 'EVT-ST01-00011');
});

test('reserve: mỗi lần đặt chỗ ra một số mới, bộ đếm tăng ngay trong list (xoá hàng không làm giảm)', async () => {
  const c = fakeClient();
  const svc = new ReportCounterService(c, 'https://x');
  assert.equal(await svc.reserve('ST01', 'DRP', 'events.eventid', 0), 1);
  assert.equal(await svc.reserve('ST01', 'DRP', 'events.eventid', 0), 2); // hàng 1 bị xoá chưa lưu: số 1 vẫn không cấp lại
  assert.equal(read(c)['events.eventid'], 2);
  assert.equal(await svc.reserve('ST01', 'DRP', 'issues.issueid', 0), 1); // khoá khác tính riêng
  assert.deepEqual(read(c), { 'events.eventid': 2, 'issues.issueid': 1 });
});

test('reserve: bảng đã có số lớn hơn bộ đếm (dự phòng cục bộ trước đó) -> cấp số sau nó', async () => {
  const c = fakeClient();
  const svc = new ReportCounterService(c, 'https://x');
  await svc.raise('ST01', 'DRP', { 'events.eventid': 3 });
  assert.equal(await svc.reserve('ST01', 'DRP', 'events.eventid', 7), 8);
  assert.equal(read(c)['events.eventid'], 8);
});

test('reserve: hai người bấm thêm hàng CÙNG LÚC ở hai báo cáo khác nhau -> hai số khác nhau', async () => {
  const c = fakeClient();
  const a = new ReportCounterService(c, 'https://x');
  const b = new ReportCounterService(c, 'https://x');
  await a.raise('ST01', 'DRP', { 'events.eventid': 5 });
  const got = await Promise.all([
    a.reserve('ST01', 'DRP', 'events.eventid', 0),
    b.reserve('ST01', 'DRP', 'events.eventid', 0),
    a.reserve('ST01', 'DRP', 'events.eventid', 0),
    b.reserve('ST01', 'DRP', 'events.eventid', 0)
  ]);
  assert.equal(new Set(got).size, 4);
  assert.deepEqual(got.slice().sort(), [6, 7, 8, 9]);
  assert.equal(read(c)['events.eventid'], 9);
  assert.ok(c.calls.conflict412 > 0); // có xung đột thật và đã được xử lý
});

test('reserve: hai người cùng đặt khi list còn trống (cùng tạo dòng đầu tiên) -> vẫn khác số', async () => {
  const c = fakeClient();
  const got = await Promise.all([
    new ReportCounterService(c, 'https://x').reserve('ST01', 'DRP', 'events.eventid', 0),
    new ReportCounterService(c, 'https://x').reserve('ST01', 'DRP', 'events.eventid', 0)
  ]);
  assert.deepEqual(got.slice().sort(), [1, 2]);
});

test('reserve: lỗi quyền/list thì ném lỗi để nơi gọi dự phòng cục bộ', async () => {
  const c = fakeClient();
  c.createItem = async () => {
    throw new Error('Failed to create item (403): Access denied');
  };
  await assert.rejects(new ReportCounterService(c, 'https://x').reserve('ST01', 'DRP', 'events.eventid', 0), /403/);
});
