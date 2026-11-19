/** Test đánh số báo cáo theo ngày bắt đầu của nhà ga (list DRPUnitConfig). Chạy: npm run test:report */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const base = path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report');
const M = require(path.join(base, 'utils', 'reportMeta'));
const { ReportUnitConfigService } = require(path.join(base, 'services', 'ReportUnitConfigService'));
const { ReportDataService, ReportNumberingError } = require(path.join(base, 'services', 'ReportDataService'));

test('reportNoForDate: ngày bắt đầu = 1, các ngày sau tăng dần', () => {
  assert.equal(M.reportNoForDate('2026-10-20', '2026-10-20'), 1);
  assert.equal(M.reportNoForDate('2026-10-20', '2026-10-30'), 11);
  assert.equal(M.reportNoForDate('2026-10-20', '2026-11-19'), 31); // qua tháng
  assert.equal(M.reportNoForDate('2026-12-25', '2027-01-05'), 12); // qua năm
  assert.equal(M.reportNoForDate('2028-02-28', '2028-03-01'), 3); // năm nhuận (29/02)
  assert.equal(M.reportNoForDate('2026-03-01', '2026-10-25'), 239);
});

test('reportNoForDate: trước ngày bắt đầu hoặc ngày không hợp lệ -> undefined', () => {
  assert.equal(M.reportNoForDate('2026-10-20', '2026-10-19'), undefined);
  assert.equal(M.reportNoForDate('2026-10-20', ''), undefined);
  assert.equal(M.reportNoForDate('2026-10-20', '2026-02-30'), undefined);
  assert.equal(M.reportNoForDate('20/10/2026', '2026-10-30'), undefined);
  assert.equal(M.reportNoForDate('', '2026-10-30'), undefined);
});

test('parseStartDate: chỉ nhận yyyy-MM-dd hợp lệ thật', () => {
  assert.equal(M.parseStartDate('2026-10-20'), '2026-10-20');
  assert.equal(M.parseStartDate(' 2026-10-20 '), '2026-10-20');
  assert.equal(M.parseStartDate('2026-10-2'), undefined);
  assert.equal(M.parseStartDate('20/10/2026'), undefined);
  assert.equal(M.parseStartDate('2026-02-30'), undefined);
  assert.equal(M.parseStartDate(46315), undefined); // số ngày của Excel
  assert.equal(M.parseStartDate(undefined), undefined);
});

test('formatReportNo: 5 chữ số', () => {
  assert.equal(M.formatReportNo(1), '00001');
  assert.equal(M.formatReportNo(11), '00011');
});

const fakeClient = (rows, fail) => ({
  getListFieldMap: async () => ({ title: 'Title', unit: 'Unit0', startdate: 'StartDate', geotag: 'GeoTag', active: 'Active' }),
  getItems: async () => {
    if (fail) throw new Error('404 list does not exist');
    return rows;
  }
});

test('ReportUnitConfigService: đọc StartDate/GeoTag theo tên nội bộ, bỏ Active = 0, sai định dạng = chưa cấu hình', async () => {
  const svc = new ReportUnitConfigService(
    fakeClient([
      { Title: 'ST01', Unit0: 'ST01', StartDate: '2026-10-20', GeoTag: '10.77, 106.69', Active: 1 },
      { Title: 'ST02', Unit0: 'ST02', StartDate: '20/10/2026', GeoTag: '', Active: 1 },
      { Title: 'ST03', Unit0: 'ST03', StartDate: '2026-10-20', GeoTag: 'x', Active: 0 },
      { Title: 'ST01', Unit0: 'ST01', StartDate: '2030-01-01', GeoTag: 'trùng', Active: 1 }
    ]),
    'https://x'
  );
  assert.deepEqual(await svc.get('st01'), { unit: 'ST01', startDate: '2026-10-20', geoTag: '10.77, 106.69' }); // không phân biệt hoa thường, dòng đầu thắng
  assert.equal((await svc.get('ST02')).startDate, undefined);
  assert.equal(await svc.get('ST03'), undefined);
  assert.equal(await svc.get('ST09'), undefined);
});

const dataService = (rows, fail) => {
  const ds = Object.create(ReportDataService.prototype);
  ds._unitConfig = new ReportUnitConfigService(fakeClient(rows, fail), 'https://x');
  return ds;
};

test('resolveNumbering: có cấu hình -> số theo ngày + tọa độ', async () => {
  const ds = dataService([{ Title: 'ST01', Unit0: 'ST01', StartDate: '2026-10-20', GeoTag: '10.77, 106.69', Active: 1 }]);
  assert.deepEqual(await ds.resolveNumbering('ST01', '2026-10-30'), { reportNo: '00011', geoTag: '10.77, 106.69' });
});

test('resolveNumbering: thiếu cấu hình, trước ngày bắt đầu, không đọc được list -> ReportNumberingError', async () => {
  const ds = dataService([{ Title: 'ST01', Unit0: 'ST01', StartDate: '2026-10-20', GeoTag: '', Active: 1 }]);
  await assert.rejects(ds.resolveNumbering('ST02', '2026-10-30'), (e) => e instanceof ReportNumberingError && e.reason === 'noConfig' && /liên hệ Admin/.test(e.message));
  await assert.rejects(ds.resolveNumbering('ST01', '2026-10-19'), (e) => e instanceof ReportNumberingError && e.reason === 'beforeStart' && /19\/10\/2026/.test(e.message));
  await assert.rejects(dataService([], true).resolveNumbering('ST01', '2026-10-30'), (e) => e instanceof ReportNumberingError && e.reason === 'configError');
});
