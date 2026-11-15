/** Test PDF đính kèm dòng sự kiện: giữ nguyên tên file, phát hiện trùng tên, cờ tính năng. Chạy: npm run test:report */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const src = path.join(root, 'temp', 'dev', 'js', 'src', 'modules', 'report');
const { ReportFileService } = require(path.join(src, 'services', 'ReportFileService'));
const { REPORT_FEATURES } = require(path.join(src, 'constants', 'reportFeatures'));
const J = require(path.join(src, 'utils', 'reportJson'));

/** Client giả: nhớ thư mục đã tạo và file đã upload. */
function fakeClient() {
  const files = {};
  return {
    files,
    folders: [],
    ensureFolder: async function (_s, p) {
      this.folders.push(p);
    },
    uploadFile: async (_s, folder, name, buf) => {
      files[`${folder}/${name}`] = buf;
    },
    getFolderFiles: async (_s, folder) =>
      Object.keys(files)
        .filter((k) => k.startsWith(`${folder}/`) && !k.slice(folder.length + 1).includes('/'))
        .map((k) => ({ name: k.slice(folder.length + 1), serverRelativeUrl: k, length: 1, timeLastModified: '' }))
  };
}
const svc = (c) => new ReportFileService(c, 'https://x.sharepoint.com/sites/s');
const buf = new ArrayBuffer(4);

test('giữ nguyên tên file PDF, chỉ đổi ký tự SharePoint cấm', () => {
  assert.equal(ReportFileService.eventPdfStoredName('bien-ban nghiem thu.pdf'), 'bien-ban nghiem thu.pdf');
  assert.equal(ReportFileService.eventPdfStoredName('Biên bản 12_11.pdf'), 'Biên bản 12_11.pdf');
  assert.equal(ReportFileService.eventPdfStoredName('a#b:c.pdf'), 'a-b-c.pdf');
});

test('upload vào pdf/eventlog của thư mục báo cáo bằng đúng tên gốc, rồi phát hiện trùng tên (không phân biệt hoa thường)', async () => {
  const c = fakeClient();
  const s = svc(c);
  const name = await s.uploadEventPdf('ST01', 'DailyReport', '2026-11-14', 'Bien-ban.pdf', buf);
  assert.equal(name, 'Bien-ban.pdf');
  const folder = s.getReportEventPdfFolderPath('ST01', 'DailyReport', '2026-11-14');
  assert.match(folder, /data\/ST01\/DRP_20261114_ST01\/pdf\/eventlog$/);
  assert.ok(c.files[`${folder}/Bien-ban.pdf`]);
  assert.equal(await s.eventPdfExists('ST01', 'DailyReport', '2026-11-14', 'Bien-ban.pdf'), true);
  assert.equal(await s.eventPdfExists('ST01', 'DailyReport', '2026-11-14', 'bien-BAN.pdf'), true);
  assert.equal(await s.eventPdfExists('ST01', 'DailyReport', '2026-11-14', 'khac.pdf'), false);
  assert.equal(await s.eventPdfExists('ST01', 'DailyReport', '2026-11-15', 'Bien-ban.pdf'), false); // báo cáo ngày khác
});

test('cờ REPORT_FEATURES.EVENT_PDF quyết định mục Sự kiện có rowPdf hay không', () => {
  assert.equal(typeof REPORT_FEATURES.EVENT_PDF, 'boolean');
  const tpl = J.parseTemplateJson(require('fs').readFileSync(path.join(root, 'docs', 'sharepoint_setup', 'DailyReportLibrary', '_template', 'template_DRP_ST01_20261114.json'), 'utf8'));
  const ev = tpl.sections.find((s) => s.id === 'events');
  assert.equal(Boolean(ev.rowPdf), REPORT_FEATURES.EVENT_PDF);
});
