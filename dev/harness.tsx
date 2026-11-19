import * as React from 'react';
import { createRoot } from 'react-dom/client';
import '../src/core/styles';
import { ReportModule } from '../src/modules/report/ReportModule';
import type { ModuleProps } from '../src/core/types';
import { dumpFiles, dumpList, readText, seedFile, seedList } from './LocalSharePointClient';

/**
 * Khung chạy thử cục bộ của module Report: dựng module trong khung `.app-layout > .main-content` như PortalShell,
 * nhưng thư viện tài liệu và list SharePoint là bản giả lập trong bộ nhớ, nạp từ docs/sharepoint_setup/DailyReportLibrary.
 * Dựng bằng `node dev/build.js`, mở `temp/dev/index.html`. Xem AGENTS.md mục "Chạy thử cục bộ".
 */
interface Seed {
  files: Record<string, string>;
  catalogRows: Array<Record<string, unknown>>;
  roleRows: Array<Record<string, unknown>>;
  unitConfigRows: Array<Record<string, unknown>>;
}

const SITE = 'https://dev.local/sites/demo';
const LIBRARY = '/sites/demo/DailyReportLibrary';
const seed = (window as unknown as { __SEED__: Seed }).__SEED__;

Object.keys(seed.files).forEach((rel) => seedFile(`${LIBRARY}/${rel}`, seed.files[rel]));
seedList('DRPRegistry', [
  {
    Title: 'DRP_20261110_ST01',
    Unit: 'ST01',
    DocID: 'DRP',
    ReportDate: '2026-11-10',
    ReportNo: 1,
    Reporter: 'truongquoctoan1@thadico.com',
    Status: 'done',
    SavedAt: '2026-11-10T09:30:00.000Z',
    SignedAt: '2026-11-11T02:00:00.000Z'
  }
]);
seedList('DRPCatalogItems', seed.catalogRows);
seedList('DRPCounters', [{ Title: 'ST01|DRP', Counters: '{"events.eventid":4,"issues.issueid":4}' }]);
// ?nocfg=1: không tạo list DRPUnitConfig để thử thông báo "chưa cấu hình"
if (new URLSearchParams(window.location.search).get('nocfg') !== '1') seedList('DRPUnitConfig', seed.unitConfigRows);
// ?norole=1: không tạo list DRPRole để thử thông báo lỗi
if (new URLSearchParams(window.location.search).get('norole') !== '1') seedList('DRPRole', seed.roleRows);
// ?as=<email>: mở như người dùng THƯỜNG có email đó (mặc định: admin) để thử phân quyền theo list DRPRole
const asEmail = new URLSearchParams(window.location.search).get('as');

(window as unknown as { __dev: unknown }).__dev = { dumpFiles, readText, dumpList, LIBRARY };

const props: ModuleProps = {
  currentUser: asEmail
    ? { username: asEmail, displayName: asEmail, email: asEmail, isAdmin: false }
    : { username: 'dev', displayName: 'Dev Admin', email: 'truongquoctoan1@thadico.com', isAdmin: true },
  spHttpClient: {} as ModuleProps['spHttpClient'],
  siteUrl: SITE,
  onGoHome: () => window.alert('Trang chủ (khung chạy thử)'),
  checkMaintenance: () => Promise.resolve(false),
  // ?todayOnly=1 và ?days=<n> thử cài đặt "Setup DailyReport" của admin
  getModuleSettings: () => {
    const q = new URLSearchParams(window.location.search);
    return Promise.resolve({ todayOnly: q.get('todayOnly') === '1', windowDays: Number(q.get('days')) || 14 });
  }
};

createRoot(document.getElementById('root') as HTMLElement).render(
  <div className="app-layout">
    <main className="main-content">
      <ReportModule {...props} />
    </main>
  </div>
);
