import { RoleEntry } from '../../../core/types';
import { ReportDayInfo, ReportDayStatus } from '../types/report.types';
import { DAY_STATUS_LABEL, formatDM, weekdayLong } from './reportCalendar';

/** Người phụ trách một báo cáo (lấy từ Report_Role.xlsx). */
export interface ReportOwner {
  name: string;
  email: string;
}

/** Một đơn vị + loại báo cáo cần theo dõi trong màn hình Tổng quan. */
export interface OverviewTarget {
  key: string;
  unit: string;
  docId: string;
  /** Vai trò dùng để mở lịch / form của đơn vị này. */
  role: RoleEntry;
  owners: ReportOwner[];
}

export interface OverviewRow {
  target: OverviewTarget;
  state: 'pending' | 'ready' | 'error';
  infos: Record<string, ReportDayInfo>;
}

export interface ReportListItem {
  target: OverviewTarget;
  dateStr: string;
  status: ReportDayStatus;
  info?: ReportDayInfo;
}

export type StatusFilter = 'ALL' | ReportDayStatus;

const collator = new Intl.Collator('vi', { numeric: true, sensitivity: 'base' });

/** Gom các phân công báo cáo ngày theo (đơn vị, loại báo cáo), xếp đơn vị theo số thứ tự tự nhiên. */
export function buildTargets(assignments: RoleEntry[]): OverviewTarget[] {
  const map: Record<string, OverviewTarget> = {};
  assignments.forEach((r) => {
    if (r.docId.toLowerCase().includes('week')) return; // báo cáo tuần là placeholder, chưa có lịch trạng thái
    const unit = r.unit.trim();
    const key = `${unit.toLowerCase()}|${r.docId.toLowerCase()}`;
    if (!map[key]) {
      map[key] = { key, unit, docId: r.docId, role: r, owners: [] };
    }
    const name = (r.siteManager || '').trim();
    const email = (r.email || '').trim();
    const exists = map[key].owners.some((o) => o.email.toLowerCase() === email.toLowerCase() && o.name === name);
    if ((name || email) && !exists) map[key].owners.push({ name, email });
  });
  return Object.keys(map)
    .map((k) => map[k])
    .sort((a, b) => collator.compare(a.unit, b.unit) || collator.compare(a.docId, b.docId));
}

export function statusOfDay(row: OverviewRow, dateStr: string): ReportDayStatus {
  const info = row.infos[dateStr];
  return info ? info.status : 'missing';
}

/** Đếm số đơn vị theo trạng thái của một ngày (chỉ tính các hàng đã đọc xong). */
export function countByStatus(rows: OverviewRow[], dateStr: string): Record<ReportDayStatus, number> {
  const c: Record<ReportDayStatus, number> = { missing: 0, unsigned: 0, done: 0 };
  rows.forEach((r) => {
    if (r.state !== 'ready') return;
    c[statusOfDay(r, dateStr)] += 1;
  });
  return c;
}

/** Số báo cáo (đơn vị × ngày) theo trạng thái trong cả khoảng ngày, chỉ tính các hàng đã đọc xong. */
export function countWindow(rows: OverviewRow[], dates: string[]): { counts: Record<ReportDayStatus, number>; total: number } {
  const counts: Record<ReportDayStatus, number> = { missing: 0, unsigned: 0, done: 0 };
  let total = 0;
  rows.forEach((r) => {
    if (r.state !== 'ready') return;
    dates.forEach((d) => {
      counts[statusOfDay(r, d)] += 1;
      total += 1;
    });
  });
  return { counts, total };
}

/** Số ngày hoàn tất / tổng số ngày của một hàng. */
export function doneCount(row: OverviewRow, dates: string[]): number {
  return dates.filter((d) => statusOfDay(row, d) === 'done').length;
}

export function matchesSearch(target: OverviewTarget, term: string): boolean {
  const t = term.trim().toLowerCase();
  if (!t) return true;
  if (target.unit.toLowerCase().includes(t) || target.docId.toLowerCase().includes(t)) return true;
  return target.owners.some((o) => o.name.toLowerCase().includes(t) || o.email.toLowerCase().includes(t));
}

/** Mọi cặp (đơn vị, ngày) với đủ trạng thái, ngày mới nhất lên trước; có thể lọc theo trạng thái và từ khoá. */
export function buildReportList(rows: OverviewRow[], dates: string[], filter: StatusFilter, term: string): ReportListItem[] {
  const items: ReportListItem[] = [];
  rows.forEach((r) => {
    if (r.state !== 'ready' || !matchesSearch(r.target, term)) return;
    dates.forEach((d) => {
      const st = statusOfDay(r, d);
      if (filter !== 'ALL' && st !== filter) return;
      items.push({ target: r.target, dateStr: d, status: st, info: r.infos[d] });
    });
  });
  return items.sort((a, b) =>
    a.dateStr === b.dateStr ? collator.compare(a.target.unit, b.target.unit) : a.dateStr < b.dateStr ? 1 : -1
  );
}

export function ownersText(owners: ReportOwner[]): string {
  if (owners.length === 0) return 'Chưa có người phụ trách';
  return owners.map((o) => (o.name && o.email ? `${o.name} (${o.email})` : o.name || o.email)).join('; ');
}

/** Văn bản dán vào email / Teams từ danh sách đang hiển thị. */
export function reportListToText(items: ReportListItem[]): string {
  return items
    .map(
      (i) =>
        `${i.target.unit} · ${weekdayLong(i.dateStr)} ${formatDM(i.dateStr)} · ${DAY_STATUS_LABEL[i.status]} · ${ownersText(i.target.owners)}`
    )
    .join('\n');
}

/** Chạy `worker` cho từng phần tử, tối đa `limit` việc cùng lúc. */
export async function runWithConcurrency<T>(items: T[], limit: number, worker: (item: T) => Promise<void>): Promise<void> {
  let next = 0;
  const runners: Promise<void>[] = [];
  const size = Math.max(1, Math.min(limit, items.length));
  for (let i = 0; i < size; i++) {
    runners.push(
      (async (): Promise<void> => {
        while (next < items.length) {
          const idx = next;
          next += 1;
          await worker(items[idx]);
        }
      })()
    );
  }
  await Promise.all(runners);
}
