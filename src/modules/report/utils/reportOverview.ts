import { RoleEntry } from '../../../core/types';
import { ReportDayInfo, ReportDayStatus } from '../types/report.types';

/** Người phụ trách một báo cáo (lấy từ list ReportRole). */
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

/** Cách hiển thị một người phụ trách dạng thẻ nhỏ: chữ cái đầu, nhãn ngắn (phần sau " - " nếu có), chú giải đầy đủ. */
export function ownerDisplay(o: ReportOwner): { initials: string; label: string; title: string } {
  const sep = ' - ';
  const cut = o.name.indexOf(sep);
  const person = (cut === -1 ? o.name : o.name.slice(0, cut)).trim();
  const org = cut === -1 ? '' : o.name.slice(cut + sep.length).trim();
  const words = person.split(/\s+/).filter(Boolean);
  let initials = '';
  if (words.length >= 2) initials = words[0].charAt(0) + words[words.length - 1].charAt(0);
  else if (words.length === 1) initials = words[0].slice(0, 2);
  else initials = (o.email.split('@')[0] || '?').slice(0, 2);
  return {
    initials: initials.toUpperCase(),
    label: org || person || o.email,
    title: [o.name, o.email].filter(Boolean).join('\n')
  };
}
