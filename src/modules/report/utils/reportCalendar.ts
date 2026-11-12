import { ReportDayStatus } from '../types/report.types';

export const WEEKDAY_LONG = ['Chủ nhật', 'Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy'];
export const MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export const WEEKDAY_SHORT_MON_FIRST = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'];

/** Nhãn hiển thị của từng trạng thái ngày. */
export const DAY_STATUS_LABEL: Record<ReportDayStatus, string> = {
  missing: 'Chưa nhập',
  unsigned: 'Chưa upload bản ký',
  done: 'Hoàn tất'
};

/** Biến thể `badge--*` của core ứng với từng trạng thái. */
export const DAY_STATUS_BADGE: Record<ReportDayStatus, string> = {
  missing: 'badge--danger',
  unsigned: 'badge--warning',
  done: 'badge--success'
};

function pad(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function toDateStr(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** Parse YYYY-MM-DD thành Date theo giờ địa phương (không lệch múi giờ). */
export function parseDateStr(dateStr: string): Date {
  const [y, m, d] = dateStr.split('-').map((x) => parseInt(x, 10));
  return new Date(y, (m || 1) - 1, d || 1);
}

/** Ngày sau `dateStr` n ngày (n có thể âm), dạng YYYY-MM-DD. */
export function addDaysStr(dateStr: string, n: number): string {
  const d = parseDateStr(dateStr);
  return toDateStr(new Date(d.getFullYear(), d.getMonth(), d.getDate() + n));
}

/** `count` ngày kết thúc tại `today`, xếp từ cũ đến mới (YYYY-MM-DD). */
export function buildRecentDates(count: number, today: Date = new Date()): string[] {
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const out: string[] = [];
  for (let i = count - 1; i >= 0; i--) {
    out.push(toDateStr(new Date(base.getFullYear(), base.getMonth(), base.getDate() - i)));
  }
  return out;
}

/** Ngày bắt đầu từ thứ hai: T2 = 0 … CN = 6. */
export function mondayIndex(dateStr: string): number {
  return (parseDateStr(dateStr).getDay() + 6) % 7;
}

export function weekdayLong(dateStr: string): string {
  return WEEKDAY_LONG[parseDateStr(dateStr).getDay()];
}

/** Tên tháng viết tắt (Jan … Dec) của một ngày. */
export function monthShort(dateStr: string): string {
  return MONTH_SHORT[parseDateStr(dateStr).getMonth()];
}

/** dd/MM */
export function formatDM(dateStr: string): string {
  const d = parseDateStr(dateStr);
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}`;
}

/** dd/MM/yyyy */
export function formatDMY(dateStr: string): string {
  return `${formatDM(dateStr)}/${parseDateStr(dateStr).getFullYear()}`;
}

/** HH:mm · dd/MM từ một mốc ISO. */
export function formatSavedAt(iso?: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '—';
  return `${pad(d.getHours())}:${pad(d.getMinutes())} · ${pad(d.getDate())}/${pad(d.getMonth() + 1)}`;
}

/** HH:mm nếu cùng ngày với `refDateStr` (YYYY-MM-DD), ngược lại "HH:mm · dd/MM". */
export function formatTimeRelative(iso: string | undefined, refDateStr: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const time = `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const sameDay = toDateStr(d) === refDateStr;
  return sameDay ? time : `${time} · ${pad(d.getDate())}/${pad(d.getMonth() + 1)}`;
}

export function formatFileSize(bytes: number): string {
  if (!bytes || bytes < 0) return '—';
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
