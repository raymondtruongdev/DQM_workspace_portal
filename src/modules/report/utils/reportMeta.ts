import { ReportMetadata } from '../types/report.types';

/** Số báo cáo luôn 5 chữ số: 1 -> "00001". */
export function formatReportNo(n: number): string {
  return ('00000' + Math.max(1, Math.floor(n))).slice(-5);
}

/** Chuỗi ngày `yyyy-MM-dd` hợp lệ thật (không chấp nhận 2026-02-30) -> số ngày kể từ 1970-01-01 (UTC); không hợp lệ -> undefined. */
function dayNumber(dateStr: string): number | undefined {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr || '').trim());
  if (!m) return undefined;
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const ms = Date.UTC(y, mo - 1, d);
  const back = new Date(ms);
  if (back.getUTCFullYear() !== y || back.getUTCMonth() !== mo - 1 || back.getUTCDate() !== d) return undefined;
  return Math.round(ms / 86400000);
}

/** Ngày bắt đầu của nhà ga (cột StartDate): chỉ nhận `yyyy-MM-dd` hợp lệ, sai định dạng thì coi như chưa cấu hình. */
export function parseStartDate(raw: unknown): string | undefined {
  const s = raw === undefined || raw === null ? '' : String(raw).trim();
  return dayNumber(s) === undefined ? undefined : s;
}

/**
 * Số báo cáo của một ngày = (ngày báo cáo - ngày bắt đầu của nhà ga) + 1 (ngày bắt đầu là 1).
 * Ngày trước ngày bắt đầu hoặc chuỗi ngày không hợp lệ thì trả undefined (không có số).
 */
export function reportNoForDate(startDate: string, dateStr: string): number | undefined {
  const start = dayNumber(startDate);
  const day = dayNumber(dateStr);
  if (start === undefined || day === undefined || day < start) return undefined;
  return day - start + 1;
}

/** Mã tài liệu = mẫu của template với {unit} (mã đơn vị) và {reportNo} (số báo cáo). */
export function expandDocCode(pattern: string, unit: string, reportNo: string): string {
  return pattern.replace(/\{unit\}/gi, unit).replace(/\{reportNo\}/gi, reportNo);
}

/**
 * Đồng bộ các trường tự động của thông tin chung: Vị trí = đơn vị, Mã tài liệu theo mẫu + số báo cáo.
 * Dùng sau khi đọc template/báo cáo và mỗi khi số báo cáo đổi.
 */
export function syncMetadata(m: ReportMetadata, unit: string): ReportMetadata {
  const next: ReportMetadata = { ...m, location: unit };
  if (next.docCodePattern) next.docCode = expandDocCode(next.docCodePattern, unit, next.reportNo || '');
  return next;
}
