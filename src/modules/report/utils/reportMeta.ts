import { ReportMetadata } from '../types/report.types';

/** Số báo cáo luôn 5 chữ số: 1 -> "00001". */
export function formatReportNo(n: number): string {
  return ('00000' + Math.max(1, Math.floor(n))).slice(-5);
}

/** Số báo cáo kế tiếp của đơn vị, tính từ số của báo cáo ngày trước ("00041" -> "00042"; không đọc được -> "00001"). */
export function nextReportNo(prev: string): string {
  const n = parseInt(String(prev || '').replace(/\D/g, ''), 10);
  return formatReportNo(isNaN(n) ? 1 : n + 1);
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
