/**
 * Cài đặt module báo cáo do admin đặt ở Admin Config → "Setup DailyReport"
 * (lưu trong `SiteAssets/portal_config.json`, mục `moduleSettings.report`; màn hình admin ghi cùng khoá, xem `DailyReportSetupView`).
 */
export interface ReportSettings {
  /** true = chỉ cho nhập/sửa báo cáo của NGÀY HÔM NAY; ngày cũ hơn tự khoá (chỉ xem) trừ khi admin mở khoá riêng ngày đó. */
  todayOnly: boolean;
  /** Số ngày gần nhất hiện ở lịch và tab Tổng quan. */
  windowDays: number;
  /** Admin bấm OPEN một báo cáo thì sau bấy nhiêu giờ tự khoá lại (cần cột `UnlockUntil` trong DRPRegistry). */
  unlockHours: number;
}

export const REPORT_SETTINGS_ID = 'report';
export const DEFAULT_WINDOW_DAYS = 14;
export const MIN_WINDOW_DAYS = 1;
export const MAX_WINDOW_DAYS = 180; // TẠM nâng để nhập dữ liệu cũ; xong thì trả về 60 (và MAX_DAYS ở DailyReportSetupView.tsx)

export const DEFAULT_UNLOCK_HOURS = 3;
export const MIN_UNLOCK_HOURS = 1;
export const MAX_UNLOCK_HOURS = 72;

export const DEFAULT_REPORT_SETTINGS: ReportSettings = { todayOnly: false, windowDays: DEFAULT_WINDOW_DAYS, unlockHours: DEFAULT_UNLOCK_HOURS };

/** Đọc cài đặt từ dữ liệu thô; thiếu hoặc sai kiểu thì dùng mặc định. */
export function parseReportSettings(raw: Record<string, unknown> | undefined): ReportSettings {
  const days = raw ? Math.round(Number(raw.windowDays)) : NaN;
  const hours = raw ? Math.round(Number(raw.unlockHours)) : NaN;
  return {
    unlockHours: Number.isFinite(hours) ? Math.min(MAX_UNLOCK_HOURS, Math.max(MIN_UNLOCK_HOURS, hours)) : DEFAULT_UNLOCK_HOURS,
    todayOnly: raw ? raw.todayOnly === true : false,
    windowDays: Number.isFinite(days) ? Math.min(MAX_WINDOW_DAYS, Math.max(MIN_WINDOW_DAYS, days)) : DEFAULT_WINDOW_DAYS
  };
}

/** Múi giờ dùng để tính "ngày" của báo cáo: giờ Việt Nam (UTC+7, không có giờ mùa hè). */
const REPORT_UTC_OFFSET_MS = 7 * 3600000;

/** Ngày (YYYY-MM-DD) theo giờ Việt Nam của một thời điểm (ms UTC). Không phụ thuộc múi giờ của máy người dùng. */
export function vnDateStr(ms: number): string {
  return new Date(ms + REPORT_UTC_OFFSET_MS).toISOString().slice(0, 10);
}

/** Ngày theo giờ Việt Nam của một thời điểm, dạng Date lúc 0h theo giờ máy (để dùng với `buildRecentDates`). */
export function vnDateAsLocal(ms: number): Date {
  const [y, m, d] = vnDateStr(ms).split('-').map(Number);
  return new Date(y, m - 1, d);
}

/** Ngày hôm nay (YYYY-MM-DD) theo giờ Việt Nam; `nowMs` nên là giờ máy chủ (xem `ServerClock`). */
export function todayDateStr(nowMs: number = Date.now()): string {
  return vnDateStr(nowMs);
}

/** Ngày báo cáo đã qua (cũ hơn hôm nay). Chuỗi YYYY-MM-DD so sánh theo thứ tự chữ được. */
export function isPastDate(dateStr: string, today: string = todayDateStr()): boolean {
  return dateStr < today;
}

/** Ngày không có dòng khoá riêng trong sổ có bị khoá theo chế độ "chỉ nhập ngày hiện tại" không. */
export function isLockedByTodayOnly(settings: ReportSettings, dateStr: string, today: string = todayDateStr()): boolean {
  return settings.todayOnly && isPastDate(dateStr, today);
}

/** Ngày đầu tiên của cửa sổ `windowDays` ngày gần nhất (hôm nay là ngày cuối), dạng YYYY-MM-DD. */
export function windowStartDate(settings: ReportSettings, today: string = todayDateStr()): string {
  const [y, m, d] = today.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d - (settings.windowDays - 1))).toISOString().slice(0, 10);
}

/** Ngày cũ hơn cửa sổ `windowDays` là LỊCH SỬ: người dùng chỉ xem (admin vẫn mở khoá riêng từng ngày bằng OPEN). */
export function isHistoryDate(settings: ReportSettings, dateStr: string, today: string = todayDateStr()): boolean {
  return dateStr < windowStartDate(settings, today);
}

/** Ngày không có dòng khoá riêng trong sổ có bị khoá do ngày cũ không: chế độ "chỉ nhập ngày hiện tại" hoặc ngày thuộc lịch sử (ngoài `windowDays`). */
export function isPastLocked(settings: ReportSettings, dateStr: string, today: string = todayDateStr()): boolean {
  return isLockedByTodayOnly(settings, dateStr, today) || isHistoryDate(settings, dateStr, today);
}

/** Giờ tự khoá lại, dạng "HH:mm" (kèm "dd/MM" khi không phải hôm nay). */
export function formatUnlockUntil(iso: string): string {
  const d = new Date(iso);
  if (!Number.isFinite(d.getTime())) return '';
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  const hm = `${p(d.getHours())}:${p(d.getMinutes())}`;
  return d.toDateString() === new Date().toDateString() ? hm : `${hm} ${p(d.getDate())}/${p(d.getMonth() + 1)}`;
}

/** Thời gian còn lại đến hạn tự khoá, dạng "2 giờ 15 phút" / "12 phút" (tối thiểu 1 phút, hết hạn trả "0 phút"). */
export function formatUnlockLeft(iso: string, nowMs: number = Date.now()): string {
  const t = Date.parse(iso);
  if (!Number.isFinite(t) || t <= nowMs) return '0 phút';
  const m = Math.ceil((t - nowMs) / 60000);
  return m >= 60 ? `${Math.floor(m / 60)} giờ ${m % 60} phút` : `${m} phút`;
}
