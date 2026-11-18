import { buildRecentDates } from './reportCalendar';

/**
 * Cài đặt module báo cáo do admin đặt ở Admin Config → "Setup DailyReport"
 * (lưu trong `SiteAssets/portal_config.json`, mục `moduleSettings.report`; màn hình admin ghi cùng khoá, xem `DailyReportSetupView`).
 */
export interface ReportSettings {
  /** true = chỉ cho nhập/sửa báo cáo của NGÀY HÔM NAY; ngày cũ hơn tự khoá (chỉ xem) trừ khi admin mở khoá riêng ngày đó. */
  todayOnly: boolean;
  /** Số ngày gần nhất hiện ở lịch và tab Tổng quan. */
  windowDays: number;
}

export const REPORT_SETTINGS_ID = 'report';
export const DEFAULT_WINDOW_DAYS = 14;
export const MIN_WINDOW_DAYS = 1;
export const MAX_WINDOW_DAYS = 60;

export const DEFAULT_REPORT_SETTINGS: ReportSettings = { todayOnly: false, windowDays: DEFAULT_WINDOW_DAYS };

/** Đọc cài đặt từ dữ liệu thô; thiếu hoặc sai kiểu thì dùng mặc định. */
export function parseReportSettings(raw: Record<string, unknown> | undefined): ReportSettings {
  const days = raw ? Math.round(Number(raw.windowDays)) : NaN;
  return {
    todayOnly: raw ? raw.todayOnly === true : false,
    windowDays: Number.isFinite(days) ? Math.min(MAX_WINDOW_DAYS, Math.max(MIN_WINDOW_DAYS, days)) : DEFAULT_WINDOW_DAYS
  };
}

/** Ngày hôm nay (YYYY-MM-DD, theo giờ máy người dùng). */
export function todayDateStr(): string {
  return buildRecentDates(1)[0];
}

/** Ngày báo cáo đã qua (cũ hơn hôm nay). Chuỗi YYYY-MM-DD so sánh theo thứ tự chữ được. */
export function isPastDate(dateStr: string, today: string = todayDateStr()): boolean {
  return dateStr < today;
}

/** Ngày không có dòng khoá riêng trong sổ có bị khoá theo chế độ "chỉ nhập ngày hiện tại" không. */
export function isLockedByTodayOnly(settings: ReportSettings, dateStr: string): boolean {
  return settings.todayOnly && isPastDate(dateStr);
}
