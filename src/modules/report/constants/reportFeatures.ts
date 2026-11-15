/**
 * Cờ bật/tắt tính năng của module báo cáo. Đổi `true` / `false` rồi lưu (dev: tự nạp lại; triển khai: build lại).
 * Tắt một tính năng chỉ ẩn giao diện của nó; dữ liệu đã lưu trong báo cáo vẫn được giữ nguyên.
 */
export const REPORT_FEATURES: { EVENT_PDF: boolean } = {
  /** PDF đính kèm cho từng dòng sự kiện (nút kẹp giấy, thư mục pdf/eventlog, hàng link PDF dưới dòng). */
  EVENT_PDF: true
};
