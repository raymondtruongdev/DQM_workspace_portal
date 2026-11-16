/**
 * Cờ bật/tắt tính năng của module báo cáo. Đổi `true` / `false` rồi lưu (dev: tự nạp lại; triển khai: build lại).
 * Tắt một tính năng chỉ ẩn giao diện của nó; dữ liệu đã lưu trong báo cáo vẫn được giữ nguyên.
 */
export const REPORT_FEATURES: { EVENT_PDF: boolean; PRINT_FOOTER_NOTE: boolean } = {
  /** PDF đính kèm cho từng dòng sự kiện (nút kẹp giấy, thư mục pdf/eventlog, hàng link PDF dưới dòng). */
  EVENT_PDF: true,
  /** Dòng chú thích cuối bản in A4/HTML: "Sinh tự động từ Microsoft 365 (SharePoint)" + đường dẫn thư mục ảnh + phiên bản template. */
  PRINT_FOOTER_NOTE: true
};
