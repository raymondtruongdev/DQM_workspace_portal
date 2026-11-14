/**
 * CSS bản in A4 của báo cáo (class `bcn-*`): dùng chung cho cửa sổ in của web (iframe) và file HTML độc lập.
 * Ngoại lệ có chủ đích với quy ước style của dự án (màu hex nằm ngoài tokens.css): đây là CSS của trang in/iframe riêng,
 * màu cố định theo mẫu giấy (tiêu đề mục #014a94, nền hàng tiêu đề #dbe6f1, DAILY REPORT đỏ #c00000).
 */
export const PRINT_CSS = `
@page {
  size: A4 portrait;
  margin: 8mm 10mm 14mm; /* chừa chỗ cho số trang ở lề dưới */
  @bottom-right {
    content: "Trang " counter(page) " / " counter(pages);
    font-family: Arial, Helvetica, sans-serif;
    font-size: 9pt;
    color: #4b5563;
  }
}
* {
  box-sizing: border-box;
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
}
body {
  margin: 0;
  padding: 0;
  font-family: "Segoe UI", Roboto, Arial, sans-serif;
  font-size: 10.5px;
  color: #111827;
  background: #ffffff;
}
.bcn-a4-page {
  width: 100%;
  padding: 0 2px;
  background: #ffffff;
}
.bcn-sec-block {
  page-break-inside: auto;
  break-inside: auto;
}
.bcn-keep-together {
  page-break-inside: avoid;
  break-inside: avoid;
}
.bcn-table {
  width: 100%;
  border-collapse: collapse;
  border: 1px solid #9ca3af;
  margin-bottom: 6px;
  font-size: 10.5px;
  page-break-inside: auto;
  break-inside: auto;
}
.bcn-table thead {
  display: table-header-group;
}
.bcn-table tr {
  page-break-inside: avoid;
  break-inside: avoid;
}
/* Không để mục bị cắt lẻ: tiêu đề + hàng tiêu đề + 2 dòng đầu đi cùng nhau,
   2 dòng cuối đi cùng nhau. Phần giữa vẫn chảy liên tục, không bỏ trống trang. */
.bcn-table thead tr,
.bcn-table tbody tr:first-child,
.bcn-table tbody tr:nth-last-child(2) {
  page-break-after: avoid;
  break-after: avoid;
}
.bcn-table tbody tr:last-child {
  page-break-after: auto;
  break-after: auto;
}
.bcn-table th,
.bcn-table td {
  border: 1px solid #9ca3af;
  padding: 3px 6px;
  vertical-align: middle;
  line-height: 1.28;
}
.bcn-table th:last-child,
.bcn-table td:last-child {
  border-right: 1px solid #9ca3af !important;
}
.bcn-header-logo-cell {
  width: 18%;
  text-align: center;
  font-weight: 800;
  font-size: 14px;
  color: #014a94;
}
.bcn-logo-img {
  display: block;
  margin: 0 auto;
  max-width: 100%;
  max-height: 16mm;
  object-fit: contain;
}
.bcn-header-title-cell {
  width: 64%;
  text-align: center;
  padding: 4px 8px;
}
.bcn-title-vi {
  font-size: 18px;
  font-weight: 800;
  color: #014a94;
  margin: 0;
}
.bcn-title-en {
  font-size: 18px;
  font-weight: 800;
  color: #c00000;
  margin: 1px 0 3px;
}
.bcn-subtitle {
  font-size: 10px;
  color: #374151;
}
.bcn-meta-label {
  background: #f3f4f6;
  font-weight: 700;
  color: #1f2937;
  width: 17%;
}
.bcn-meta-label span {
  font-weight: 400;
  color: #6b7280;
}
.bcn-sec-title {
  background: #014a94 !important;
  color: #ffffff !important;
  font-weight: 700;
  font-size: 11px;
  padding: 4px 8px;
  margin-top: 5px;
  text-align: left;
  border: 1px solid #014a94;
  page-break-after: avoid;
  break-after: avoid;
  page-break-inside: avoid;
  break-inside: avoid;
}
.bcn-sec-title span {
  font-weight: 400;
  color: #e2e8f0 !important;
  opacity: 0.95;
}
.bcn-th-row th {
  background: #dbe6f1;
  color: #1f2937;
  font-weight: 700;
  text-align: center;
}
.bcn-th-row th span {
  display: block;
  font-style: italic;
  font-weight: 400;
  color: #4b5563;
}
.bcn-total-row td {
  background: #f3f4f6;
  font-weight: 700;
}
.bcn-total-row td span {
  font-weight: 400;
  color: #4b5563;
}
.text-center { text-align: center; }
.text-right { text-align: right; }
.font-bold { font-weight: 700; }
.bcn-tone-danger { color: #b91c1c; font-weight: 700; }
.bcn-tone-warning { color: #9a3412; font-weight: 700; }
.bcn-tone-success { color: #15803d; font-weight: 700; }
.bcn-tone-info { color: #1d4ed8; font-weight: 700; }
.bcn-tone-muted { color: #6b7280; font-weight: 700; }
.bcn-preline { white-space: pre-line; }
.bcn-ruled-line {
  min-height: 8mm;
  line-height: 8mm;
  padding: 0 2mm;
  white-space: pre-wrap;
  border-bottom: 1px solid #c9cdd3;
  break-inside: avoid;
}
.bcn-ruled .bcn-ruled-line:last-child { border-bottom: 0; }
.bcn-sign-box {
  height: 110px; /* chừa chỗ trống phía trên tên để ký */
  vertical-align: bottom !important;
  text-align: center;
  padding-bottom: 8px;
}
.bcn-table-photos {
  table-layout: fixed;
}
.bcn-table tr.bcn-photo-row {
  page-break-inside: avoid;
  break-inside: avoid;
}
.bcn-table tr.bcn-photo-dept td {
  background: #1e4e78;
  color: #ffffff;
  font-weight: 700;
  text-align: left;
  padding: 4px 8px;
}
.bcn-table td.bcn-photo-cell {
  padding: 6px;
  vertical-align: top;
}
.bcn-photo-item {
  width: 100%;
  margin: 0;
  border: 1px solid #9ca3af;
  page-break-inside: avoid;
  break-inside: avoid;
}
.bcn-photo-item img {
  display: block;
  width: 100%;
  height: auto;
  aspect-ratio: 4 / 3;
  object-fit: cover;
  background: #f3f4f6;
}
.bcn-photo-cap {
  padding: 3px 6px;
  font-size: 9px;
  color: #4b5563;
  border-top: 1px solid #9ca3af;
  overflow-wrap: anywhere;
}
.bcn-footer-note {
  font-size: 9px;
  color: #4b5563;
  line-height: 1.35;
  margin-top: 4px;
  page-break-inside: avoid;
  break-inside: avoid;
}
`;
