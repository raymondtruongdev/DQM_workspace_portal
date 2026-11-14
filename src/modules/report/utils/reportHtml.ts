import * as React from 'react';
import { ReportDoc } from '../types/report.types';
import { ReportPrintDocument } from '../components/ReportPrintDocument';
import { PRINT_CSS } from './reportPrintCss';

/** Ảnh nhúng sẵn vào file HTML: logo và ảnh hiện trường (khoá `<id nhóm>/<tên file>`), đều là data URI. */
export interface ReportHtmlAssets {
  logos: { thaco?: string; crec?: string };
  photos: Record<string, string>;
}

/** Tên file PDF mặc định khi in = Mã tài liệu (thiếu thì BCN_<đơn vị>_<ngày>). */
export function getPdfFileName(doc: Pick<ReportDoc, 'unit' | 'dateStr' | 'metadata'>): string {
  return (doc.metadata.docCode || `BCN_${doc.unit}_${doc.dateStr}`).replace(/[\\/:*?"<>|\s]+/g, '_');
}

/** Khung giấy A4 + thanh công cụ khi xem trên màn hình (khi in chỉ còn nội dung giấy). */
const SCREEN_CSS = `
  .bcn-toolbar { display: none; }
  @media screen {
    body { background: #eef1f5; padding: 0 12px 32px; }
    .bcn-toolbar {
      display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap;
      max-width: 794px; margin: 0 auto; padding: 14px 0; font-family: "Segoe UI", Roboto, Arial, sans-serif;
    }
    .bcn-toolbar__title { font-size: 14px; font-weight: 700; color: #1f2937; overflow-wrap: anywhere; min-width: 0; }
    .bcn-toolbar__btn {
      min-height: 40px; padding: 0 16px; border: 0; border-radius: 8px; background: #1d4ed8; color: #ffffff;
      font: 600 14px "Segoe UI", Roboto, Arial, sans-serif; cursor: pointer;
    }
    .bcn-toolbar__btn:hover { background: #1e40af; }
    .bcn-sheet { max-width: 794px; margin: 0 auto; }
    .bcn-a4-page {
      background: #ffffff; border: 1px solid #cbd5e1; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
      padding: 28px 32px; color: #111827; font-size: 11.5px;
    }
    .bcn-table { margin-bottom: 8px; font-size: 11.5px; }
    .bcn-table th, .bcn-table td { padding: 5px 7px; line-height: 1.35; }
    .bcn-header-logo-cell { font-size: 15px; }
    .bcn-header-title-cell { padding: 6px 8px; }
    .bcn-subtitle { font-size: 11px; }
    .bcn-sec-title { font-size: 12px; padding: 5px 8px; margin-top: 6px; }
  }
  @media screen and (max-width: 820px) {
    body { overflow-x: auto; }
    .bcn-sheet, .bcn-toolbar { min-width: 680px; }
    .bcn-a4-page { padding: 18px 16px; }
  }
  @media print { .bcn-toolbar { display: none !important; } }
`;

const esc = (s: string): string => s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));

/**
 * File HTML độc lập của báo cáo: nội dung bản in A4 đã render sẵn (giống hệt cửa sổ xem trước của web) + CSS in + logo/ảnh nhúng
 * + toàn bộ dữ liệu `.data.json` trong thẻ `<script type="application/json" id="report-data">`.
 * Mở bằng trình duyệt không cần mạng hay web server; bấm nút In để lưu PDF. `react-dom/server` chỉ được tải khi gọi hàm này.
 */
export async function buildStandaloneHtml(doc: ReportDoc, jsonText: string, assets: ReportHtmlAssets): Promise<string> {
  const { renderToStaticMarkup } = await import(/* webpackChunkName: 'report-html' */ 'react-dom/server');
  const body = renderToStaticMarkup(
    React.createElement(ReportPrintDocument, {
      doc,
      logoSrc: assets.logos,
      photoSrc: (groupId: string, name: string) => assets.photos[`${groupId}/${name}`] || `photos/${groupId}/${name}`
    })
  );
  const title = getPdfFileName(doc);
  // "<" trong JSON thành < để dữ liệu không thể đóng thẻ <script> sớm; JSON.parse vẫn đọc đúng
  const data = jsonText.replace(/</g, '\\u003c');
  return `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${esc(title)}</title>
<style>${PRINT_CSS}${SCREEN_CSS}</style>
</head>
<body>
<div class="bcn-toolbar">
  <div class="bcn-toolbar__title">${esc(doc.metadata.docCode || title)}</div>
  <button type="button" class="bcn-toolbar__btn" onclick="window.print()">In / Lưu thành PDF (A4)</button>
</div>
<div class="bcn-sheet">${body}</div>
<script type="application/json" id="report-data">
${data}</script>
</body>
</html>
`;
}
