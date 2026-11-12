import * as React from 'react';
import { ReportDoc } from '../types/report.types';
import { buildSectionNumbering } from '../utils/reportSections';
import { ReportPrintTable } from './ReportPrintTable';
import { ReportPrintForm } from './ReportPrintForm';
import { ReportPrintPhoto } from './ReportPrintPhoto';
import { formatTemplateVersionDisplay } from '../utils/reportFileNaming';

export interface IReportPrintPreviewModalProps {
  doc: ReportDoc;
  /** Thư mục ảnh của báo cáo (server-relative): data/<Unit>/<DocID>_<yyyyMMdd>/photos */
  photosFolder: string;
  onClose: () => void;
}

function formatDateVN(dateStr: string): string {
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  return dateStr;
}

export const ReportPrintPreviewModal: React.FC<IReportPrintPreviewModalProps> = ({
  doc,
  photosFolder,
  onClose
}) => {
  const printAreaRef = React.useRef<HTMLDivElement | null>(null);
  const m = doc.metadata;
  const num = buildSectionNumbering(doc);

  const handlePrint = (): void => {
    if (!printAreaRef.current) return;
    const htmlContent = printAreaRef.current.innerHTML;

    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);

    const iframeDoc = iframe.contentWindow?.document;
    if (!iframeDoc) {
      document.body.removeChild(iframe);
      return;
    }

    // Tên file PDF mặc định = Mã tài liệu (vd HCM2_1-23_CS-THAC_DCEM-CPZZ-ST01-CWS-DRP-00001); thiếu thì BCN_<unit>_<yyyy-MM-dd>
    const pdfFileName = (m.docCode || `BCN_${doc.unit}_${doc.dateStr}`).replace(/[\\/:*?"<>|\s]+/g, '_');

    iframeDoc.open();
    iframeDoc.write(`<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${pdfFileName}</title>
  <style>
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
  </style>
</head>
<body>
  ${htmlContent}
</body>
</html>`);
    iframeDoc.close();

    // Đợi mọi ảnh (phụ lục ảnh hiện trường) tải xong rồi mới mở hộp thoại in
    const waitForImages = (): Promise<void> =>
      new Promise<void>((resolve) => {
        const pending = Array.from(iframeDoc.images).filter((img) => !img.complete);
        if (pending.length === 0) {
          resolve();
          return;
        }
        let left = pending.length;
        const done = (): void => {
          left -= 1;
          if (left <= 0) resolve();
        };
        pending.forEach((img) => {
          img.addEventListener('load', done);
          img.addEventListener('error', done);
        });
        window.setTimeout(resolve, 20000); // không chờ quá 20s
      });

    waitForImages()
      .then(() => {
        window.setTimeout(() => {
          // Chrome/Edge lấy tên file PDF từ tiêu đề của TRANG CHÍNH (không phải iframe):
          // đặt tạm document.title rồi trả lại sau khi in xong.
          const previousTitle = document.title;
          let restored = false;
          const restoreTitle = (): void => {
            if (restored) return;
            restored = true;
            document.title = previousTitle;
            window.removeEventListener('afterprint', restoreTitle);
          };
          document.title = pdfFileName;
          window.addEventListener('afterprint', restoreTitle);
          iframe.contentWindow?.addEventListener('afterprint', restoreTitle);
          window.setTimeout(restoreTitle, 120000); // dự phòng nếu trình duyệt không bắn afterprint

          iframe.contentWindow?.focus();
          iframe.contentWindow?.print();
          setTimeout(() => {
            if (document.body.contains(iframe)) {
              document.body.removeChild(iframe);
            }
          }, 1500);
        }, 150);
      })
      .catch((err: unknown) => console.error('[print] image wait failed:', err));
  };

  return (
    <div className="bcn-preview-backdrop" onClick={onClose}>
      <div className="bcn-preview-modal" onClick={(e) => e.stopPropagation()}>
        {/* Top Modal Action Bar */}
        <div className="bcn-preview-toolbar">
          <div className="bcn-preview-toolbar-left">
            <span className="bcn-preview-modal-title">
              Xem bản in A4 — {m.docCode || `${doc.unit} (${doc.dateStr})`}
            </span>
          </div>
          <div className="bcn-preview-toolbar-right">
            <button type="button" className="btn btn--primary" onClick={handlePrint}>
              <svg
                width="16"
                height="16"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <polyline points="6 9 6 2 18 2 18 9" />
                <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
                <rect x="6" y="14" width="12" height="8" />
              </svg>
              <span>In / Lưu thành PDF (A4)</span>
            </button>
            <button type="button" className="btn btn--secondary" onClick={onClose}>
              Đóng
            </button>
          </div>
        </div>

        {/* Scrollable 2-Page A4 Preview Container */}
        <div className="bcn-preview-scroll-body">
          <div ref={printAreaRef} className="bcn-print-container">
            {/* ==================== PAGE 1 ==================== */}
            <div className="bcn-a4-page">
              {/* Top Title Banner Table */}
              <table className="bcn-table">
                <tbody>
                  <tr>
                    <td className="bcn-header-logo-cell">THACO</td>
                    <td className="bcn-header-title-cell">
                      <div className="bcn-title-vi">{m.docTitle || 'BÁO CÁO NGÀY'}</div>
                      <div className="bcn-title-en">{m.docTitleEn || 'DAILY REPORT'}</div>
                      {m.subTitle.split('|').map((line, i) => (
                        <div key={i} className="bcn-subtitle">
                          {line.trim()}
                        </div>
                      ))}
                    </td>
                    <td className="bcn-header-logo-cell">CREC</td>
                  </tr>
                </tbody>
              </table>

              {/* Document Metadata Table */}
              <table className="bcn-table">
                <tbody>
                  <tr>
                    <td className="bcn-meta-label">
                      Mã tài liệu <span>| Doc. code</span>
                    </td>
                    <td className="font-bold" colSpan={3} style={{ wordBreak: 'break-all' }}>
                      {m.docCode}
                    </td>
                  </tr>
                  <tr>
                    <td className="bcn-meta-label">
                      Vị trí <span>| Location</span>
                    </td>
                    <td style={{ width: '49%' }}>{m.location}</td>
                    <td className="bcn-meta-label" style={{ width: '20%', whiteSpace: 'nowrap' }}>
                      Số BCN <span>| Report no.</span>
                    </td>
                    <td style={{ width: '14%' }}>{m.reportNo}</td>
                  </tr>
                  <tr>
                    <td className="bcn-meta-label">
                      Đơn vị thi công <span>| Subcontractor</span>
                    </td>
                    <td>{m.subcontractor}</td>
                    <td className="bcn-meta-label" style={{ whiteSpace: 'nowrap' }}>
                      Ngày <span>| Date</span>
                    </td>
                    <td style={{ width: '14%' }}>{formatDateVN(doc.dateStr)}</td>
                  </tr>
                </tbody>
              </table>

              {/* Các mục của báo cáo (mỗi sheet của template = một mục, theo thứ tự sheet) */}
              {doc.sections.map((sec) =>
                sec.kind === 'table' ? (
                  <ReportPrintTable key={sec.key} table={sec} roman={num.roman(sec.key)} dateStr={doc.dateStr} />
                ) : sec.kind === 'photo' ? (
                  <ReportPrintPhoto
                    key={sec.key}
                    section={sec}
                    roman={num.roman(sec.key)}
                    photosFolder={photosFolder}
                  />
                ) : (
                  <ReportPrintForm key={sec.key} form={sec} roman={num.roman(sec.key)} dateStr={doc.dateStr} />
                )
              )}

              <div className="bcn-keep-together">
                <div className="bcn-footer-note">
                  <div>
                    Sinh tự động từ Microsoft 365 (SharePoint)
                  </div>
                  <div>
                    Ảnh: {`data/${doc.unit}/${doc.docId}_${doc.dateStr.replace(/-/g, '')}/photos/`}
                    {' | '}
                    Template version:{' '}
                    {formatTemplateVersionDisplay(m.templateVersion || doc.latestTemplateInfo?.version)}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
