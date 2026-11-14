import * as React from 'react';
import { ReportDoc } from '../types/report.types';
import { PRINT_CSS } from '../utils/reportPrintCss';
import { ReportPrintDocument } from './ReportPrintDocument';

export interface IReportPrintPreviewModalProps {
  doc: ReportDoc;
  /** Thư mục ảnh của báo cáo (server-relative): data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/photos */
  photosFolder: string;
  /** Thư mục logo (server-relative): <thư viện>/resource, chứa thaco_logo.jpg và crec_logo.jpg. */
  resourceFolder: string;
  onClose: () => void;
}

/** URL tuyệt đối để iframe in tải được bằng phiên SharePoint. */
const absoluteUrl = (path: string): string => new URL(path, window.location.href).href;

export const ReportPrintPreviewModal: React.FC<IReportPrintPreviewModalProps> = ({
  doc,
  photosFolder,
  resourceFolder,
  onClose
}) => {
  const printAreaRef = React.useRef<HTMLDivElement | null>(null);
  const m = doc.metadata;
  const logoSrc = React.useMemo(
    () => ({ thaco: absoluteUrl(`${resourceFolder}/thaco_logo.jpg`), crec: absoluteUrl(`${resourceFolder}/crec_logo.jpg`) }),
    [resourceFolder]
  );
  const photoSrc = React.useCallback(
    (groupId: string, name: string): string => absoluteUrl(`${photosFolder}/${groupId}/${name}`),
    [photosFolder]
  );

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
  <style>${PRINT_CSS}</style>
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

        {/* Khung xem trước giấy A4 */}
        <div className="bcn-preview-scroll-body">
          <div ref={printAreaRef} className="bcn-print-container">
            <ReportPrintDocument doc={doc} logoSrc={logoSrc} photoSrc={photoSrc} />
          </div>
        </div>
      </div>
    </div>
  );
};
