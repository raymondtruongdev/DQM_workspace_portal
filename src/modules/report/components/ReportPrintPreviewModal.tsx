import * as React from 'react';
import { ReportDoc } from '../types/report.types';
import { buildSectionNumbering, hasSection } from '../utils/reportSections';
import { formatTemplateVersionDisplay } from '../utils/reportFileNaming';

export interface IReportPrintPreviewModalProps {
  doc: ReportDoc;
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
  onClose
}) => {
  const printAreaRef = React.useRef<HTMLDivElement | null>(null);
  const m = doc.metadata;
  const num = buildSectionNumbering(doc);

  const totalNos = doc.manpower.reduce((acc, r) => acc + (r.nos || 0), 0);
  const totalManHrs = doc.manpower.reduce((acc, r) => acc + (r.manHrs || 0), 0);
  // URL tuyệt đối để iframe in tải được ảnh (cùng phiên đăng nhập SharePoint)
  const photoItems = (doc.photos || []).map((ph) => ({
    name: ph.name,
    url: new URL(ph.serverRelativeUrl, window.location.href).href
  }));
  const photosCount = doc.photos && doc.photos.length > 0 ? doc.photos.length : m.photosCount ?? 0;

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

    // Tên file PDF mặc định: BCN_<unit>_<yyyy-MM-dd>
    const pdfFileName = `BCN_${doc.unit}_${doc.dateStr}`.replace(/[\\/:*?"<>|\s]+/g, '_');

    iframeDoc.open();
    iframeDoc.write(`<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${pdfFileName}</title>
  <style>
    @page {
      size: A4 portrait;
      margin: 8mm 10mm;
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
      font-weight: 400;
      color: #4b5563;
    }
    .bcn-total-row td {
      background: #f3f4f6;
      font-weight: 700;
    }
    .text-center { text-align: center; }
    .text-right { text-align: right; }
    .font-bold { font-weight: 700; }
    .sev-cao { color: #b91c1c; font-weight: 700; }
    .sev-tb { color: #9a3412; font-weight: 700; }
    .bcn-preline { white-space: pre-line; }
    .bcn-sign-box {
      height: 110px; /* chừa chỗ trống phía trên tên để ký */
      vertical-align: bottom !important;
      text-align: center;
      padding-bottom: 8px;
    }
    .bcn-table tr.bcn-photo-row {
      page-break-inside: auto;
      break-inside: auto;
    }
    .bcn-photo-grid {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 5px;
    }
    .bcn-photo-item {
      width: calc(50% - 4px);
      margin: 0;
      border: 1px solid #9ca3af;
      page-break-inside: avoid;
      break-inside: avoid;
    }
    .bcn-photo-item img {
      display: block;
      width: 100%;
      height: 62mm;
      object-fit: contain;
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
                      <div className="bcn-subtitle">{m.subTitle}</div>
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
                    <td style={{ width: '48%' }}>{m.location}</td>
                    <td className="bcn-meta-label" style={{ width: '17%' }}>
                      Số BCN <span>| Report no.</span>
                    </td>
                    <td style={{ width: '18%' }}>{m.reportNo}</td>
                  </tr>
                  <tr>
                    <td className="bcn-meta-label">
                      Ngày <span>| Date</span>
                    </td>
                    <td>{formatDateVN(doc.dateStr)}</td>
                    <td className="bcn-meta-label">
                      Ca <span>| Shift</span>
                    </td>
                    <td>{m.shift}</td>
                  </tr>
                  <tr>
                    <td className="bcn-meta-label">
                      Phiên bản mẫu <span>| Template</span>
                    </td>
                    <td colSpan={3}>
                      {m.templateVersion || doc.latestTemplateInfo?.version || `default_${doc.unit}`}
                    </td>
                  </tr>
                </tbody>
              </table>

              {/* I. THÔNG TIN CHUNG & THỜI TIẾT */}
              <div className="bcn-sec-title">
                I. THÔNG TIN CHUNG &amp; THỜI TIẾT <span>| GENERAL &amp; WEATHER</span>
              </div>
              <table className="bcn-table">
                <thead>
                  <tr className="bcn-th-row">
                    <th style={{ width: '16%' }}>
                      Sáng <span>| AM</span>
                    </th>
                    <th style={{ width: '16%' }}>
                      Chiều <span>| PM</span>
                    </th>
                    <th style={{ width: '22%' }}>
                      Nhiệt độ max <span>| Max temp</span>
                    </th>
                    <th style={{ width: '24%' }}>
                      Giờ mất do thời tiết <span>| Hrs lost</span>
                    </th>
                    <th style={{ width: '22%' }}>
                      Đánh giá tiến độ <span>| Progress</span>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td className="text-center">{m.weatherAM}</td>
                    <td className="text-center">{m.weatherPM}</td>
                    <td className="text-center">
                      {m.maxTemp ? (m.maxTemp.includes('°C') ? m.maxTemp : `${m.maxTemp} °C`) : ''}
                    </td>
                    <td className="text-center">
                      {m.hrsLost !== undefined ? String(m.hrsLost).replace('.', ',') : '0'}
                    </td>
                    <td className="text-center font-bold">{m.progressEval}</td>
                  </tr>
                  <tr>
                    <td colSpan={5}>
                      <strong>Tóm tắt | Summary:</strong> {m.summary}
                    </td>
                  </tr>
                </tbody>
              </table>

              {/* II. NHÂN LỰC | MANPOWER */}
              {hasSection(doc, 'manpower') && (
                <>
                  <div className="bcn-sec-title">
                    {num.roman('manpower')}. NHÂN LỰC <span>| MANPOWER</span>
                  </div>
                  <table className="bcn-table">
                    <thead>
                      <tr className="bcn-th-row">
                        <th style={{ width: '20%' }}>
                          Đơn vị <span>| Org.</span>
                        </th>
                        <th style={{ width: '38%' }}>
                          Loại <span>| Category</span>
                        </th>
                        <th style={{ width: '14%' }}>
                          Số người <span>| Nos.</span>
                        </th>
                        <th style={{ width: '14%' }}>
                          Giờ/người <span>| Hrs</span>
                        </th>
                        <th style={{ width: '14%' }}>
                          Giờ công <span>| Man-hrs</span>
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {doc.manpower.map((r, i) => (
                        <tr key={r.id || i}>
                          <td>{r.org}</td>
                          <td>{r.category}</td>
                          <td className="text-right">{r.nos ?? 0}</td>
                          <td className="text-right">{r.hrs ?? 0}</td>
                          <td className="text-right">{(r.manHrs || 0).toLocaleString('vi-VN')}</td>
                        </tr>
                      ))}
                      <tr className="bcn-total-row">
                        <td colSpan={2}>
                          Tổng <span style={{ fontWeight: 400, color: '#4b5563' }}>| Total</span>
                        </td>
                        <td className="text-right">{totalNos.toLocaleString('vi-VN')}</td>
                        <td />
                        <td className="text-right">{totalManHrs.toLocaleString('vi-VN')}</td>
                      </tr>
                      {m.manpowerSummary && (
                        <tr>
                          <td colSpan={5}>
                            <strong>Tóm tắt | Summary:</strong> {m.manpowerSummary}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </>
              )}

              {/* III. THIẾT BỊ | EQUIPMENT */}
              {hasSection(doc, 'equipment') && (
                <>
                  <div className="bcn-sec-title">
                    {num.roman('equipment')}. THIẾT BỊ <span>| EQUIPMENT</span>
                  </div>
                  <table className="bcn-table">
                    <thead>
                      <tr className="bcn-th-row">
                        <th style={{ width: '36%' }}>
                          Thiết bị <span>| Equipment</span>
                        </th>
                        <th style={{ width: '10%' }}>
                          SL <span>| Qty</span>
                        </th>
                        <th style={{ width: '12%' }}>
                          Giờ HĐ <span>| Hrs</span>
                        </th>
                        <th style={{ width: '16%' }}>
                          Tình trạng <span>| Status</span>
                        </th>
                        <th style={{ width: '26%' }}>
                          Ghi chú <span>| Remark</span>
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {doc.equipment.map((r, i) => (
                        <tr key={r.id || i}>
                          <td>{r.equipment}</td>
                          <td className="text-right">{r.qty ?? 0}</td>
                          <td className="text-right">
                            {r.hrs !== undefined ? String(r.hrs).replace('.', ',') : '0'}
                          </td>
                          <td className="text-center">{r.status}</td>
                          <td>{r.remark}</td>
                        </tr>
                      ))}
                      {m.equipmentSummary && (
                        <tr>
                          <td colSpan={5}>
                            <strong>Tóm tắt | Summary:</strong> {m.equipmentSummary}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </>
              )}

              {/* IV. KHỐI LƯỢNG THỰC HIỆN | WORK QUANTITIES */}
              {hasSection(doc, 'workqty') && (
                <>
                  <div className="bcn-sec-title">
                    {num.roman('workqty')}. KHỐI LƯỢNG THỰC HIỆN <span>| WORK QUANTITIES</span>
                  </div>
                  <table className="bcn-table">
                    <thead>
                      <tr className="bcn-th-row">
                        <th style={{ width: '16%' }}>Activity ID</th>
                        <th style={{ width: '30%' }}>
                          Hạng mục <span>| Item</span>
                        </th>
                        <th style={{ width: '8%' }}>
                          ĐV <span>| Unit</span>
                        </th>
                        <th style={{ width: '10%' }}>
                          KH ngày <span>| Plan</span>
                        </th>
                        <th style={{ width: '10%' }}>
                          Thực hiện <span>| Actual</span>
                        </th>
                        <th style={{ width: '8%' }}>%</th>
                        <th style={{ width: '18%' }}>
                          Ghi chú <span>| Remark</span>
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {doc.workItems.map((r, i) => (
                        <tr key={r.id || i}>
                          <td>{r.activityId}</td>
                          <td>{r.item}</td>
                          <td className="text-center">{r.unit}</td>
                          <td className="text-right">{r.plan ?? 0}</td>
                          <td className="text-right">{r.actual ?? 0}</td>
                          <td className="text-right">{r.percent !== undefined ? `${r.percent}%` : ''}</td>
                          <td>{r.remark}</td>
                        </tr>
                      ))}
                      {m.workQtySummary && (
                        <tr>
                          <td colSpan={7}>
                            <strong>Tóm tắt | Summary:</strong> {m.workQtySummary}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </>
              )}

              {/* V. VẬT TƯ | MATERIALS */}
              {hasSection(doc, 'materials') && (
                <>
                  <div className="bcn-sec-title">
                    {num.roman('materials')}. VẬT TƯ <span>| MATERIALS</span>
                  </div>
                  <table className="bcn-table">
                    <thead>
                      <tr className="bcn-th-row">
                        <th style={{ width: '38%' }}>
                          Vật tư <span>| Material</span>
                        </th>
                        <th style={{ width: '10%' }}>
                          ĐV <span>| Unit</span>
                        </th>
                        <th style={{ width: '13%' }}>
                          Nhập về <span>| Recv.</span>
                        </th>
                        <th style={{ width: '13%' }}>
                          Sử dụng <span>| Used</span>
                        </th>
                        <th style={{ width: '26%' }}>Phiếu MIR/MAR</th>
                      </tr>
                    </thead>
                    <tbody>
                      {doc.materials.map((r, i) => (
                        <tr key={r.id || i}>
                          <td>{r.material}</td>
                          <td className="text-center">{r.unit}</td>
                          <td className="text-right">
                            {r.recv !== undefined ? String(r.recv).replace('.', ',') : '0'}
                          </td>
                          <td className="text-right">
                            {r.used !== undefined ? String(r.used).replace('.', ',') : '0'}
                          </td>
                          <td>{r.mirMar}</td>
                        </tr>
                      ))}
                      {m.materialsSummary && (
                        <tr>
                          <td colSpan={5}>
                            <strong>Tóm tắt | Summary:</strong> {m.materialsSummary}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </>
              )}

              {/* VI. CÔNG VIỆC THỰC HIỆN & KẾ HOẠCH NGÀY MAI */}
              <div className="bcn-sec-title">
                {num.roman('worksplan')}. CÔNG VIỆC THỰC HIỆN &amp; KẾ HOẠCH NGÀY MAI{' '}
                <span>| WORKS TODAY &amp; PLAN FOR TOMORROW</span>
              </div>
              <table className="bcn-table">
                <thead>
                  <tr className="bcn-th-row">
                    <th style={{ width: '50%' }}>
                      Hôm nay <span>| Today</span>
                    </th>
                    <th style={{ width: '50%' }}>
                      Ngày mai <span>| Tomorrow</span>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td className="bcn-preline" style={{ verticalAlign: 'top' }}>
                      {m.worksToday}
                    </td>
                    <td className="bcn-preline" style={{ verticalAlign: 'top' }}>
                      {m.planTomorrow}
                    </td>
                  </tr>
                  {m.worksPlanSummary && (
                    <tr>
                      <td colSpan={2}>
                        <strong>Tóm tắt | Summary:</strong> {m.worksPlanSummary}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>

              {/* VII. VƯỚNG MẮC / RÀNG BUỘC ĐANG MỞ */}
              {hasSection(doc, 'issues') && (
                <>
                  <div className="bcn-sec-title">
                    {num.roman('issues')}. VƯỚNG MẮC / RÀNG BUỘC ĐANG MỞ <span>| OPEN ISSUES &amp; CONSTRAINTS</span>
                  </div>
                  <table className="bcn-table">
                    <thead>
                      <tr className="bcn-th-row">
                        <th style={{ width: '13%' }}>
                          Mã <span>| ID</span>
                        </th>
                        <th style={{ width: '38%' }}>
                          Nội dung <span>| Description</span>
                        </th>
                        <th style={{ width: '10%' }}>
                          Mức <span>| Sev.</span>
                        </th>
                        <th style={{ width: '8%' }}>
                          Găng <span>| CP</span>
                        </th>
                        <th style={{ width: '15%' }}>
                          Bên xử lý <span>| By</span>
                        </th>
                        <th style={{ width: '8%' }}>
                          Hạn <span>| Due</span>
                        </th>
                        <th style={{ width: '8%' }}>Trạng thái</th>
                      </tr>
                    </thead>
                    <tbody>
                      {doc.issues.map((r, i) => (
                        <tr key={r.id || i}>
                          <td>{r.issueId}</td>
                          <td>{r.description}</td>
                          <td
                            className={`text-center ${
                              r.severity === 'Cao' ? 'sev-cao' : r.severity === 'Trung bình' ? 'sev-tb' : ''
                            }`}
                          >
                            {r.severity}
                          </td>
                          <td className="text-center font-bold">{r.criticalPath}</td>
                          <td>{r.owner}</td>
                          <td className="text-center">{r.dueDate}</td>
                          <td className="text-center">{r.status}</td>
                        </tr>
                      ))}
                      {m.issuesSummary && (
                        <tr>
                          <td colSpan={7}>
                            <strong>Tóm tắt | Summary:</strong> {m.issuesSummary}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </>
              )}

              {/* VIII. AN TOÀN – CHẤT LƯỢNG – CHỈ THỊ TVGS */}
              <div className="bcn-sec-title">
                {num.roman('hse')}. AN TOÀN – CHẤT LƯỢNG – CHỈ THỊ TVGS{' '}
                <span>| HSE – QUALITY – ENGINEER&apos;S INSTRUCTIONS</span>
              </div>
              <table className="bcn-table">
                <thead>
                  <tr className="bcn-th-row">
                    <th style={{ width: '18%' }}>
                      Sự cố HSE <span>| Incidents</span>
                    </th>
                    <th style={{ width: '24%' }}>
                      Kiểm tra ATLĐ <span>| HSE inspections</span>
                    </th>
                    <th style={{ width: '19%' }}>
                      Nghiệm thu <span>| IR</span>
                    </th>
                    <th style={{ width: '19%' }}>
                      NCR mới <span>| New NCR</span>
                    </th>
                    <th style={{ width: '20%' }}>
                      Ảnh hiện trường <span>| Photos</span>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td className="text-center">{m.hseIncidents ?? 0}</td>
                    <td className="text-center">{m.hseInspections ?? 0}</td>
                    <td className="text-center">{m.inspectionRequests ?? 0}</td>
                    <td className="text-center">{m.newNCR ?? 0}</td>
                    <td className="text-center">{photosCount} ảnh</td>
                  </tr>
                  <tr>
                    <td colSpan={5}>
                      <strong>Mô tả sự cố | Incident details:</strong> {m.incidentDetails || 'Không có | None'}
                    </td>
                  </tr>
                  <tr>
                    <td colSpan={5}>
                      <strong>Chỉ thị/ý kiến TVGS | Engineer&apos;s instructions:</strong>{' '}
                      {m.engineerInstructions}
                    </td>
                  </tr>
                  <tr>
                    <td colSpan={5}>
                      <strong>Kế hoạch ngày mai (tóm tắt) | Plan summary:</strong> {m.planSummary}
                    </td>
                  </tr>
                  {photoItems.length > 0 && (
                    <tr className="bcn-photo-row">
                      <td colSpan={5}>
                        <strong>Ảnh hiện trường | Site photos ({photoItems.length}):</strong>
                        <div className="bcn-photo-grid">
                          {photoItems.map((ph) => (
                            <figure key={ph.url} className="bcn-photo-item">
                              <img src={ph.url} alt={ph.name} />
                              <figcaption className="bcn-photo-cap">{ph.name}</figcaption>
                            </figure>
                          ))}
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>

              {/* SIGNATURES TABLE */}
              <div className="bcn-keep-together">
                <table className="bcn-table" style={{ marginTop: '8px' }}>
                  <thead>
                    <tr className="bcn-th-row">
                      <th style={{ width: '25%' }}>
                        Người lập
                        <br />
                        <span>Prepared by</span>
                      </th>
                      <th style={{ width: '25%' }}>
                        Trưởng ga / CHT
                        <br />
                        <span>Site manager</span>
                      </th>
                      <th style={{ width: '25%' }}>
                        Đại diện TCJV
                        <br />
                        <span>TCJV representative</span>
                      </th>
                      <th style={{ width: '25%' }}>
                        TVGS xác nhận
                        <br />
                        <span>Engineer acknowledged</span>
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="bcn-sign-box">
                        <div style={{ fontWeight: 600 }}>
                          {m.preparedBy || m.email || '[Người lập]'}
                        </div>
                      </td>
                      <td className="bcn-sign-box">
                        <div style={{ fontWeight: 600 }}>
                          {m.siteManager || '[Trưởng ga]'}
                        </div>
                      </td>
                      <td className="bcn-sign-box">
                        <div style={{ fontWeight: 600 }}>{m.tcjvRep || ''}</div>
                      </td>
                      <td className="bcn-sign-box">
                        <div style={{ fontWeight: 600 }}>{m.engineerAck || ''}</div>
                      </td>
                    </tr>
                  </tbody>
                </table>

                {/* FOOTER NOTE */}
                <div className="bcn-footer-note">
                  {m.approvalComment && (
                    <div>
                      <strong>Ý kiến duyệt:</strong> {m.approvalComment}
                    </div>
                  )}
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
