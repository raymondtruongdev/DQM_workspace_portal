import * as React from 'react';
import { ReportDoc } from '../types/report.types';
import { buildSectionNumbering } from '../utils/reportSections';
import { formatTemplateVersionDisplay, getReportFolderName } from '../utils/reportFileNaming';
import { isoToVN } from '../utils/reportTable';
import { ReportPrintTable } from './ReportPrintTable';
import { ReportPrintForm } from './ReportPrintForm';
import { ReportPrintPhoto } from './ReportPrintPhoto';

export interface IReportPrintDocumentProps {
  doc: ReportDoc;
  /** Địa chỉ logo (URL hoặc data URI); thiếu hoặc không tải được thì hiện chữ THACO / CREC. */
  logoSrc: { thaco?: string; crec?: string };
  /** Địa chỉ ảnh hiện trường của một bộ phận. */
  photoSrc: (groupId: string, name: string) => string;
}

/** Logo ở đầu bản in; không có/không tải được (chưa có file) thì hiện chữ thay thế. */
const HeaderLogo: React.FC<{ src?: string; fallback: string }> = ({ src, fallback }) => {
  const [failed, setFailed] = React.useState<boolean>(false);
  if (!src || failed) return <>{fallback}</>;
  return <img className="bcn-logo-img" src={src} alt={fallback} onError={() => setFailed(true)} />;
};

/**
 * Một trang báo cáo A4 song ngữ (header THACO/CREC, bảng thông tin tài liệu, các mục theo thứ tự của file báo cáo, chân trang).
 * Dùng cho cả cửa sổ xem trước/in trên web và file HTML độc lập (render tĩnh), nên không dùng hook/DOM của trình duyệt.
 */
export const ReportPrintDocument: React.FC<IReportPrintDocumentProps> = ({ doc, logoSrc, photoSrc }) => {
  const m = doc.metadata;
  const num = buildSectionNumbering(doc);

  return (
    <div className="bcn-a4-page">
      {/* Top Title Banner Table */}
      <table className="bcn-table">
        <tbody>
          <tr>
            <td className="bcn-header-logo-cell">
              <HeaderLogo src={logoSrc.thaco} fallback="THACO" />
            </td>
            <td className="bcn-header-title-cell">
              <div className="bcn-title-vi">{m.docTitle || 'BÁO CÁO NGÀY'}</div>
              <div className="bcn-title-en">{m.docTitleEn || 'DAILY REPORT'}</div>
              {m.subTitle && <div className="bcn-subtitle">{m.subTitle}</div>}
            </td>
            <td className="bcn-header-logo-cell">
              <HeaderLogo src={logoSrc.crec} fallback="CREC" />
            </td>
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
            <td>{m.subcontractors.join(', ')}</td>
            <td className="bcn-meta-label" style={{ whiteSpace: 'nowrap' }}>
              Ngày <span>| Date</span>
            </td>
            <td style={{ width: '14%' }}>{isoToVN(doc.dateStr)}</td>
          </tr>
        </tbody>
      </table>

      {/* Các mục của báo cáo, theo thứ tự trong file */}
      {doc.sections.map((sec) =>
        sec.kind === 'table' ? (
          <ReportPrintTable key={sec.key} table={sec} roman={num.roman(sec.key)} dateStr={doc.dateStr} />
        ) : sec.kind === 'photo' ? (
          <ReportPrintPhoto key={sec.key} section={sec} roman={num.roman(sec.key)} photoSrc={photoSrc} />
        ) : (
          <ReportPrintForm key={sec.key} form={sec} roman={num.roman(sec.key)} dateStr={doc.dateStr} />
        )
      )}

      <div className="bcn-keep-together">
        <div className="bcn-footer-note">
          <div>Sinh tự động từ Microsoft 365 (SharePoint)</div>
          <div>
            Ảnh: {`data/${doc.unit}/${getReportFolderName(doc.docId, doc.unit, doc.dateStr)}/photos/`}
            {' | '}
            Template version: {formatTemplateVersionDisplay(m.templateVersion || doc.latestTemplateInfo?.version)}
          </div>
        </div>
      </div>
    </div>
  );
};
