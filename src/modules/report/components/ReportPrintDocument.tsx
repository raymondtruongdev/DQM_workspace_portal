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

/** Ô nhãn song ngữ: tiếng Việt đậm, tiếng Anh xám bên dưới (`inline` = cùng một dòng, vd "Ngày | Date:"). */
const MetaLabel: React.FC<{ vi: string; en: string; inline?: boolean }> = ({ vi, en, inline }) => (
  <td className={`bcn-meta-label bcn-meta-label--stack${inline ? ' bcn-meta-label--inline' : ''}`}>
    <b>{vi}</b>
    {inline ? <span> | {en}</span> : <span>{en}</span>}
  </td>
);

/** Mã tài liệu: mã đơn vị và số báo cáo tô đỏ (mã theo mẫu `...-{unit}-...-{reportNo}`). */
const DocCode: React.FC<{ code: string; unit: string; reportNo: string }> = ({ code, unit, reportNo }) => {
  const midToken = `-${unit}-`;
  const tailToken = `-${reportNo}`;
  const mid = unit ? code.indexOf(midToken) : -1;
  const tail = reportNo && code.endsWith(tailToken) ? code.length - tailToken.length : -1;
  if (mid === -1 || tail === -1 || tail < mid + midToken.length) return <>{code}</>;
  return (
    <>
      {code.slice(0, mid + 1)}
      <em>{unit}</em>
      {code.slice(mid + 1 + unit.length, tail + 1)}
      <em>{reportNo}</em>
    </>
  );
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

      {/* Bảng thông tin tài liệu: 6 cột (nhãn | giá trị | nhãn | giá trị | nhãn | giá trị) */}
      <table className="bcn-table bcn-meta-table">
        <colgroup>
          <col style={{ width: '14%' }} />
          <col style={{ width: '17%' }} />
          <col style={{ width: '16%' }} />
          <col style={{ width: '25%' }} />
          <col style={{ width: '14%' }} />
          <col style={{ width: '14%' }} />
        </colgroup>
        <tbody>
          <tr>
            <MetaLabel vi="Mã tài liệu" en="Doc. code" />
            <td className="bcn-meta-doccode" colSpan={3}>
              <DocCode code={m.docCode} unit={doc.unit} reportNo={m.reportNo} />
            </td>
            <MetaLabel vi="Số báo cáo" en="Report no.:" />
            <td className="bcn-meta-hot">{m.reportNo}</td>
          </tr>
          <tr>
            <MetaLabel vi="Vị trí" en="Location" />
            <td className="bcn-meta-hot bcn-meta-hot--xl">{m.location}</td>
            <MetaLabel vi="Tọa độ" en="Geo-tag" />
            <td className="bcn-meta-geo">{m.geoTag}</td>
            <MetaLabel vi="Ngày" en="Date:" inline />
            <td className="bcn-meta-hot">{isoToVN(doc.dateStr)}</td>
          </tr>
          <tr className="bcn-meta-tall">
            <MetaLabel vi="Đơn vị thi công" en="Subcontractor" />
            <td className="bcn-meta-subs">
              {m.subcontractors.map((x) => (
                <div key={x}>{x}</div>
              ))}
            </td>
            <MetaLabel vi="Tài liệu đính kèm" en="Attached Documents" />
            <td className="bcn-meta-docs" colSpan={3}>
              {m.attachedDocs}
            </td>
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
