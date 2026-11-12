import * as React from 'react';
import { ReportPhotoSection } from '../types/report.types';

export interface IReportPrintPhotoProps {
  section: ReportPhotoSection;
  roman: string;
  /** Thư mục ảnh của báo cáo (server-relative). */
  photosFolder: string;
}

/** Mục ảnh hiện trường trong bản in A4: mỗi bộ phận một dải tiêu đề, ảnh theo hàng, rồi dòng Tóm tắt. */
export const ReportPrintPhoto: React.FC<IReportPrintPhotoProps> = ({ section, roman, photosFolder }) => {
  const cols = section.columns;
  // URL tuyệt đối để iframe in tải được ảnh (cùng phiên đăng nhập SharePoint)
  const urlOf = (groupId: string, name: string): string =>
    new URL(`${photosFolder}/${groupId}/${name}`, window.location.href).href;

  const chunk = <T,>(arr: T[], n: number): T[][] => {
    const out: T[][] = [];
    for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
    return out;
  };

  return (
    <>
      <div className="bcn-sec-title">
        {roman}. {section.title} {section.titleEn && <span>| {section.titleEn}</span>}
      </div>
      <table className="bcn-table bcn-table-photos">
        <tbody>
          {section.groups.length === 0 && (
            <tr>
              <td colSpan={cols} className="text-center">
                Không có ảnh | No photos
              </td>
            </tr>
          )}
          {section.groups.map((g) => (
            <React.Fragment key={g.id}>
              <tr className="bcn-photo-dept">
                <td colSpan={cols}>BỘ PHẬN {g.department.toUpperCase()}</td>
              </tr>
              {chunk(g.files, cols).map((row, ri) => (
                <tr key={`${g.id}-${ri}`} className="bcn-photo-row">
                  {row.map((name) => (
                    <td key={name} className="bcn-photo-cell" style={{ width: `${100 / cols}%` }}>
                      <figure className="bcn-photo-item">
                        <img src={urlOf(g.id, name)} alt={name} />
                        <figcaption className="bcn-photo-cap">{name}</figcaption>
                      </figure>
                    </td>
                  ))}
                  {Array.from({ length: cols - row.length }).map((_, k) => (
                    <td key={`pad-${k}`} className="bcn-photo-cell" style={{ width: `${100 / cols}%` }} />
                  ))}
                </tr>
              ))}
              {g.summary && (
                <tr>
                  <td colSpan={cols}>
                    <strong>Tóm tắt | Summary:</strong> <span style={{ whiteSpace: 'pre-line' }}>{g.summary}</span>
                  </td>
                </tr>
              )}
            </React.Fragment>
          ))}
          {section.showSummary && section.summary && (
            <tr>
              <td colSpan={cols}>
                <strong>Tóm tắt | Summary:</strong>{' '}
                <span style={{ whiteSpace: 'pre-line' }}>{section.summary}</span>
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </>
  );
};
