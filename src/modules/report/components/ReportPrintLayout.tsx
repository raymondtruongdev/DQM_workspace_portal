import * as React from 'react';
import { ReportPhotoSection } from '../types/report.types';

export interface IReportPrintLayoutProps {
  section: ReportPhotoSection;
  roman: string;
  /** Địa chỉ ảnh của một hàng (URL tuyệt đối hoặc data URI). */
  photoSrc: (groupId: string, name: string) => string;
}

/** Mục SƠ ĐỒ LAYOUT trong bản in A4: mỗi sơ đồ gồm cột Tên sơ đồ (15%) và cột 85% có Hình vẽ ở trên, Mô tả ở dưới. */
export const ReportPrintLayout: React.FC<IReportPrintLayoutProps> = ({ section, roman, photoSrc }) => (
  <>
    <div className="bcn-sec-title">
      {roman}. {section.title} {section.titleEn && <span>| {section.titleEn}</span>}
    </div>
    <table className="bcn-table bcn-layout-table">
      <colgroup>
        <col style={{ width: '15%' }} />
        <col style={{ width: '85%' }} />
      </colgroup>
      {section.groups.length === 0 && (
        <tbody>
          <tr>
            <td colSpan={2} className="text-center">
              Không có sơ đồ | No layout
            </td>
          </tr>
        </tbody>
      )}
      {section.groups.map((g, i) => (
        <tbody key={g.id} className="bcn-layout-group">
          <tr>
            <td rowSpan={2} className="bcn-layout-name">
              <strong>{i + 1}.</strong> {g.department}
            </td>
            <td className="bcn-layout-img">
              {g.files.map((name) => (
                <img key={name} src={photoSrc(g.id, name)} alt={name} />
              ))}
            </td>
          </tr>
          <tr>
            <td className="bcn-layout-desc">{g.summary}</td>
          </tr>
        </tbody>
      ))}
    </table>
  </>
);
