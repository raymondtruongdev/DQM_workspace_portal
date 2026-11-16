import * as React from 'react';
import { ReportPhotoSection } from '../types/report.types';

export interface IReportPrintLayoutProps {
  section: ReportPhotoSection;
  roman: string;
  /** Địa chỉ ảnh của một hàng (URL tuyệt đối hoặc data URI). */
  photoSrc: (groupId: string, name: string) => string;
}

/** Mục SƠ ĐỒ LAYOUT trong bản in A4: bảng Tên sơ đồ | Hình vẽ | Mô tả, mỗi hàng một sơ đồ. */
export const ReportPrintLayout: React.FC<IReportPrintLayoutProps> = ({ section, roman, photoSrc }) => (
  <>
    <div className="bcn-sec-title">
      {roman}. {section.title} {section.titleEn && <span>| {section.titleEn}</span>}
    </div>
    <table className="bcn-table bcn-layout-table">
      <colgroup>
        <col style={{ width: '15%' }} />
        <col style={{ width: '55%' }} />
        <col style={{ width: '30%' }} />
      </colgroup>
      <thead>
        <tr className="bcn-th-row">
          <th>
            Tên sơ đồ <span>Layout name</span>
          </th>
          <th>
            Hình vẽ <span>Layout</span>
          </th>
          <th>
            Mô tả <span>Description</span>
          </th>
        </tr>
      </thead>
      <tbody>
        {section.groups.length === 0 && (
          <tr>
            <td colSpan={3} className="text-center">
              Không có sơ đồ | No layout
            </td>
          </tr>
        )}
        {section.groups.map((g, i) => (
          <tr key={g.id}>
            <td className="bcn-layout-name">
              <strong>{i + 1}.</strong> {g.department}
            </td>
            <td className="bcn-layout-img">
              {g.files.map((name) => (
                <img key={name} src={photoSrc(g.id, name)} alt={name} />
              ))}
            </td>
            <td className="bcn-layout-desc">{g.summary}</td>
          </tr>
        ))}
      </tbody>
    </table>
  </>
);
