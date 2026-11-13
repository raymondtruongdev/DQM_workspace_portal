import * as React from 'react';
import { ReportDoc, ReportSection } from '../types/report.types';
import { ReportCatalog } from '../utils/reportCatalog';
import { buildSectionNumbering, sectionAnchor } from '../utils/reportSections';
import { ReportDynamicTable } from './ReportDynamicTable';
import { ReportDynamicForm } from './ReportDynamicForm';
import { ReportPhotoSectionEditor } from './ReportPhotoSectionEditor';

export interface IReportSectionsEditorProps {
  doc: ReportDoc;
  disabled?: boolean;
  isUploadingPhoto?: boolean;
  /** Danh mục Activity/cấu kiện cho các cột select:@… (chưa tải xong thì undefined). */
  catalog?: ReportCatalog;
  /** Tô đỏ các trường bắt buộc đang trống. */
  showErrors?: boolean;
  onSectionChange: (section: ReportSection) => void;
  /** Thư mục ảnh của báo cáo (server-relative). */
  photosFolder: string;
  onUploadGroupPhotos: (sectionKey: string, groupId: string, files: FileList) => Promise<void>;
  onDeleteGroupPhoto: (sectionKey: string, groupId: string, name: string) => Promise<void>;
  onDeletePhotoGroup: (sectionKey: string, groupId: string) => Promise<void>;
}

/** Các mục của báo cáo: mỗi sheet của template = một mục (bảng hoặc form), theo thứ tự sheet. */
export const ReportSectionsEditor: React.FC<IReportSectionsEditorProps> = ({
  doc,
  disabled = false,
  isUploadingPhoto = false,
  catalog,
  showErrors = false,
  onSectionChange,
  photosFolder,
  onUploadGroupPhotos,
  onDeleteGroupPhoto,
  onDeletePhotoGroup
}) => {
  const num = buildSectionNumbering(doc);

  return (
    <div className="report-sections-stack">
      {doc.sections.map((sec) => {
        const roman = num.roman(sec.key);
        const anchor = sectionAnchor(sec.key);
        if (sec.kind === 'table') {
          return (
            <ReportDynamicTable
              key={sec.key}
              table={sec}
              unit={doc.unit}
              catalog={catalog}
              dateStr={doc.dateStr}
              roman={roman}
              anchor={anchor}
              disabled={disabled}
              onChange={onSectionChange}
            />
          );
        }
        if (sec.kind === 'photo') {
          return (
            <ReportPhotoSectionEditor
              key={sec.key}
              section={sec}
              roman={roman}
              anchor={anchor}
              photosFolder={photosFolder}
              disabled={disabled}
              isUploading={isUploadingPhoto}
              onChange={onSectionChange}
              onUpload={(groupId, files) => onUploadGroupPhotos(sec.key, groupId, files)}
              onDeletePhoto={(groupId, name) => onDeleteGroupPhoto(sec.key, groupId, name)}
              onDeleteGroup={(groupId) => onDeletePhotoGroup(sec.key, groupId)}
            />
          );
        }
        return (
          <ReportDynamicForm
            key={sec.key}
            form={sec}
            dateStr={doc.dateStr}
            roman={roman}
            anchor={anchor}
            disabled={disabled}
            showErrors={showErrors}
            onChange={onSectionChange}
          />
        );
      })}
    </div>
  );
};
