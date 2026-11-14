import * as React from 'react';
import { ReportDoc, ReportSection } from '../types/report.types';
import { ReportCatalog } from '../utils/reportCatalog';
import { buildSectionNumbering, sectionAnchor } from '../utils/reportSections';
import { IReportDynamicTableProps, ReportDynamicTable } from './ReportDynamicTable';
import { ReportDynamicForm } from './ReportDynamicForm';
import { ReportPhotoSectionEditor } from './ReportPhotoSectionEditor';

export interface IReportSectionsEditorProps {
  doc: ReportDoc;
  disabled?: boolean;
  isUploadingPhoto?: boolean;
  /** Danh mục Activity/cấu kiện cho các cột select:@… (chưa tải xong thì undefined). */
  catalog?: ReportCatalog;
  /** Thêm giá trị mới vào danh mục dùng chung, kèm hỏi liên kết (xem ReportDynamicTable). */
  onAddCatalogItem?: IReportDynamicTableProps['onAddCatalogItem'];
  /** Tô đỏ các trường bắt buộc đang trống. */
  showErrors?: boolean;
  onSectionChange: (section: ReportSection) => void;
  /** Thư mục ảnh của báo cáo (server-relative). */
  photosFolder: string;
  onUploadGroupPhotos: (sectionKey: string, groupId: string, files: FileList) => Promise<void>;
  onDeleteGroupPhoto: (sectionKey: string, groupId: string, name: string) => Promise<void>;
  onDeletePhotoGroup: (sectionKey: string, groupId: string) => Promise<void>;
}

/** Các mục của báo cáo theo thứ tự trong file (bảng, ảnh hoặc các trường văn bản). */
export const ReportSectionsEditor: React.FC<IReportSectionsEditorProps> = ({
  doc,
  disabled = false,
  isUploadingPhoto = false,
  catalog,
  onAddCatalogItem,
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
              onAddCatalogItem={onAddCatalogItem}
              idCounters={doc.idCounters}
              dateStr={doc.dateStr}
              roman={roman}
              anchor={anchor}
              disabled={disabled}
              showErrors={showErrors}
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
