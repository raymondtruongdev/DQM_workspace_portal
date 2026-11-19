import * as React from 'react';
import { ReportDoc, ReportSection } from '../types/report.types';
import { ReportCatalog } from '../utils/reportCatalog';
import { buildSectionNumbering, sectionAnchor } from '../utils/reportSections';
import { IReportDynamicTableProps, ReportDynamicTable } from './ReportDynamicTable';
import { ReportDynamicForm } from './ReportDynamicForm';
import { ReportPhotoSectionEditor } from './ReportPhotoSectionEditor';
import { ReportLayoutSectionEditor } from './ReportLayoutSectionEditor';

export interface IReportSectionsEditorProps {
  doc: ReportDoc;
  disabled?: boolean;
  isUploadingPhoto?: boolean;
  /** Danh mục Activity/cấu kiện cho các cột select:@… (chưa tải xong thì undefined). */
  catalog?: ReportCatalog;
  /** Thêm giá trị mới vào danh mục dùng chung, kèm hỏi liên kết (xem ReportDynamicTable). */
  onAddCatalogItem?: IReportDynamicTableProps['onAddCatalogItem'];
  /** Cập nhật liên kết của mục do người dùng thêm (xem ReportDynamicTable). */
  onUpdateCatalogLink?: IReportDynamicTableProps['onUpdateCatalogLink'];
  /** Đặt chỗ số mã tự tăng khi thêm hàng (xem ReportDynamicTable). */
  onReserveAutoId?: IReportDynamicTableProps['onReserveAutoId'];
  /** Thư mục PDF đính kèm sự kiện (server-relative) và các thao tác tải / xoá PDF của dòng (xem ReportDynamicTable). */
  eventPdfFolder?: string;
  onUploadRowPdf?: (sectionKey: string, rowId: string, files: FileList) => Promise<void>;
  onDeleteRowPdf?: (sectionKey: string, rowId: string, name: string) => Promise<void>;
  onPurgeRowPdf?: (sectionKey: string, rowId: string, names: string[]) => Promise<void>;
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
  onUpdateCatalogLink,
  onReserveAutoId,
  eventPdfFolder,
  onUploadRowPdf,
  onDeleteRowPdf,
  onPurgeRowPdf,
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
              onUpdateCatalogLink={onUpdateCatalogLink}
              rowPdfFolder={eventPdfFolder}
              onUploadRowPdf={onUploadRowPdf ? (rowId, files) => onUploadRowPdf(sec.key, rowId, files) : undefined}
              onDeleteRowPdf={onDeleteRowPdf ? (rowId, name) => onDeleteRowPdf(sec.key, rowId, name) : undefined}
              onPurgeRowPdf={onPurgeRowPdf ? (rowId, names) => onPurgeRowPdf(sec.key, rowId, names) : undefined}
              idCounters={doc.idCounters}
              onReserveAutoId={onReserveAutoId}
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
          const Editor = sec.variant === 'layout' ? ReportLayoutSectionEditor : ReportPhotoSectionEditor;
          return (
            <Editor
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
