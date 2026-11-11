import * as React from 'react';
import { ReportDoc, ReportSection } from '../types/report.types';
import { buildSectionNumbering, sectionAnchor } from '../utils/reportSections';
import { ReportDynamicTable } from './ReportDynamicTable';
import { ReportDynamicForm } from './ReportDynamicForm';
import { ReportPhotoManager } from './ReportPhotoManager';

export interface IReportSectionsEditorProps {
  doc: ReportDoc;
  disabled?: boolean;
  isUploadingPhoto?: boolean;
  /** Tô đỏ các trường bắt buộc đang trống. */
  showErrors?: boolean;
  onSectionChange: (section: ReportSection) => void;
  onUploadPhotos: (files: FileList) => Promise<void>;
  onDeletePhoto: (serverRelativeUrl: string, name: string) => Promise<void>;
}

/** Các mục của báo cáo: mỗi sheet của template = một mục (bảng hoặc form), theo thứ tự sheet. */
export const ReportSectionsEditor: React.FC<IReportSectionsEditorProps> = ({
  doc,
  disabled = false,
  isUploadingPhoto = false,
  showErrors = false,
  onSectionChange,
  onUploadPhotos,
  onDeletePhoto
}) => {
  const num = buildSectionNumbering(doc);

  return (
    <div className="report-sections-stack">
      {doc.sections.map((sec) =>
        sec.kind === 'table' ? (
          <ReportDynamicTable
            key={sec.key}
            table={sec}
            unit={doc.unit}
            roman={num.roman(sec.key)}
            anchor={sectionAnchor(sec.key)}
            disabled={disabled}
            onChange={onSectionChange}
          />
        ) : (
          <ReportDynamicForm
            key={sec.key}
            form={sec}
            roman={num.roman(sec.key)}
            anchor={sectionAnchor(sec.key)}
            disabled={disabled}
            showErrors={showErrors}
            renderPhotos={() => (
              <ReportPhotoManager
                unit={doc.unit}
                docId={doc.docId}
                dateStr={doc.dateStr}
                photos={doc.photos || []}
                disabled={disabled}
                isUploading={isUploadingPhoto}
                onUpload={onUploadPhotos}
                onDelete={onDeletePhoto}
              />
            )}
            onChange={onSectionChange}
          />
        )
      )}
    </div>
  );
};
