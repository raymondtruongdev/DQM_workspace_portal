import * as React from 'react';
import { ReportReadOnlyContext } from './ReportReadOnlyContext';
import { PhotoGroup, ReportPhotoSection } from '../types/report.types';
import { Icon, EmptyState } from '../../../core/components';
import { ReportPhotoManager } from './ReportPhotoManager';
import { LAYOUT_GROUP_PREFIX, photoSubfolder } from '../utils/reportFileNaming';

export interface IReportLayoutSectionEditorProps {
  section: ReportPhotoSection;
  roman: string;
  anchor: string;
  /** Thư mục ảnh của báo cáo (server-relative): data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/photos */
  photosFolder: string;
  disabled?: boolean;
  isUploading?: boolean;
  onChange: (section: ReportPhotoSection) => void;
  onUpload: (groupId: string, files: FileList) => Promise<void>;
  onDeletePhoto: (groupId: string, name: string) => Promise<void>;
  onDeleteGroup: (groupId: string) => Promise<void>;
}

/**
 * Mục SƠ ĐỒ LAYOUT: mỗi hàng gồm cột Tên sơ đồ (15%, nhập) và cột 85% có Hình vẽ (tải ảnh lên, lưu ở photos/layout/ giữ nguyên tên file) ở trên, Mô tả (nhập) ở dưới.
 * Dùng chung dữ liệu và các hàm tải/xoá ảnh của mục ảnh: Tên sơ đồ = `department`, Mô tả = `summary`, Hình vẽ = `files`.
 */
export const ReportLayoutSectionEditor: React.FC<IReportLayoutSectionEditorProps> = ({
  section,
  roman,
  anchor,
  photosFolder,
  disabled = false,
  isUploading = false,
  onChange,
  onUpload,
  onDeletePhoto,
  onDeleteGroup
}) => {
  const readOnlyMode = React.useContext(ReportReadOnlyContext);
  const updateRow = (id: string, patch: Partial<PhotoGroup>): void =>
    onChange({ ...section, groups: section.groups.map((g) => (g.id === id ? { ...g, ...patch } : g)) });

  // id = `layout-<số>`: ảnh của mọi hàng nằm chung trong photos/layout/ (giữ nguyên tên file)
  const addRow = (): void =>
    onChange({
      ...section,
      groups: [...section.groups, { id: `${LAYOUT_GROUP_PREFIX}${Date.now()}`, department: '', summary: '', files: [] }]
    });

  return (
    <div id={anchor} className="section-card">
      <div className="section-card__banner">
        <span className="section-card__title">
          {roman}. {section.title}{' '}
          {section.titleEn && <span className="section-card__title-sub">| {section.titleEn}</span>}
        </span>
        <button type="button" className="btn btn--light btn--sm" onClick={addRow} disabled={disabled || isUploading}>
          + Thêm hàng
        </button>
      </div>

      <div className="section-card__body">
        {section.groups.length === 0 ? (
          <EmptyState
            icon="image"
            title="Chưa có sơ đồ nào"
            message="Bấm “Thêm hàng” để nhập tên sơ đồ, tải hình vẽ lên và nhập mô tả."
          />
        ) : (
          <div className="report-layout-list">
            {section.groups.map((g) => (
              <div key={g.id} className="report-layout-row">
                <div className="field">
                  <label className="field__label report-layout-label">Tên sơ đồ | Layout name</label>
                  <textarea
                    className="textarea"
                    rows={3}
                    placeholder="Nhập tên sơ đồ..."
                    aria-label="Tên sơ đồ"
                    value={g.department}
                    onChange={(e) => updateRow(g.id, { department: e.target.value })}
                    disabled={disabled && !readOnlyMode}
                    readOnly={readOnlyMode}
                  />
                </div>
                <div className="report-layout-main">
                  <div className="field">
                    <label className="field__label report-layout-label">Hình vẽ | Layout</label>
                    <ReportPhotoManager
                      folder={`${photosFolder}/${photoSubfolder(g.id)}`}
                      files={g.files}
                      disabled={disabled}
                      isUploading={isUploading}
                      onUpload={(files) => onUpload(g.id, files)}
                      onDelete={(name) => onDeletePhoto(g.id, name)}
                    />
                  </div>
                  <div className="field">
                    <label className="field__label report-layout-label">Mô tả | Description</label>
                    <textarea
                      className="textarea"
                      rows={5}
                      placeholder="Nhập mô tả..."
                      aria-label="Mô tả"
                      value={g.summary}
                      onChange={(e) => updateRow(g.id, { summary: e.target.value })}
                      disabled={disabled && !readOnlyMode}
                      readOnly={readOnlyMode}
                    />
                  </div>
                </div>
                <div className="report-layout-actions">
                  <button
                    type="button"
                    className="btn btn--ghost-danger btn--icon"
                    onClick={() => onDeleteGroup(g.id)}
                    disabled={disabled || isUploading}
                    title="Xoá hàng này"
                    aria-label="Xoá hàng này"
                  >
                    <Icon name="dismiss" size={16} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
