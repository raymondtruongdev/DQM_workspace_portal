import * as React from 'react';
import { PhotoGroup, ReportPhotoSection } from '../types/report.types';
import { Icon, EmptyState } from '../../../core/components';
import { ReportPhotoManager } from './ReportPhotoManager';

export interface IReportPhotoSectionEditorProps {
  section: ReportPhotoSection;
  roman: string;
  anchor: string;
  /** Thư mục ảnh của báo cáo (server-relative): data/<Unit>/<DocID>_<yyyyMMdd>/photos */
  photosFolder: string;
  disabled?: boolean;
  isUploading?: boolean;
  onChange: (section: ReportPhotoSection) => void;
  onUpload: (groupId: string, files: FileList) => Promise<void>;
  onDeletePhoto: (groupId: string, name: string) => Promise<void>;
  onDeleteGroup: (groupId: string) => Promise<void>;
}

/** Mục ảnh hiện trường: mỗi bộ phận một nhóm (tên bộ phận + ảnh + tóm tắt), thêm bằng nút "Thêm bộ phận". */
export const ReportPhotoSectionEditor: React.FC<IReportPhotoSectionEditorProps> = ({
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
  const updateGroup = (id: string, patch: Partial<PhotoGroup>): void =>
    onChange({ ...section, groups: section.groups.map((g) => (g.id === id ? { ...g, ...patch } : g)) });

  const addGroup = (): void =>
    onChange({
      ...section,
      groups: [...section.groups, { id: `g${Date.now()}`, department: '', summary: '', files: [] }]
    });

  return (
    <div id={anchor} className="section-card">
      <div className="section-card__banner">
        <span className="section-card__title">
          {roman}. {section.title}{' '}
          {section.titleEn && <span className="section-card__title-sub">| {section.titleEn}</span>}
        </span>
        <button type="button" className="btn btn--light btn--sm" onClick={addGroup} disabled={disabled || isUploading}>
          + Thêm bộ phận
        </button>
      </div>

      <div className="section-card__body">
        {section.groups.length === 0 ? (
          <EmptyState
            icon="image"
            title="Chưa có ảnh nào"
            message="Bấm “Thêm bộ phận” để nhập tên bộ phận, tải ảnh và nhập tóm tắt."
          />
        ) : (
          <div className="report-photo-group-list">
            {section.groups.map((g) => (
              <div key={g.id} className="report-photo-group">
                <div className="report-photo-group__bar">
                  <input
                    type="text"
                    className="input"
                    placeholder="Tên bộ phận (VD: DCEM, HSE...)"
                    aria-label="Tên bộ phận"
                    value={g.department}
                    onChange={(e) => updateGroup(g.id, { department: e.target.value })}
                    disabled={disabled}
                  />
                  <button
                    type="button"
                    className="btn btn--light btn--icon"
                    onClick={() => onDeleteGroup(g.id)}
                    disabled={disabled || isUploading}
                    title="Xoá bộ phận này"
                    aria-label="Xoá bộ phận này"
                  >
                    <Icon name="dismiss" size={16} />
                  </button>
                </div>

                <div className="report-photo-group__body">
                  <ReportPhotoManager
                    folder={`${photosFolder}/${g.id}`}
                    files={g.files}
                    disabled={disabled}
                    isUploading={isUploading}
                    onUpload={(files) => onUpload(g.id, files)}
                    onDelete={(name) => onDeletePhoto(g.id, name)}
                  />
                  <div className="field">
                    <label className="field__label">Tóm tắt | Summary</label>
                    <textarea
                      className="textarea"
                      rows={2}
                      placeholder="Nhập tóm tắt nội dung ảnh của bộ phận..."
                      value={g.summary}
                      onChange={(e) => updateGroup(g.id, { summary: e.target.value })}
                      disabled={disabled}
                    />
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {section.showSummary && (
        <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
          <div className="field">
            <label className="field__label">Tóm tắt chung | Summary</label>
            <textarea
              className="textarea"
              rows={2}
              value={section.summary}
              onChange={(e) => onChange({ ...section, summary: e.target.value })}
              disabled={disabled}
            />
          </div>
        </div>
      )}
    </div>
  );
};
