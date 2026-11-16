import * as React from 'react';
import { ReportMetadata, RequirableField } from '../types/report.types';
import { CreatableSelect } from '../../../core/components/CreatableSelect';
import { Icon } from '../../../core/components/Icon';
import { PromptDialog } from '../../../core/components/PromptDialog';
import { sanitizeOption } from '../utils/reportTable';
import { REQUIRED_FIELDS, getMissingRequiredFields } from '../utils/reportValidation';

export interface IReportMetaFormProps {
  metadata: ReportMetadata;
  dateStr: string;
  onDateChange: (newDate: string) => void;
  onMetadataChange: <K extends keyof ReportMetadata>(field: K, value: ReportMetadata[K]) => void;
  isExisting: boolean;
  /** Danh sách đơn vị thi công để chọn (danh mục dùng chung `subcontractor`). */
  subcontractorOptions: string[];
  /** Thêm đơn vị thi công mới vào danh mục dùng chung; trả về giá trị đã lưu (hoặc chọn sẵn nếu đã có), undefined nếu không ghi được. */
  onAddSubcontractor?: (value: string) => Promise<string | undefined>;
  /** Tọa độ để chọn (danh mục dùng chung `location`): `value` = tọa độ lưu vào báo cáo, `label` = chữ hiện trong danh sách (vd "ST01 - 10.771091, 106.698838"). */
  locationOptions: Array<{ value: string; label: string }>;
  /** Thêm tọa độ mới (kèm tên vị trí) vào danh mục dùng chung; trả về giá trị để chọn, undefined nếu không ghi được. */
  onAddLocation?: (geoTag: string, name: string) => Promise<string | undefined>;
  disabled?: boolean;
  /** When true, blank required fields are highlighted with an error message. */
  showErrors?: boolean;
}

export const ReportMetaForm: React.FC<IReportMetaFormProps> = ({
  metadata,
  dateStr,
  onDateChange,
  onMetadataChange,
  subcontractorOptions,
  onAddSubcontractor,
  locationOptions,
  onAddLocation,
  disabled = false,
  showErrors = false
}) => {
  const required = REQUIRED_FIELDS;
  const missing = showErrors ? getMissingRequiredFields(metadata) : [];
  const isReq = (f: RequirableField): boolean => required.indexOf(f) !== -1;
  const isBad = (f: RequirableField): boolean => missing.indexOf(f) !== -1;
  const mark = (f: RequirableField): React.ReactNode =>
    isReq(f) ? <span className="report-required-mark" aria-hidden="true"> *</span> : null;
  const err = (f: RequirableField): React.ReactNode =>
    isBad(f) ? <span className="report-field-error" role="alert">Bắt buộc nhập, không được để trống</span> : null;
  const selectedSubs = metadata.subcontractors;
  const availableSubs = subcontractorOptions.filter((o) => selectedSubs.indexOf(o) === -1);
  const setSubs = (next: string[]): void => onMetadataChange('subcontractors', next);
  // Tọa độ mới: hỏi tiếp tên vị trí (bắt buộc) rồi ghi vào danh mục dùng chung
  const [newGeo, setNewGeo] = React.useState<string | undefined>(undefined);
  const geoValues = locationOptions.map((o) => o.value);
  const geoLabels: Record<string, string> = {};
  locationOptions.forEach((o) => {
    geoLabels[o.value] = o.label;
  });
  const cls = (f: RequirableField): string => `input${isBad(f) ? ' input--danger' : ''}`;

  return (
    <>
      <div className="card report-info-card">
        {/* 1. Top Section: Doc Code + Template Version + Subtitle on Left, Date Picker on Right */}
        <div className="report-info-top-row">
          <div className="report-info-left-group">
            <div className="report-info-badges-inline">
              {metadata.docCode && (
                <input
                  type="text"
                  readOnly
                  className="report-doc-code"
                  value={metadata.docCode}
                  title={metadata.docCode}
                  aria-label="Mã tài liệu (chỉ đọc, có thể sao chép)"
                  onFocus={(e) => e.currentTarget.select()}
                />
              )}
            </div>
            {metadata.subTitle && (
              <div className="report-info-subtitle">
                {metadata.subTitle.replace(/\s*\|\s*/g, ' – ')}
              </div>
            )}
          </div>

          <div className="report-info-date-group">
            <label htmlFor="report-date-picker-input" className="field__label">
              Ngày báo cáo
            </label>
            <div className="report-date-wrap">
              <input
                id="report-date-picker-input"
                type="date"
                className="input report-date-width"
                value={dateStr}
                onChange={(e) => onDateChange(e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>
        </div>

        {/* 2. Form Grid (2 Columns x 2 Rows) */}
        <div className="report-info-grid-2x2">
          {/* Row 1, Col 1: Vị trí / Trạm (tự điền theo đơn vị, không sửa) */}
          <div className="field">
            <label htmlFor="report-location-input" className="field__label">Vị trí / Trạm (Location)</label>
            <input
              id="report-location-input"
              type="text"
              className="input input--readonly"
              value={metadata.location || ''}
              readOnly
              title="Tự điền theo công trường viết báo cáo"
            />
          </div>

          {/* Row 1, Col 2: Tọa độ (chọn từ danh mục `location`) */}
          <div className="field">
            <label htmlFor="report-geotag-input" className="field__label">Tọa độ (Geo-tag)</label>
            <CreatableSelect
              id="report-geotag-input"
              className="input"
              value={metadata.geoTag || ''}
              placeholder="Chọn tọa độ…"
              options={geoValues}
              optionLabels={geoLabels}
              disabled={disabled}
              dialogTitle="Thêm tọa độ mới (vd 10.771091, 106.698838)"
              onChange={(v, isNew) => {
                const clean = v.trim().replace(/\s+/g, ' ');
                if (!clean) return;
                if (isNew) setNewGeo(clean);
                else onMetadataChange('geoTag', clean);
              }}
            />
          </div>

          {/* Row 2, Col 1: Số báo cáo (tự nhảy theo từng đơn vị, không sửa) */}
          <div className="field">
            <label htmlFor="report-no-input" className="field__label">Số báo cáo (Report no.)</label>
            <input
              id="report-no-input"
              type="text"
              className="input input--readonly"
              value={metadata.reportNo || ''}
              readOnly
              title="Tự động tăng theo từng công trường"
            />
          </div>

          {/* Row 2, Col 2: Người báo cáo (Email) */}
          <div className="field">
            <label className="field__label">Người báo cáo (Email)</label>
            <input
              type="text"
              className="input input--readonly"
              value={metadata.email || ''}
              readOnly
            />
          </div>

          {/* Row 3, Col 1: Đơn vị thi công */}
          <div className="field">
            <label htmlFor="report-subcontractor-input" className="field__label">
              Đơn vị thi công (Subcontractor){mark('subcontractor')}
            </label>
            <div className="report-subcon-list">
              {selectedSubs.map((name) => (
                <span key={name} className="badge badge--info report-subcon-item">
                  {name}
                  {!disabled && (
                    <button
                      type="button"
                      className="report-subcon-remove"
                      onClick={() => setSubs(selectedSubs.filter((x) => x !== name))}
                      title={`Bỏ ${name}`}
                      aria-label={`Bỏ ${name}`}
                    >
                      <Icon name="dismiss" size={16} />
                    </button>
                  )}
                </span>
              ))}
              <CreatableSelect
                id="report-subcontractor-input"
                className={cls('subcontractor')}
                wrapperClassName="report-subcon-select"
                value=""
                placeholder={selectedSubs.length ? '+ Thêm đơn vị…' : 'Chọn đơn vị thi công…'}
                options={availableSubs}
                disabled={disabled}
                ariaInvalid={isBad('subcontractor')}
                dialogTitle="Thêm đơn vị thi công mới"
                onChange={(v, isNew) => {
                  const clean = sanitizeOption(v);
                  if (!clean) return;
                  const add = (value: string): void => {
                    if (selectedSubs.indexOf(value) === -1) setSubs(selectedSubs.concat(value));
                  };
                  // Đơn vị mới: ghi vào danh mục dùng chung (để mọi báo cáo sau thấy) rồi chọn luôn
                  if (isNew && onAddSubcontractor) {
                    onAddSubcontractor(clean).then((saved) => add(saved || clean)).catch(() => add(clean));
                  } else add(clean);
                }}
              />
            </div>
            {err('subcontractor')}
          </div>

          {/* Row 3, Col 2: Tài liệu đính kèm (người dùng tự nhập, mỗi dòng một tài liệu) */}
          <div className="field">
            <label htmlFor="report-attached-input" className="field__label">Tài liệu đính kèm (Attached Documents)</label>
            <textarea
              id="report-attached-input"
              className="textarea"
              rows={3}
              value={metadata.attachedDocs || ''}
              readOnly={disabled}
              placeholder="Mỗi dòng một tài liệu, vd: - Bản vẽ thiết kế (Design Drawings)"
              onChange={(e) => onMetadataChange('attachedDocs', e.target.value)}
            />
          </div>
        </div>
      </div>
      {newGeo !== undefined && (
        <PromptDialog
          title={`Thêm tọa độ "${newGeo}"`}
          label="Tên vị trí của tọa độ này (vd ST01)"
          onSubmit={(name) => {
            const geo = newGeo;
            setNewGeo(undefined);
            const add = (value: string): void => onMetadataChange('geoTag', value);
            if (onAddLocation) onAddLocation(geo, name.trim()).then((saved) => add(saved || geo)).catch(() => add(geo));
            else add(geo);
          }}
          onCancel={() => setNewGeo(undefined)}
        />
      )}
    </>
  );
};
