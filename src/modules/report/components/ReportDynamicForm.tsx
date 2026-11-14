import * as React from 'react';
import { FormField, ReportForm } from '../types/report.types';
import { effectiveColumns } from '../utils/reportForm';
import { resolveTitle } from '../utils/reportTable';
import { ReportFieldsGrid } from './ReportFieldsGrid';

export interface IReportDynamicFormProps {
  form: ReportForm;
  dateStr: string;
  roman: string;
  anchor: string;
  disabled?: boolean;
  /** Tô đỏ các trường bắt buộc đang trống. */
  showErrors?: boolean;
  onChange: (form: ReportForm) => void;
}

/** Mục chỉ có các trường văn bản (ký duyệt, chỉ đạo…): lưới trường nhập theo số cột khai báo trong template. */
export const ReportDynamicForm: React.FC<IReportDynamicFormProps> = ({
  form,
  dateStr,
  roman,
  anchor,
  disabled = false,
  showErrors = false,
  onChange
}) => {
  const setFields = (fields: FormField[]): void => onChange({ ...form, fields });

  return (
    <div id={anchor} className="section-card">
      <div className="section-card__banner">
        <span className="section-card__title">
          {roman}. {resolveTitle(form.title, dateStr)}{' '}
          {form.titleEn && <span className="section-card__title-sub">| {form.titleEn}</span>}
        </span>
      </div>

      <div className="section-card__body">
        <ReportFieldsGrid
          fields={form.fields}
          cols={effectiveColumns(form)}
          anchor={anchor}
          disabled={disabled}
          showErrors={showErrors}
          onChange={setFields}
        />
      </div>

      {form.showSummary && (
        <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
          <div className="field">
            <label className="field__label">Tóm tắt | Summary</label>
            <textarea
              className="textarea"
              rows={2}
              placeholder="Nhập tóm tắt (nếu có)..."
              value={form.summary}
              onChange={(e) => onChange({ ...form, summary: e.target.value })}
              disabled={disabled}
            />
          </div>
        </div>
      )}
    </div>
  );
};
