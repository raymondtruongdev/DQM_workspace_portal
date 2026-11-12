import * as React from 'react';
import { FormField, ReportForm } from '../types/report.types';
import { CreatableSelect } from '../../../core/components/CreatableSelect';
import { effectiveColumns, fieldWithOption } from '../utils/reportForm';
import { resolveTitle, sanitizeOption, splitLabel } from '../utils/reportTable';

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

function parseInputNum(val: string): number | undefined {
  if (val.trim() === '') return undefined;
  const n = parseFloat(val);
  return isNaN(n) ? undefined : n;
}

/** Mục dạng form sinh từ sheet của template: lưới trường nhập theo số cột khai báo. */
export const ReportDynamicForm: React.FC<IReportDynamicFormProps> = ({
  form,
  dateStr,
  roman,
  anchor,
  disabled = false,
  showErrors = false,
  onChange
}) => {
  const cols = effectiveColumns(form);
  const setValue = (key: string, value: FormField['value']): void =>
    onChange({ ...form, fields: form.fields.map((f) => (f.key === key ? { ...f, value } : f)) });

  const renderControl = (f: FormField, bad: boolean): React.ReactNode => {
    const id = `${anchor}-${f.key}`;
    const common = { id, disabled, 'aria-invalid': bad || undefined };
    if (f.type === 'textarea') {
      return (
        <textarea
          {...common}
          className={`textarea${bad ? ' input--danger' : ''}`}
          rows={f.rows || 3}
          placeholder={f.placeholder}
          value={typeof f.value === 'string' ? f.value : ''}
          onChange={(e) => setValue(f.key, e.target.value)}
        />
      );
    }
    if (f.type === 'number') {
      return (
        <input
          {...common}
          type="number"
          step="any"
          min="0"
          className={`input${bad ? ' input--danger' : ''}`}
          placeholder={f.placeholder}
          value={typeof f.value === 'number' ? f.value : ''}
          onChange={(e) => setValue(f.key, parseInputNum(e.target.value))}
        />
      );
    }
    if (f.type === 'select') {
      const options = f.options || [];
      const current = typeof f.value === 'string' && f.value ? f.value : options[0] || '';
      return (
        <CreatableSelect
          id={id}
          className={`input${bad ? ' input--danger' : ''}`}
          value={current}
          options={options}
          disabled={disabled}
          ariaInvalid={bad}
          dialogTitle={`Thêm giá trị mới – ${splitLabel(f.label).vi}`}
          onChange={(v, isNew) => {
            if (!isNew) {
              setValue(f.key, v);
              return;
            }
            const clean = sanitizeOption(v);
            if (!clean) return;
            onChange({
              ...form,
              fields: form.fields.map((x) => (x.key === f.key ? { ...fieldWithOption(x, clean), value: clean } : x))
            });
          }}
        />
      );
    }
    return (
      <input
        {...common}
        type="text"
        className={`input${bad ? ' input--danger' : ''}`}
        placeholder={f.placeholder}
        value={f.value === undefined ? '' : String(f.value)}
        onChange={(e) => setValue(f.key, e.target.value)}
      />
    );
  };

  return (
    <div id={anchor} className="section-card">
      <div className="section-card__banner">
        <span className="section-card__title">
          {roman}. {resolveTitle(form.title, dateStr)}{' '}
          {form.titleEn && <span className="section-card__title-sub">| {form.titleEn}</span>}
        </span>
      </div>

      <div className="section-card__body">
        <div
          className={`report-form-grid${cols > 3 ? ' report-form-grid--wide' : ''}`}
          style={{ ['--report-form-cols' as string]: String(cols) }}
        >
          {form.fields.map((f) => {
            const empty = f.value === undefined || String(f.value).trim() === '';
            const bad = showErrors && Boolean(f.required) && empty;
            const lbl = splitLabel(f.label);
            const span = f.full ? cols : Math.min(f.span || 1, cols);
            return (
              <div
                key={f.key}
                className={`field report-form-cell${f.full ? ' report-form-cell--full' : ''}${
                  f.type === 'textarea' ? ' report-form-cell--area' : ''
                }`}
                style={f.full ? undefined : { ['--report-cell-span' as string]: String(span) }}
              >
                <label htmlFor={`${anchor}-${f.key}`} className="field__label field__label--bilingual">
                  <span className="field__label-main">
                    {lbl.vi}
                    {f.required && (
                      <span className="report-required-mark" aria-hidden="true">
                        {' '}
                        *
                      </span>
                    )}
                  </span>
                  {lbl.en && <span className="field__label-sub">{lbl.en}</span>}
                </label>
                {renderControl(f, bad)}
                {bad && (
                  <span className="report-field-error" role="alert">
                    Bắt buộc nhập, không được để trống
                  </span>
                )}
              </div>
            );
          })}
        </div>
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
