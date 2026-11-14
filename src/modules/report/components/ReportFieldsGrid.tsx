import * as React from 'react';
import { FormField } from '../types/report.types';
import { CreatableSelect } from '../../../core/components/CreatableSelect';
import { fieldWithOption } from '../utils/reportForm';
import { sanitizeOption, splitLabel } from '../utils/reportTable';

export interface IReportFieldsGridProps {
  fields: FormField[];
  /** Số cột lưới. */
  cols: number;
  /** Tiền tố id của các ô (để nhãn trỏ đúng ô và để cuộn tới ô lỗi). */
  anchor: string;
  disabled?: boolean;
  /** Tô đỏ các trường bắt buộc đang trống. */
  showErrors?: boolean;
  onChange: (fields: FormField[]) => void;
}

function parseInputNum(val: string): number | undefined {
  if (val.trim() === '') return undefined;
  const n = parseFloat(val);
  return isNaN(n) ? undefined : n;
}

/** Lưới các trường văn bản của một mục (ghi chú, ký duyệt, chỉ đạo…): dùng trong mục dạng form và dưới bảng. */
export const ReportFieldsGrid: React.FC<IReportFieldsGridProps> = ({
  fields,
  cols,
  anchor,
  disabled = false,
  showErrors = false,
  onChange
}) => {
  const setValue = (key: string, value: FormField['value']): void =>
    onChange(fields.map((f) => (f.key === key ? { ...f, value } : f)));

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
            onChange(fields.map((x) => (x.key === f.key ? { ...fieldWithOption(x, clean), value: clean } : x)));
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
    <div
      className={`report-form-grid${cols > 3 ? ' report-form-grid--wide' : ''}`}
      style={{ ['--report-form-cols' as string]: String(cols) }}
    >
      {fields.map((f) => {
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
  );
};
