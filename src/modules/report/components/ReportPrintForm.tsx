import * as React from 'react';
import { FormField, ReportForm } from '../types/report.types';
import { buildFormBlocks, effectiveColumns, formFieldDisplay } from '../utils/reportForm';
import { resolveTitle, splitLabel } from '../utils/reportTable';

export interface IReportPrintFormProps {
  form: ReportForm;
  roman: string;
  dateStr: string;
}

/** Mục dạng form trong bản in A4 (dùng class `bcn-*` do ReportPrintPreviewModal định nghĩa). */
export const ReportPrintForm: React.FC<IReportPrintFormProps> = ({ form, roman, dateStr }) => {
  const cols = effectiveColumns(form);
  const blocks = buildFormBlocks(form);

  const cellWidth = (f: FormField): string => `${(Math.min(f.span || 1, cols) / cols) * 100}%`;

  const headCell = (f: FormField): React.ReactNode => {
    const lbl = splitLabel(f.label);
    return (
      <th key={f.key} style={{ width: cellWidth(f) }} colSpan={Math.min(f.span || 1, cols)}>
        {lbl.vi}
        {lbl.en && <span>{lbl.en}</span>}
      </th>
    );
  };

  const valueCell = (f: FormField): React.ReactNode => {
    const span = Math.min(f.span || 1, cols);
    if (form.sign) {
      return (
        <td key={f.key} colSpan={span} className="bcn-sign-box">
          <div style={{ fontWeight: 600 }}>{formFieldDisplay(f)}</div>
        </td>
      );
    }
    const multiline = f.type === 'textarea';
    return (
      <td
        key={f.key}
        colSpan={span}
        className={multiline ? 'bcn-preline' : 'text-center'}
        style={multiline ? { verticalAlign: 'top' } : undefined}
      >
        {formFieldDisplay(f)}
      </td>
    );
  };

  const first = blocks[0] && blocks[0].kind === 'short' ? blocks[0] : undefined;
  const rest = first ? blocks.slice(1) : blocks;

  return (
    <div className={form.sign ? 'bcn-keep-together' : undefined}>
      <div className="bcn-sec-title">
        {roman}. {resolveTitle(form.title, dateStr)} {form.titleEn && <span>| {form.titleEn}</span>}
      </div>
      <table className="bcn-table">
        {first && (
          <thead>
            <tr className="bcn-th-row">{first.fields.map(headCell)}</tr>
          </thead>
        )}
        <tbody>
          {first && <tr>{first.fields.map(valueCell)}</tr>}
          {rest.map((b, i) =>
            b.kind === 'full' ? (
              <tr key={`f-${b.field.key}`}>
                <td colSpan={cols}>
                  <strong>{b.field.label}:</strong> <span style={{ whiteSpace: 'pre-line' }}>{formFieldDisplay(b.field)}</span>
                </td>
              </tr>
            ) : (
              <React.Fragment key={`s-${i}`}>
                <tr className="bcn-th-row">{b.fields.map(headCell)}</tr>
                <tr>{b.fields.map(valueCell)}</tr>
              </React.Fragment>
            )
          )}
          {form.showSummary && form.summary && (
            <tr>
              <td colSpan={cols}>
                <strong>Tóm tắt | Summary:</strong> <span style={{ whiteSpace: 'pre-line' }}>{form.summary}</span>
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
};
