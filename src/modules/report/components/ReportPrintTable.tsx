import * as React from 'react';
import { ReportTable } from '../types/report.types';
import { formFieldDisplay } from '../utils/reportForm';
import {
  columnAlign,
  columnTotal,
  firstTotalIndex,
  formatNumber,
  hasTotals,
  isoToVN,
  resolveTitle,
  splitLabel,
  toneOf
} from '../utils/reportTable';

export interface IReportPrintTableProps {
  table: ReportTable;
  roman: string;
  dateStr: string;
}

/** Mục dạng bảng trong bản in A4 (dùng class `bcn-*` do ReportPrintPreviewModal định nghĩa). */
export const ReportPrintTable: React.FC<IReportPrintTableProps> = ({ table, roman, dateStr }) => {
  const { columns, rows } = table;
  const defaultWidth = 100 / Math.max(1, columns.length);
  const totalSpan = firstTotalIndex(table);

  const cellClass = (align: string, tone?: string): string =>
    [align === 'right' ? 'text-right' : align === 'center' ? 'text-center' : '', tone ? `bcn-tone-${tone}` : '']
      .filter(Boolean)
      .join(' ');

  return (
    <>
      <div className="bcn-sec-title">
        {roman}. {resolveTitle(table.title, dateStr)} {table.titleEn && <span>| {table.titleEn}</span>}
      </div>
      <table className="bcn-table">
        <thead>
          <tr className="bcn-th-row">
            {columns.map((col) => {
              const lbl = splitLabel(col.label);
              return (
                <th key={col.code} style={{ width: `${col.width || defaultWidth}%` }}>
                  {lbl.vi}
                  {lbl.en && <span>{lbl.en}</span>}
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={r.id || i}>
              {columns.map((col) => {
                const v = r.values[col.code];
                const text =
                  col.type === 'number' || col.type === 'calc'
                    ? formatNumber(v ?? 0)
                    : col.type === 'date'
                    ? isoToVN(String(v ?? ''))
                    : v ?? '';
                return (
                  <td key={col.code} className={cellClass(columnAlign(col), toneOf(col, v))}>
                    {text}
                  </td>
                );
              })}
            </tr>
          ))}
          {hasTotals(table) && (
            <tr className="bcn-total-row">
              <td colSpan={totalSpan}>
                Tổng <span>| Total</span>
              </td>
              {columns.slice(totalSpan).map((col) => (
                <td key={col.code} className="text-right">
                  {col.total ? columnTotal(table, col).toLocaleString('vi-VN') : ''}
                </td>
              ))}
            </tr>
          )}
          {(table.fields || []).map((f) => (
            <tr key={`field-${f.key}`}>
              <td colSpan={columns.length}>
                <strong>{f.label}:</strong> <span style={{ whiteSpace: 'pre-line' }}>{formFieldDisplay(f)}</span>
              </td>
            </tr>
          ))}
          {table.showSummary && (
            <tr>
              <td colSpan={columns.length}>
                <strong>Tóm tắt | Summary:</strong> <span style={{ whiteSpace: 'pre-line' }}>{table.summary}</span>
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </>
  );
};
