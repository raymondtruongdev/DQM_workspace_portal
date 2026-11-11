import * as React from 'react';
import { CellValue, ReportTable, TableColumn, TableRow } from '../types/report.types';
import { Icon } from '../../../core/components/Icon';
import {
  columnAlign,
  columnTotal,
  createRow,
  firstTotalIndex,
  formatNumber,
  hasTotals,
  recalcRow,
  splitLabel,
  toneOf
} from '../utils/reportTable';

export interface IReportDynamicTableProps {
  table: ReportTable;
  unit: string;
  roman: string;
  anchor: string;
  disabled?: boolean;
  onChange: (table: ReportTable) => void;
}

function parseInputNum(val: string): number | undefined {
  if (val.trim() === '') return undefined;
  const n = parseFloat(val);
  return isNaN(n) ? undefined : n;
}

const alignClass = (col: TableColumn): string => {
  const a = columnAlign(col);
  return a === 'right' ? ' text-right' : a === 'center' ? ' text-center' : '';
};

/** Mục dạng bảng sinh từ sheet của template: tiêu đề, bảng nhập, dòng Tổng, ô Tóm tắt. */
export const ReportDynamicTable: React.FC<IReportDynamicTableProps> = ({
  table,
  unit,
  roman,
  anchor,
  disabled = false,
  onChange
}) => {
  const { columns, rows } = table;

  const setRows = (next: TableRow[]): void => onChange({ ...table, rows: next });

  const updateCell = (index: number, code: string, value: CellValue): void => {
    setRows(
      rows.map((row, idx) =>
        idx === index ? recalcRow(columns, { ...row, values: { ...row.values, [code]: value } }) : row
      )
    );
  };

  const addRow = (): void => setRows([...rows, createRow(table, unit)]);

  const removeRow = (index: number): void => {
    if (rows[index]?.isLocked) return;
    setRows(rows.filter((_, idx) => idx !== index));
  };

  const defaultWidth = 100 / Math.max(1, columns.length);
  const renderCell = (col: TableColumn, row: TableRow, idx: number): React.ReactNode => {
    const value = row.values[col.code];
    const locked = Boolean(row.isLocked) && Boolean(col.key);
    const align = alignClass(col);
    if (col.type === 'calc') {
      return <td className={`font-semibold${align}`}>{formatNumber(value) || '0'}</td>;
    }
    if (col.type === 'number') {
      return (
        <td>
          <input
            type="number"
            min="0"
            step="any"
            className={`input input--cell${align}`}
            value={typeof value === 'number' ? value : ''}
            onChange={(e) => updateCell(idx, col.code, parseInputNum(e.target.value))}
            disabled={disabled}
          />
        </td>
      );
    }
    if (col.type === 'select') {
      const current = typeof value === 'string' ? value : '';
      const options = col.options || [];
      const list = current && options.indexOf(current) === -1 ? [current].concat(options) : options;
      const tone = toneOf(col, current);
      return (
        <td>
          <select
            className={`input input--cell${align}${tone ? ` input--${tone}` : ''}`}
            value={current}
            onChange={(e) => updateCell(idx, col.code, e.target.value)}
            disabled={disabled}
          >
            {list.map((o) => (
              <option key={o} value={o}>
                {o}
              </option>
            ))}
          </select>
        </td>
      );
    }
    return (
      <td>
        <input
          type="text"
          className={`input input--cell${locked ? ' input--locked' : ''}${align}`}
          value={typeof value === 'string' ? value : value === undefined ? '' : String(value)}
          onChange={(e) => updateCell(idx, col.code, e.target.value)}
          disabled={disabled}
          readOnly={locked}
          title={splitLabel(col.label).vi}
        />
      </td>
    );
  };

  const totalLabelSpan = firstTotalIndex(table);

  return (
    <div id={anchor} className="section-card">
      <div className="section-card__banner">
        <span className="section-card__title">
          {roman}. {table.title}{' '}
          {table.titleEn && <span className="section-card__title-sub">| {table.titleEn}</span>}
        </span>
        <button type="button" className="btn btn--light btn--sm" onClick={addRow} disabled={disabled}>
          + Thêm dòng
        </button>
      </div>

      <div className="table-responsive">
        <table className="portal-table report-grid-table">
          <thead>
            <tr>
              {columns.map((col) => {
                const lbl = splitLabel(col.label);
                return (
                  <th
                    key={col.code}
                    style={{ width: `${col.width || defaultWidth}%`, textAlign: columnAlign(col) }}
                  >
                    {lbl.vi}
                    {lbl.en ? ` | ${lbl.en}` : ''}
                  </th>
                );
              })}
              <th style={{ width: '5%', textAlign: 'center' }} />
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={row.id || idx} className={row.isLocked ? 'report-row-locked' : ''}>
                {columns.map((col) => (
                  <React.Fragment key={col.code}>{renderCell(col, row, idx)}</React.Fragment>
                ))}
                <td className="text-center">
                  {row.isLocked ? (
                    <span className="report-row-lock-badge" title="Mục cố định theo template (không được xoá)">
                      <Icon name="lock" size={16} />
                    </span>
                  ) : (
                    <button
                      type="button"
                      className="btn btn--ghost-danger btn--icon btn--sm"
                      onClick={() => removeRow(idx)}
                      disabled={disabled}
                      title="Xoá dòng"
                    >
                      <Icon name="dismiss" size={16} />
                    </button>
                  )}
                </td>
              </tr>
            ))}
            {hasTotals(table) && (
              <tr className="report-table-total-row">
                <td colSpan={totalLabelSpan}>
                  <strong>Tổng | Total</strong>
                </td>
                {columns.slice(totalLabelSpan).map((col) => (
                  <td key={col.code} className="text-right">
                    {col.total && <strong>{columnTotal(table, col).toLocaleString('vi-VN')}</strong>}
                  </td>
                ))}
                <td />
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
        <div className="field">
          <label className="field__label">Tóm tắt | Summary</label>
          <textarea
            className="textarea"
            rows={2}
            placeholder="Nhập tóm tắt (nếu có)..."
            value={table.summary}
            onChange={(e) => onChange({ ...table, summary: e.target.value })}
            disabled={disabled}
          />
        </div>
      </div>
    </div>
  );
};
