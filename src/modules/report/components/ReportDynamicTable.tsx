import * as React from 'react';
import { CellValue, ReportTable, TableColumn, TableRow } from '../types/report.types';
import { ReportCatalog, applyCatalogLinks, catalogOptions, linkedValue } from '../utils/reportCatalog';
import { Icon } from '../../../core/components/Icon';
import { CreatableSelect } from '../../../core/components/CreatableSelect';
import {
  columnAlign,
  columnMinPx,
  columnTotal,
  columnWithOption,
  createRow,
  firstTotalIndex,
  formatNumber,
  hasTotals,
  isoToVN,
  recalcRow,
  resolveTitle,
  sanitizeOption,
  splitLabel,
  toneOf,
  vnToIso
} from '../utils/reportTable';

export interface IReportDynamicTableProps {
  table: ReportTable;
  unit: string;
  /** Danh mục cho các cột select:@… và from=… (chưa tải xong thì undefined). */
  catalog?: ReportCatalog;
  dateStr: string;
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
  catalog,
  dateStr,
  roman,
  anchor,
  disabled = false,
  onChange
}) => {
  const { columns, rows } = table;

  const setRows = (next: TableRow[]): void => onChange({ ...table, rows: next });

  const updateCell = (index: number, code: string, value: CellValue): void => {
    setRows(
      rows.map((row, idx) => {
        if (idx !== index) return row;
        const values = { ...row.values, [code]: value };
        // Cột from= (vd Activity theo cấu kiện) tự điền giá trị ràng buộc của mã vừa chọn
        applyCatalogLinks(columns, values, code, catalog);
        return recalcRow(columns, { ...row, values });
      })
    );
  };

  /** Giá trị mới của ô select: thêm vào danh sách chọn của cột (ghi vào spec) và gán cho ô đang sửa. */
  const selectValue = (index: number, col: TableColumn, value: string, isNew: boolean): void => {
    if (!isNew) {
      updateCell(index, col.code, value);
      return;
    }
    const clean = sanitizeOption(value);
    if (!clean) return;
    const nextCols = columns.map((c) => (c.code === col.code ? columnWithOption(c, clean) : c));
    onChange({
      ...table,
      columns: nextCols,
      rows: rows.map((row, idx) =>
        idx === index ? recalcRow(nextCols, { ...row, values: { ...row.values, [col.code]: clean } }) : row
      )
    });
  };

  const addRow = (): void => setRows([...rows, createRow(table, unit)]);

  const removeRow = (index: number): void => {
    if (rows[index]?.isLocked) return;
    setRows(rows.filter((_, idx) => idx !== index));
  };

  const tableMinWidth = columns.reduce((acc, c) => acc + columnMinPx(c), table.single ? 0 : 56);
  const renderCell = (col: TableColumn, row: TableRow, idx: number): React.ReactNode => {
    const value = row.values[col.code];
    // Cột dropdown khoá (key/readonly): dòng thêm tay luôn giữ giá trị mặc định (vd "---"), không chọn đổi được
    // Cột danh mục (@…) chỉ khoá khi khai báo readonly (hoặc key ở dòng IsLocked); cột from=… khoá khi mã đã chọn có giá trị ràng buộc
    const locked = col.optionsRef
      ? Boolean(col.readonly) || (Boolean(col.key) && Boolean(row.isLocked)) || linkedValue(columns, col, row.values, catalog) !== undefined
      : Boolean(col.readonly) || (Boolean(col.key) && (Boolean(row.isLocked) || col.type === 'select')) || linkedValue(columns, col, row.values, catalog) !== undefined;
    const align = alignClass(col);
    if (col.type === 'calc') {
      return (
        <td>
          <input
            type="text"
            readOnly
            tabIndex={-1}
            className={`input input--cell input--locked${align}`}
            value={formatNumber(value) || '0'}
            aria-label={splitLabel(col.label).vi}
          />
        </td>
      );
    }
    if (col.type === 'number') {
      return (
        <td>
          <input
            type="number"
            min="0"
            step="any"
            className={`input input--cell${locked ? ' input--locked' : ''}${align}`}
            value={typeof value === 'number' ? value : ''}
            onChange={(e) => updateCell(idx, col.code, parseInputNum(e.target.value))}
            disabled={disabled}
            readOnly={locked}
          />
        </td>
      );
    }
    if (col.type === 'date') {
      const text = typeof value === 'string' ? value : '';
      return (
        <td>
          <input
            type="date"
            className={`input input--cell${locked ? ' input--locked' : ''}${align}`}
            value={vnToIso(text)}
            onChange={(e) => updateCell(idx, col.code, isoToVN(e.target.value))}
            disabled={disabled || locked}
            title={splitLabel(col.label).vi}
          />
        </td>
      );
    }
    if (col.type === 'select') {
      const current = typeof value === 'string' ? value : '';
      const tone = toneOf(col, current);
      return (
        <td>
          <CreatableSelect
            className={`input input--cell${locked ? ' input--locked' : ''}${align}${tone ? ` input--${tone}` : ''}`}
            value={current}
            options={col.optionsRef ? catalogOptions(catalog, col.optionsRef) : col.options || []}
            allowOther={!col.optionsRef}
            onChange={(v, isNew) => selectValue(idx, col, v, isNew)}
            disabled={disabled || locked}
            dialogTitle={`Thêm giá trị mới – ${splitLabel(col.label).vi}`}
          />
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
          {roman}. {resolveTitle(table.title, dateStr)}{' '}
          {table.titleEn && <span className="section-card__title-sub">| {table.titleEn}</span>}
        </span>
        {!table.single && (
          <button type="button" className="btn btn--light btn--sm" onClick={addRow} disabled={disabled}>
            + Thêm dòng
          </button>
        )}
      </div>

      <div className="table-responsive">
        <table className="portal-table report-grid-table" style={{ minWidth: tableMinWidth }}>
          <thead>
            <tr>
              {columns.map((col) => {
                const lbl = splitLabel(col.label);
                return (
                  <th
                    key={col.code}
                    style={{ minWidth: columnMinPx(col), textAlign: columnAlign(col) }}
                  >
                    <span className="report-th-vi">{lbl.vi}</span>
                    {lbl.en && <span className="report-th-en">{lbl.en}</span>}
                  </th>
                );
              })}
              {!table.single && <th style={{ width: 56, minWidth: 56, textAlign: 'center' }} />}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={row.id || idx} className={row.isLocked ? 'report-row-locked' : ''}>
                {columns.map((col) => (
                  <React.Fragment key={col.code}>{renderCell(col, row, idx)}</React.Fragment>
                ))}
                {!table.single && (
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
                )}
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
                {!table.single && <td />}
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {table.showSummary && (
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
      )}
    </div>
  );
};
