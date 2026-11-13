import * as React from 'react';
import { CellValue, ReportTable, TableColumn, TableRow } from '../types/report.types';
import { ReportCatalog, applyCatalogLinks, canAddToTable, catalogOptions, linkedValue } from '../utils/reportCatalog';
import { Icon } from '../../../core/components/Icon';
import { PromptDialog } from '../../../core/components/PromptDialog';
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

/** Hỏi giá trị liên kết của mục mới; `undefined` = người dùng huỷ. */
export type AskCatalogLink = () => Promise<{ link: string } | undefined>;

/** Kết quả thêm mục danh mục: giá trị gán cho ô + danh mục mới, hoặc người dùng đã huỷ. */
export type CatalogAddResult = { value: string; catalog: ReportCatalog; cancelled?: false } | { cancelled: true };

export interface IReportDynamicTableProps {
  table: ReportTable;
  unit: string;
  /** Danh mục cho các cột select:@… và from=… (chưa tải xong thì undefined). */
  catalog?: ReportCatalog;
  /**
   * Người dùng thêm giá trị mới ("Khác…") vào cột danh mục: ghi vào list CatalogItems (chống trùng). Trả về giá trị cần gán cho ô
   * và danh mục mới; trả undefined nếu không ghi được (khi đó giá trị chỉ lưu trong spec của báo cáo này), `{ cancelled }` nếu người dùng huỷ.
   * `askLink` hỏi giá trị liên kết của mục mới (vd Activity của cấu kiện mới; link '' = bỏ qua, undefined = huỷ).
   */
  onAddCatalogItem?: (ref: string, value: string, byLink: boolean, askLink?: AskCatalogLink) => Promise<CatalogAddResult | undefined>;
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
  onAddCatalogItem,
  dateStr,
  roman,
  anchor,
  disabled = false,
  onChange
}) => {
  const { columns, rows } = table;

  const setRows = (next: TableRow[]): void => onChange({ ...table, rows: next });

  // Hộp thoại hỏi liên kết khi thêm mã mới (bước 2 của "Khác…")
  const [linkAsk, setLinkAsk] = React.useState<
    { title: string; label: string; options?: string[]; resolve: (v: { link: string } | undefined) => void } | undefined
  >(undefined);

  /** Cột phụ thuộc vào cột mã `col` (có from=<col>) mà mục mới cần liên kết; hỏi người dùng giá trị (có thể bỏ qua). */
  const makeAskLink = (col: TableColumn, value: string): AskCatalogLink | undefined => {
    const dep = columns.filter((c) => c.from && c.from.toLowerCase() === col.code.toLowerCase())[0];
    if (!dep) return undefined;
    const options =
      dep.type === 'select'
        ? (dep.optionsRef ? catalogOptions(catalog, dep.optionsRef) : dep.options || []).filter((o) => o !== '---')
        : undefined;
    if (options && options.length === 0) return undefined; // chưa có gì để chọn: bỏ qua bước hỏi
    return () =>
      new Promise<{ link: string } | undefined>((resolve) =>
        setLinkAsk({
          title: `Thêm "${value}"`,
          label: `${splitLabel(dep.label).vi} của "${value}" (không bắt buộc, bỏ trống nếu chưa có)`,
          options,
          resolve
        })
      );
  };

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
  const selectValue = async (index: number, col: TableColumn, value: string, isNew: boolean): Promise<void> => {
    if (!isNew) {
      updateCell(index, col.code, value);
      return;
    }
    // Cột danh mục: ghi vào list CatalogItems (có chống trùng); chọn lại mục sẵn có nếu đã tồn tại
    if (col.optionsRef && onAddCatalogItem) {
      const byLink = col.optionsBy === 'link';
      const added = await onAddCatalogItem(col.optionsRef, value, byLink, byLink ? undefined : makeAskLink(col, value.trim()));
      if (added && added.cancelled) return; // người dùng huỷ ở bước hỏi liên kết
      if (added) {
        setRows(
          rows.map((row, idx) => {
            if (idx !== index) return row;
            const values = { ...row.values, [col.code]: added.value };
            applyCatalogLinks(columns, values, col.code, added.catalog);
            return recalcRow(columns, { ...row, values });
          })
        );
        return;
      }
    }
    const clean = sanitizeOption(value);
    if (!clean) return;
    const nextCols = columns.map((c) => (c.code === col.code ? columnWithOption(c, clean) : c));
    onChange({
      ...table,
      columns: nextCols,
      rows: rows.map((row, idx) => {
        if (idx !== index) return row;
        const values = { ...row.values, [col.code]: clean };
        // Giá trị mới chưa có ràng buộc: các ô from= của cột này về "---" (hoặc trống)
        applyCatalogLinks(nextCols, values, col.code, catalog);
        return recalcRow(nextCols, { ...row, values });
      })
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
            options={
              col.optionsRef
                ? ((base) => base.concat((col.options || []).filter((o) => base.indexOf(o) === -1)))(
                    catalogOptions(catalog, col.optionsRef, col.optionsBy === 'link')
                  )
                : col.options || []
            }
            allowOther={col.optionsRef ? canAddToTable(catalog, col.optionsRef) : true}
            onChange={(v, isNew) => {
              selectValue(idx, col, v, isNew).catch(() => undefined);
            }}
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
      {linkAsk && (
        <PromptDialog
          title={linkAsk.title}
          label={linkAsk.label}
          options={linkAsk.options}
          allowEmpty
          onSubmit={(v) => {
            linkAsk.resolve({ link: v });
            setLinkAsk(undefined);
          }}
          onCancel={() => {
            linkAsk.resolve(undefined);
            setLinkAsk(undefined);
          }}
        />
      )}
    </div>
  );
};
