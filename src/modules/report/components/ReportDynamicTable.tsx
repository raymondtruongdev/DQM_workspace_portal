import * as React from 'react';
import { ReportReadOnlyContext } from './ReportReadOnlyContext';
import { CellValue, FormField, ReportTable, TableColumn, TableRow } from '../types/report.types';
import { ReportFieldsGrid } from './ReportFieldsGrid';
import { ReportCatalog, applyCatalogLinks, canAddToTable, catalogOptionLabels, catalogOptions, findCatalogItem, linkedValue } from '../utils/reportCatalog';
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
  recalcRow,
  resolveTitle,
  sanitizeOption,
  splitLabel,
  toneOf
} from '../utils/reportTable';

/** Hỏi giá trị liên kết của mục mới; `undefined` = người dùng huỷ. */
/** `field` = cột của danh mục mà giá trị nhập được ghi vào (cột phụ thuộc có `fromField`); không có = Link. */
export type AskCatalogLink = () => Promise<{ link: string; field?: 'note' | 'nameVn' } | undefined>;

/** Kết quả thêm mục danh mục: giá trị gán cho ô + danh mục mới, hoặc người dùng đã huỷ. */
export type CatalogAddResult = { value: string; catalog: ReportCatalog; cancelled?: false } | { cancelled: true };

export interface IReportDynamicTableProps {
  table: ReportTable;
  unit: string;
  /** Danh mục cho các cột select:@… và from=… (chưa tải xong thì undefined). */
  catalog?: ReportCatalog;
  /**
   * Người dùng chọn lại giá trị cho cột phụ thuộc (vd Activity) của một mục do người dùng thêm (`AddedByAdmin = 0`):
   * ghi liên kết mới vào danh mục dùng chung (`ref` = bảng của cột nguồn, `code` = mã đang chọn).
   */
  onUpdateCatalogLink?: (ref: string, code: string, link: string) => Promise<void>;
  /**
   * Người dùng thêm giá trị mới ("Thêm mới…") vào cột danh mục: ghi vào list DRPCatalogItems (chống trùng). Trả về giá trị cần gán cho ô
   * và danh mục mới; trả undefined nếu không ghi được (khi đó giá trị chỉ lưu trong spec của báo cáo này), `{ cancelled }` nếu người dùng huỷ.
   * `askLink` hỏi giá trị liên kết của mục mới (vd Activity của cấu kiện mới; link '' = bỏ qua, undefined = huỷ).
   */
  onAddCatalogItem?: (ref: string, value: string, byLink: boolean, askLink?: AskCatalogLink) => Promise<CatalogAddResult | undefined>;
  /** Số lớn nhất đã dùng ở các báo cáo khác cho mã tự tăng (Event ID, Issue ID…), để dòng mới nhận số kế tiếp. */
  idCounters?: Record<string, number>;
  /** Thư mục (server-relative) chứa PDF đính kèm của các dòng (mục có `rowPdf`, vd Sự kiện). */
  rowPdfFolder?: string;
  /** Tải PDF đính kèm cho một dòng. Nơi gọi tải lên SharePoint (hỏi Thay thế/Huỷ khi trùng tên) và cập nhật `files` của dòng. */
  onUploadRowPdf?: (rowId: string, files: FileList) => Promise<void>;
  /** Xoá một PDF đính kèm của dòng (xoá file và bỏ khỏi `files`). */
  onDeleteRowPdf?: (rowId: string, name: string) => Promise<void>;
  /** Xoá các file PDF của một dòng sắp bị xoá (chỉ xoá file, bỏ qua file dòng khác còn dùng; dòng do bảng tự bỏ). */
  onPurgeRowPdf?: (rowId: string, names: string[]) => Promise<void>;
  dateStr: string;
  roman: string;
  anchor: string;
  disabled?: boolean;
  /** Tô đỏ các trường bắt buộc đang trống. */
  showErrors?: boolean;
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

/** Mục dạng bảng: tiêu đề, bảng nhập, dòng Tổng, các trường văn bản (nếu có), ô Tóm tắt. */
export const ReportDynamicTable: React.FC<IReportDynamicTableProps> = ({
  table,
  unit,
  catalog,
  onAddCatalogItem,
  onUpdateCatalogLink,
  rowPdfFolder,
  onUploadRowPdf,
  onDeleteRowPdf,
  onPurgeRowPdf,
  idCounters,
  dateStr,
  roman,
  anchor,
  disabled = false,
  showErrors = false,
  onChange
}) => {
  const readOnlyMode = React.useContext(ReportReadOnlyContext);
  const { columns, rows } = table;

  const setRows = (next: TableRow[]): void => onChange({ ...table, rows: next });

  // Hộp thoại hỏi liên kết khi thêm mã mới (bước 2 của "Thêm mới…")
  const [linkAsk, setLinkAsk] = React.useState<
    { title: string; label: string; options?: string[]; field?: 'note' | 'nameVn'; resolve: (v: { link: string; field?: 'note' | 'nameVn' } | undefined) => void } | undefined
  >(undefined);

  /** Cột mã mà mục mới phải kèm tên Việt (vật tư, nhân lực, thiết bị): hộp thoại đầu hỏi "Nhập <tên cột> mới" thay vì "Thêm giá trị mới". */
  const codeDialogTitle = (col: TableColumn): string => {
    const label = splitLabel(col.label).vi;
    const asksName = col.askName === 'nameVn' || columns.some((c) => c.from && c.from.toLowerCase() === col.code.toLowerCase() && c.fromField === 'nameVn');
    return asksName ? `Nhập ${label.charAt(0).toLowerCase()}${label.slice(1)} mới` : `Thêm giá trị mới – ${label}`;
  };

  /** Cột phụ thuộc vào cột mã `col` (có from=<col>) mà mục mới cần liên kết; hỏi người dùng giá trị (có thể bỏ qua). */
  const makeAskLink = (col: TableColumn, value: string): AskCatalogLink | undefined => {
    const dep = columns.filter((c) => c.from && c.from.toLowerCase() === col.code.toLowerCase())[0];
    if (!dep && col.askName) {
      return () =>
        new Promise<{ link: string; field?: 'note' | 'nameVn' } | undefined>((resolve) =>
          setLinkAsk({ title: 'Nhập tên VN', label: `Tên tiếng Việt của mã "${value}"`, field: col.askName, resolve })
        );
    }
    if (!dep) return undefined;
    const options =
      dep.type === 'select'
        ? (dep.optionsRef ? catalogOptions(catalog, dep.optionsRef) : dep.options || []).filter((o) => o !== '---')
        : undefined;
    if (options && options.length === 0) return undefined; // chưa có gì để chọn: bỏ qua bước hỏi
    return () =>
      new Promise<{ link: string; field?: 'note' | 'nameVn' } | undefined>((resolve) =>
        setLinkAsk({
          title: dep.fromField === 'nameVn' ? 'Nhập tên VN' : `Thêm "${value}"`,
          label: dep.fromField === 'nameVn'
            ? `Tên tiếng Việt của mã "${value}"`
            : dep.fromField
            ? `${splitLabel(dep.label).vi} của mã "${value}"`
            : `${splitLabel(dep.label).vi} của "${value}" (không bắt buộc, bỏ trống nếu chưa có)`,
          options,
          field: dep.fromField,
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
    // Cột phụ thuộc (vd Activity) đổi giá trị cho mục do người dùng thêm: cập nhật liên kết trong danh mục
    const dep = columns.filter((c) => c.code === code)[0];
    const src = dep && dep.from && !dep.fromField ? columns.filter((c) => c.code.toLowerCase() === dep.from!.toLowerCase())[0] : undefined;
    if (src && src.optionsRef && onUpdateCatalogLink && typeof value === 'string' && value !== '---' && value !== '') {
      const srcCode = String(rows[index]?.values[src.code] ?? '');
      const item = findCatalogItem(catalog, src.optionsRef, srcCode);
      if (item && item.addedByUser && item.link !== value) onUpdateCatalogLink(src.optionsRef, srcCode, value).catch(() => undefined);
    }
  };

  /** Giá trị mới của ô select: thêm vào danh sách chọn của cột (ghi vào spec) và gán cho ô đang sửa. */
  const selectValue = async (index: number, col: TableColumn, value: string, isNew: boolean): Promise<void> => {
    if (!isNew) {
      updateCell(index, col.code, value);
      return;
    }
    // Cột danh mục: ghi vào list DRPCatalogItems (có chống trùng); chọn lại mục sẵn có nếu đã tồn tại
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

  const addRow = (): void => setRows([...rows, createRow(table, unit, idCounters)]);

  const removeRow = async (index: number): Promise<void> => {
    const row = rows[index];
    if (!row || row.isLocked) return;
    // Dòng có PDF đính kèm: xoá cả file trên SharePoint (có xác nhận), tránh để lại file mồ côi
    if (table.rowPdf && row.files && row.files.length > 0 && onPurgeRowPdf) {
      if (!window.confirm(`Xoá dòng này và ${row.files.length} file PDF đính kèm khỏi SharePoint?`)) return;
      try {
        await onPurgeRowPdf(row.id, row.files);
      } catch {
        return;
      }
    }
    setRows(rows.filter((r) => r.id !== row.id));
  };

  // PDF đính kèm của dòng: một ô chọn file dùng chung, nhớ dòng đang đính kèm
  const pdfInputRef = React.useRef<HTMLInputElement | null>(null);
  const pdfTarget = React.useRef<string | null>(null);
  const [pdfBusyRow, setPdfBusyRow] = React.useState<string | null>(null);
  const pickPdf = (row: TableRow): void => {
    pdfTarget.current = row.id;
    if (pdfInputRef.current) {
      pdfInputRef.current.value = '';
      pdfInputRef.current.click();
    }
  };
  const handlePdfPicked = async (files: FileList | null): Promise<void> => {
    const target = pdfTarget.current;
    if (!files || files.length === 0 || !target || !onUploadRowPdf) return;
    setPdfBusyRow(target);
    try {
      await onUploadRowPdf(target, files);
    } finally {
      setPdfBusyRow(null);
    }
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
            disabled={disabled && !readOnlyMode}
            readOnly={locked || readOnlyMode}
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
            value={text}
            onChange={(e) => updateCell(idx, col.code, e.target.value)}
            disabled={(disabled && !readOnlyMode) || locked}
            readOnly={readOnlyMode}
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
                    catalogOptions(catalog, col.optionsRef, col.optionsBy === 'link', col.noNone)
                  )
                : col.options || []
            }
            compactSelected
            placeholder={current === '' ? 'Chọn…' : undefined}
            optionLabels={col.optionsRef ? catalogOptionLabels(catalog, col.optionsRef, col.optionsBy === 'link') : undefined}
            allowOther={col.optionsRef ? canAddToTable(catalog, col.optionsRef) : true}
            onChange={(v, isNew) => {
              selectValue(idx, col, v, isNew).catch(() => undefined);
            }}
            disabled={disabled || locked}
            dialogTitle={codeDialogTitle(col)}
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
          disabled={disabled && !readOnlyMode}
          readOnly={locked || readOnlyMode}
          placeholder={col.placeholder}
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
        <table className={`portal-table report-grid-table${table.single ? ' report-grid-table--fixed' : ''}`} style={{ minWidth: tableMinWidth }}>
          <thead>
            <tr>
              {columns.map((col) => {
                const lbl = splitLabel(col.label);
                return (
                  <th
                    key={col.code}
                    style={{ minWidth: columnMinPx(col), ...(table.single || col.fit ? { width: columnMinPx(col) } : {}), textAlign: columnAlign(col) }}
                  >
                    <span className="report-th-vi">{lbl.vi}</span>
                    {lbl.en && <span className="report-th-en">{lbl.en}</span>}
                  </th>
                );
              })}
              {!table.single && <th style={{ width: table.rowPdf ? 96 : 56, minWidth: table.rowPdf ? 96 : 56, textAlign: 'center' }} />}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <React.Fragment key={row.id || idx}>
              <tr className={row.isLocked ? 'report-row-locked' : ''}>
                {columns.map((col) => (
                  <React.Fragment key={col.code}>{renderCell(col, row, idx)}</React.Fragment>
                ))}
                {!table.single && (
                  <td className="text-center">
                    {table.rowPdf && onUploadRowPdf && (
                      <button
                        type="button"
                        className="btn btn--ghost btn--icon btn--sm"
                        onClick={() => pickPdf(row)}
                        disabled={disabled || pdfBusyRow === row.id}
                        title="Đính kèm file PDF cho dòng này"
                        aria-label="Đính kèm file PDF"
                      >
                        <Icon name={pdfBusyRow === row.id ? 'refresh' : 'attach'} size={16} spin={pdfBusyRow === row.id} />
                      </button>
                    )}
                    {row.isLocked ? (
                      <span className="report-row-lock-badge" title="Mục cố định theo template (không được xoá)">
                        <Icon name="lock" size={16} />
                      </span>
                    ) : (
                      <button
                        type="button"
                        className="btn btn--ghost-danger btn--icon btn--sm"
                        onClick={() => {
                          removeRow(idx).catch(() => undefined);
                        }}
                        disabled={disabled}
                        title="Xoá dòng"
                      >
                        <Icon name="dismiss" size={16} />
                      </button>
                    )}
                  </td>
                )}
              </tr>
              {table.rowPdf && row.files && row.files.length > 0 && (
                <tr className="report-row-pdf">
                  <td colSpan={columns.length + 1}>
                    <div className="report-row-pdf__list">
                      <span className="report-row-pdf__label">
                        <Icon name="document-pdf" size={16} />
                        PDF đính kèm
                      </span>
                      {row.files.map((name) => (
                        <span key={name} className="report-row-pdf__item">
                          <a
                            className="btn btn--secondary btn--sm"
                            href={`${rowPdfFolder || ''}/${encodeURIComponent(name)}`}
                            target="_blank"
                            rel="noopener noreferrer"
                            title={`Mở ${name}`}
                          >
                            <Icon name="document-pdf" size={16} />
                            <span className="report-row-pdf__name">{name}</span>
                          </a>
                          {onDeleteRowPdf && (
                            <button
                              type="button"
                              className="btn btn--ghost-danger btn--icon btn--sm"
                              onClick={() => {
                                if (window.confirm(`Xoá file "${name}" khỏi SharePoint?`)) onDeleteRowPdf(row.id, name).catch(() => undefined);
                              }}
                              disabled={disabled}
                              title="Xoá file PDF này"
                              aria-label={`Xoá file ${name}`}
                            >
                              <Icon name="dismiss" size={16} />
                            </button>
                          )}
                        </span>
                      ))}
                    </div>
                  </td>
                </tr>
              )}
              </React.Fragment>
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

      {table.rowPdf && onUploadRowPdf && (
        <input
          ref={pdfInputRef}
          type="file"
          accept="application/pdf,.pdf"
          multiple
          style={{ display: 'none' }}
          tabIndex={-1}
          aria-label="Chọn file PDF đính kèm"
          onChange={(e) => {
            handlePdfPicked(e.target.files).catch(() => undefined);
          }}
        />
      )}

      {table.fields && table.fields.length > 0 && (
        <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
          <ReportFieldsGrid
            fields={table.fields}
            cols={table.fieldColumns || 1}
            anchor={`${anchor}-f`}
            disabled={disabled}
            showErrors={showErrors}
            onChange={(fields: FormField[]) => onChange({ ...table, fields })}
          />
        </div>
      )}

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
            disabled={disabled && !readOnlyMode}
            readOnly={readOnlyMode}
          />
        </div>
      </div>
      )}
      {linkAsk && (
        <PromptDialog
          title={linkAsk.title}
          label={linkAsk.label}
          options={linkAsk.options}
          allowEmpty={!linkAsk.field}
          onSubmit={(v) => {
            linkAsk.resolve({ link: v, field: linkAsk.field });
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
