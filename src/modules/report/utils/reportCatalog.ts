import type * as XLSXTypes from 'xlsx';
import { TableColumn, CellValue } from '../types/report.types';

/**
 * Danh mục dùng chung cho các dropdown, gộp từ hai nguồn (xem AGENTS.md mục 5.3):
 *  - File `ActivityP6.xlsx` (chỉ admin ghi, người dùng chỉ đọc): MỘT sheet, mỗi dòng một Activity với các cột
 *    `ActivityID` (mã), `Unit` (đơn vị, vd ST01; `ALL` = dùng chung mọi đơn vị) và `Description` (mô tả, không bắt buộc, hiện kèm mã
 *    trong dropdown). Web lọc theo đơn vị đang mở. Muốn bỏ một mã thì xoá dòng; báo cáo cũ vẫn giữ giá trị đã lưu.
 *  - SharePoint List `CatalogItems` (người dùng được tự thêm bằng "Khác…", xem ReportCatalogService): Element, Resource, Vật tư,
 *    và các danh sách chọn dùng chung (thời tiết, trạng thái…).
 * Mã "---" luôn có sẵn trong dropdown làm giá trị "chưa chọn".
 */

export interface CatalogItem {
  /** Mô tả (không bắt buộc), hiện kèm mã trong dropdown. */
  description?: string;
  /** Mã; '' khi mục chỉ có giá trị ràng buộc (vd vật tư người dùng thêm theo tên, chưa được cấp mã). */
  id: string;
  /** Giá trị ràng buộc theo mã ('' nếu để trống hoặc "---"). */
  link: string;
}

export interface CatalogTable {
  /** Tên bảng = tiêu đề cột mã, ví dụ "Element_ID". */
  name: string;
  /** 'file' = từ ActivityP6.xlsx (chỉ admin sửa, người dùng không thêm được); 'list' = từ list CatalogItems hoặc danh mục cục bộ (người dùng được thêm). */
  source: 'file' | 'list';
  items: CatalogItem[];
}

/** Một dòng của list CatalogItems (tên bảng đã chuẩn hoá chữ thường). */
export interface CatalogRow {
  table: string;
  code: string;
  link: string;
  description?: string;
}

export interface ReportCatalog {
  unit: string;
  /** Bảng theo tên (chữ thường). */
  tables: Record<string, CatalogTable>;
  /** Cảnh báo dữ liệu (trùng mã, ràng buộc trỏ tới mã không có…) để admin sửa file. */
  warnings: string[];
}

/** Giá trị "chưa chọn" của cột mã. */
export const CATALOG_NONE = '---';

/** Danh mục rỗng (chưa đọc được file), kèm một cảnh báo. */
export function emptyCatalog(unit: string, warning: string): ReportCatalog {
  return { unit, tables: {}, warnings: [warning] };
}

/**
 * `select:@Element_ID` -> { table: 'element_id' }; `select:@Material_ID.Material` -> { table: 'material_id', byLink: true }
 * (dropdown liệt kê cột ràng buộc "Material", chọn thì mã được tự điền). Không phải tham chiếu thì undefined.
 */
export function parseCatalogRef(option: string): { table: string; byLink: boolean } | undefined {
  const m = /^@\s*([A-Za-z0-9_]+)(\.[A-Za-z0-9_]+)?$/.exec(option.trim());
  return m ? { table: m[1].toLowerCase(), byLink: Boolean(m[2]) } : undefined;
}

function norm(v: unknown): string {
  return v === undefined || v === null ? '' : String(v).trim();
}

/** Tên bảng Activity (chữ thường dùng làm khoá, khớp `catalog` của cột trong template). */
export const ACTIVITY_TABLE = 'ActivityP6_ID';

const HEADER_ALIASES = {
  code: ['activityid', 'activity id', 'activityp6_id', 'activityp6id', 'id', 'code', 'ma'],
  unit: ['unit', 'don vi', 'donvi'],
  description: ['description', 'mo ta', 'ghi chu']
};

function headerIndex(header: unknown[], aliases: string[]): number {
  for (let c = 0; c < header.length; c++) if (aliases.indexOf(normKey(norm(header[c]))) !== -1) return c;
  return -1;
}

/**
 * Đọc các Activity của `unit` từ file ActivityP6.xlsx: sheet đầu tiên (hoặc sheet tên "ActivityP6"), dòng 1 là tiêu đề với cột
 * `ActivityID`, `Unit`, `Description` (tìm theo tên cột, thứ tự cột tuỳ ý; `Description` không bắt buộc). Chỉ lấy dòng có `Unit` trùng
 * đơn vị (không phân biệt hoa thường) hoặc `ALL`. Thiếu cột/không có Activity cho đơn vị thì trả bảng rỗng kèm cảnh báo.
 */
export function parseCatalogFromWorkbook(XLSX: typeof XLSXTypes, wb: XLSXTypes.WorkBook, unit: string): ReportCatalog {
  const catalog: ReportCatalog = { unit, tables: {}, warnings: [] };
  const sheetName = wb.SheetNames.filter((n) => n.trim().toLowerCase() === 'activityp6')[0] || wb.SheetNames[0];
  const table: CatalogTable = { name: ACTIVITY_TABLE, source: 'file', items: [] };
  catalog.tables[ACTIVITY_TABLE.toLowerCase()] = table;
  if (!sheetName) {
    catalog.warnings.push('File ActivityP6.xlsx không có sheet nào.');
    return catalog;
  }
  const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sheetName], { header: 1, defval: '', blankrows: false });
  const header = rows[0] || [];
  const codeCol = headerIndex(header, HEADER_ALIASES.code);
  const unitCol = headerIndex(header, HEADER_ALIASES.unit);
  const descCol = headerIndex(header, HEADER_ALIASES.description);
  if (codeCol < 0 || unitCol < 0) {
    catalog.warnings.push(`Sheet "${sheetName}" cần có cột ActivityID và Unit ở dòng 1.`);
    return catalog;
  }
  const target = unit.trim().toLowerCase();
  const seen = new Set<string>();
  for (let r = 1; r < rows.length; r++) {
    const id = norm(rows[r][codeCol]);
    const rowUnit = norm(rows[r][unitCol]).toLowerCase();
    if (!id || id === CATALOG_NONE || (rowUnit !== target && rowUnit !== 'all')) continue;
    if (seen.has(id.toLowerCase())) {
      catalog.warnings.push(`Trùng mã ${id} của đơn vị ${unit} trong ActivityP6.xlsx.`);
      continue;
    }
    seen.add(id.toLowerCase());
    const description = descCol >= 0 ? norm(rows[r][descCol]) : '';
    table.items.push(description ? { id, link: '', description } : { id, link: '' });
  }
  if (table.items.length === 0) catalog.warnings.push(`Chưa có Activity nào của đơn vị ${unit} trong ActivityP6.xlsx.`);
  return catalog;
}

/** Mã trong bảng `ref` (không phân biệt hoa thường); undefined nếu là "---", trống hoặc không có. */
export function findCatalogItem(catalog: ReportCatalog | undefined, ref: string, id: string): CatalogItem | undefined {
  if (!catalog || !id || id === CATALOG_NONE) return undefined;
  const table = catalog.tables[ref.toLowerCase()];
  const k = id.toLowerCase();
  return table ? table.items.filter((i) => i.id !== '' && i.id.toLowerCase() === k)[0] : undefined;
}

/** Mục có giá trị ràng buộc = `link` (không phân biệt hoa thường); undefined nếu không có. */
export function findCatalogItemByLink(catalog: ReportCatalog | undefined, ref: string, link: string): CatalogItem | undefined {
  if (!catalog || !link) return undefined;
  const table = catalog.tables[ref.toLowerCase()];
  const k = link.toLowerCase();
  return table ? table.items.filter((i) => i.link && i.link.toLowerCase() === k)[0] : undefined;
}

/** Danh sách chọn của cột tham chiếu: "---" (trừ `noNone`) + mọi mã của bảng (hoặc, với `byLink`, các giá trị ràng buộc không trùng nhau). */
export function catalogOptions(catalog: ReportCatalog | undefined, ref: string, byLink = false, noNone = false): string[] {
  const table = catalog ? catalog.tables[ref.toLowerCase()] : undefined;
  if (byLink) {
    const out: string[] = [];
    (table ? table.items : []).forEach((i) => {
      if (i.link && out.indexOf(i.link) === -1) out.push(i.link);
    });
    return out;
  }
  const ids = table ? table.items.filter((i) => i.id !== '').map((i) => i.id) : [];
  return noNone ? ids : [CATALOG_NONE].concat(ids);
}

/** Nhãn hiển thị của các mã trong dropdown: "mã – mô tả" cho mục có mô tả (mục không có mô tả giữ nguyên mã). */
export function catalogOptionLabels(catalog: ReportCatalog | undefined, ref: string): Record<string, string> | undefined {
  const table = catalog ? catalog.tables[ref.toLowerCase()] : undefined;
  if (!table) return undefined;
  const out: Record<string, string> = {};
  let any = false;
  table.items.forEach((i) => {
    if (i.id && i.description) {
      out[i.id] = `${i.id} – ${i.description}`;
      any = true;
    }
  });
  return any ? out : undefined;
}

function columnByCode(columns: TableColumn[], code: string): TableColumn | undefined {
  return columns.filter((c) => c.code.toLowerCase() === code.toLowerCase())[0];
}

/** Mã đang chọn ở cột nguồn của `col` (cột có `from=`) và giá trị ràng buộc của nó; undefined nếu chưa chọn mã có ràng buộc. */
export function linkedValue(
  columns: TableColumn[],
  col: TableColumn,
  values: Record<string, CellValue>,
  catalog: ReportCatalog | undefined
): string | undefined {
  const src = col.from ? columnByCode(columns, col.from) : undefined;
  if (!src || !src.optionsRef) return undefined;
  const cur = String(values[src.code] ?? '');
  if (src.optionsBy) return findCatalogItemByLink(catalog, src.optionsRef, cur)?.id || undefined;
  const item = findCatalogItem(catalog, src.optionsRef, cur);
  return item && item.link ? item.link : undefined;
}

/**
 * Sau khi cột `changed` đổi giá trị (sửa `values` tại chỗ): mọi cột `from=<changed>` nhận giá trị ràng buộc của mã mới
 * (mã chưa có ràng buộc hoặc mã mới thêm bằng "Khác…": cột select về "---", cột text xoá để nhập tay; mã "---" thì giữ nguyên).
 */
export function applyCatalogLinks(
  columns: TableColumn[],
  values: Record<string, CellValue>,
  changed: string,
  catalog: ReportCatalog | undefined
): void {
  const src = columnByCode(columns, changed);
  if (!src || !src.optionsRef) return;
  const current = String(values[changed] ?? '');
  if (current === CATALOG_NONE || current === '') return;
  // Mã không có trong danh mục (giá trị mới thêm bằng "Khác…") thì chưa có ràng buộc: ô ràng buộc về "---" / để trống
  const linkValue = src.optionsBy
    ? findCatalogItemByLink(catalog, src.optionsRef, current)?.id
    : findCatalogItem(catalog, src.optionsRef, current)?.link;
  columns.forEach((c) => {
    if (c.from && c.from.toLowerCase() === changed.toLowerCase()) {
      values[c.code] = linkValue || (c.type === 'select' ? CATALOG_NONE : '');
    }
  });
}

/** Khoá so sánh chống trùng: bỏ khoảng trắng thừa, không phân biệt hoa thường và dấu tiếng Việt ("Thép D32" = "thep  d32"). */
export function normKey(v: string): string {
  return v
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/gi, 'd')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Mục đã có trong bảng `ref` có cùng giá trị (theo `normKey`) với `value`: so với mã (`byLink` = false) hoặc với giá trị
 * ràng buộc (`byLink` = true). Dùng để chống trùng: có rồi thì chọn mục sẵn có thay vì tạo mới.
 */
export function findDuplicate(catalog: ReportCatalog | undefined, ref: string, byLink: boolean, value: string): CatalogItem | undefined {
  const table = catalog ? catalog.tables[ref.toLowerCase()] : undefined;
  const k = normKey(value);
  if (!table || !k) return undefined;
  return table.items.filter((i) => normKey(byLink ? i.link : i.id) === k)[0];
}

/** Gộp các dòng của list CatalogItems vào danh mục (bảng có sẵn từ file thì thêm mục, bảng chưa có thì tạo mới `source: 'list'`). */
export function mergeCatalogRows(catalog: ReportCatalog, rows: CatalogRow[]): ReportCatalog {
  const tables: Record<string, CatalogTable> = {};
  Object.keys(catalog.tables).forEach((k) => {
    tables[k] = { ...catalog.tables[k], items: catalog.tables[k].items.slice() };
  });
  rows.forEach((r) => {
    const key = r.table.toLowerCase();
    const t = tables[key] || (tables[key] = { name: r.table, source: 'list', items: [] });
    const dup = t.items.some((i) => (r.code ? normKey(i.id) === normKey(r.code) : false) || (!r.code && normKey(i.link) === normKey(r.link)));
    if (!dup) t.items.push(r.description ? { id: r.code, link: r.link, description: r.description } : { id: r.code, link: r.link });
  });
  return { ...catalog, tables };
}

/** Danh mục sau khi thêm một mục (không sửa danh mục cũ). */
export function catalogWithRow(catalog: ReportCatalog | undefined, unit: string, row: CatalogRow): ReportCatalog {
  return mergeCatalogRows(catalog || { unit, tables: {}, warnings: [] }, [row]);
}

/** Bảng chỉ admin được thêm/bớt (nằm trong ActivityP6.xlsx), dù file có đọc được hay không. */
const ADMIN_ONLY_TABLES = ['activityp6_id'];

/** Người dùng có được thêm giá trị mới ("Khác…") vào bảng này không: bảng admin-only hoặc đến từ file thì không; còn lại được. */
export function canAddToTable(catalog: ReportCatalog | undefined, ref: string): boolean {
  if (!catalog || ADMIN_ONLY_TABLES.indexOf(ref.toLowerCase()) !== -1) return false;
  const t = catalog.tables[ref.toLowerCase()];
  return !t || t.source === 'list';
}
