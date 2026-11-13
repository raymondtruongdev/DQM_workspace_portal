import type * as XLSXTypes from 'xlsx';
import { TableColumn, CellValue } from '../types/report.types';

/**
 * Danh mục dùng chung của một đơn vị, đọc từ file ActivityP6.xlsx (xem AGENTS.md mục 5.3).
 *
 * File có một sheet cho mỗi đơn vị (tên sheet = mã đơn vị, ví dụ ST01). Trong sheet, các bảng danh mục xếp NGANG
 * (cách nhau ít nhất một cột trống), dòng 1 là tiêu đề, dữ liệu từ dòng 2 trở xuống đến ô mã đầu tiên bị trống:
 *   - cột 1 = MÃ (ID); tên tiêu đề cột 1 chính là tên bảng, template tham chiếu bằng `select:@<tên bảng>`;
 *   - cột 2 (không bắt buộc) = giá trị RÀNG BUỘC của mã đó (Element_ID → ActivityP6_ID, ResHuman_ID → Category…).
 *     Cột ở template khai báo `from=<cột mã>` sẽ tự điền giá trị này khi người dùng chọn mã.
 * Không có mô tả hay trạng thái: muốn bỏ một mã thì xoá dòng; báo cáo cũ vẫn giữ nguyên giá trị đã lưu.
 * Mã "---" luôn có sẵn trong dropdown làm giá trị "chưa chọn"; mã không có ràng buộc để trống ô cột 2.
 */

export interface CatalogItem {
  id: string;
  /** Giá trị ràng buộc theo mã ('' nếu để trống hoặc "---"). */
  link: string;
}

export interface CatalogTable {
  /** Tên bảng = tiêu đề cột mã, ví dụ "Element_ID". */
  name: string;
  /** Tiêu đề cột ràng buộc (vd "ActivityP6_ID"); undefined nếu bảng chỉ có cột mã. */
  linkName?: string;
  items: CatalogItem[];
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

/** `select:@Element_ID` -> tên bảng (chữ thường) "element_id"; không phải tham chiếu thì undefined. */
export function parseCatalogRef(option: string): string | undefined {
  const m = /^@\s*([A-Za-z0-9_]+)$/.exec(option.trim());
  return m ? m[1].toLowerCase() : undefined;
}

function norm(v: unknown): string {
  return v === undefined || v === null ? '' : String(v).trim();
}

/** Đọc danh mục của `unit` từ nội dung file ActivityP6.xlsx; không có sheet của đơn vị thì trả danh mục rỗng kèm cảnh báo. */
export function parseCatalogFromWorkbook(XLSX: typeof XLSXTypes, wb: XLSXTypes.WorkBook, unit: string): ReportCatalog {
  const sheetName = wb.SheetNames.filter((n) => n.trim().toLowerCase() === unit.trim().toLowerCase())[0];
  if (!sheetName) return emptyCatalog(unit, `Chưa có sheet "${unit}" trong file danh mục.`);
  const catalog: ReportCatalog = { unit, tables: {}, warnings: [] };
  const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sheetName], { header: 1, defval: '', blankrows: true });
  const header = rows[0] || [];

  // Bảng bắt đầu ở ô tiêu đề không trống có ô bên trái trống; rộng tối đa 2 cột (mã + ràng buộc)
  for (let c = 0; c < header.length; c++) {
    const name = norm(header[c]);
    if (!name || (c > 0 && norm(header[c - 1]) !== '')) continue;
    const linkName = norm(header[c + 1]) || undefined;
    const table: CatalogTable = { name, linkName, items: [] };
    const seen = new Set<string>();
    for (let r = 1; r < rows.length; r++) {
      const id = norm(rows[r][c]);
      if (!id) break; // ô mã trống = hết bảng
      if (id === CATALOG_NONE) continue;
      const raw = linkName ? norm(rows[r][c + 1]) : '';
      if (seen.has(id.toLowerCase())) catalog.warnings.push(`Trùng mã ${id} trong bảng ${name}.`);
      seen.add(id.toLowerCase());
      table.items.push({ id, link: raw === CATALOG_NONE ? '' : raw });
    }
    catalog.tables[name.toLowerCase()] = table;
  }

  const tableNames = Object.keys(catalog.tables);
  if (tableNames.length === 0) catalog.warnings.push(`Sheet "${sheetName}" chưa có bảng danh mục nào (tiêu đề ở dòng 1).`);

  // Ràng buộc trỏ tới bảng khác (tên cột ràng buộc trùng tên một bảng) thì mã phải có trong bảng đó
  tableNames.forEach((key) => {
    const t = catalog.tables[key];
    const target = t.linkName ? catalog.tables[t.linkName.toLowerCase()] : undefined;
    if (!target) return;
    const ids = new Set(target.items.map((i) => i.id.toLowerCase()));
    t.items.forEach((i) => {
      if (i.link && !ids.has(i.link.toLowerCase())) {
        catalog.warnings.push(`${i.id} (${t.name}) trỏ tới ${i.link} không có trong bảng ${target.name}.`);
      }
    });
  });
  return catalog;
}

/** Mã trong bảng `ref` (không phân biệt hoa thường); undefined nếu là "---", trống hoặc không có. */
export function findCatalogItem(catalog: ReportCatalog | undefined, ref: string, id: string): CatalogItem | undefined {
  if (!catalog || !id || id === CATALOG_NONE) return undefined;
  const table = catalog.tables[ref.toLowerCase()];
  const k = id.toLowerCase();
  return table ? table.items.filter((i) => i.id.toLowerCase() === k)[0] : undefined;
}

/** Danh sách chọn của cột tham chiếu: "---" + mọi mã của bảng. */
export function catalogOptions(catalog: ReportCatalog | undefined, ref: string): string[] {
  const table = catalog ? catalog.tables[ref.toLowerCase()] : undefined;
  return [CATALOG_NONE].concat(table ? table.items.map((i) => i.id) : []);
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
  const item = findCatalogItem(catalog, src.optionsRef, String(values[src.code] ?? ''));
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
  const item = findCatalogItem(catalog, src.optionsRef, current);
  columns.forEach((c) => {
    if (c.from && c.from.toLowerCase() === changed.toLowerCase()) {
      values[c.code] = (item && item.link) || (c.type === 'select' ? CATALOG_NONE : '');
    }
  });
}
