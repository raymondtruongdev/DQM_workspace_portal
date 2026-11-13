import type * as XLSXTypes from 'xlsx';
import { CatalogRef, ReportDoc, ReportTable, TableColumn, CellValue } from '../types/report.types';

/**
 * Danh mục Activity P6 + cấu kiện của một đơn vị, đọc từ file ActivityP6.xlsx (xem AGENTS.md mục 5.3).
 *
 * File có một sheet cho mỗi đơn vị (tên sheet = mã đơn vị, ví dụ ST01). Trong sheet có hai bảng xếp dọc,
 * nhận diện bằng TÊN CỘT ở dòng tiêu đề (không dựa vào đối tượng Table của Excel):
 *   - bảng Activity: ActivityP6_ID | Description | Status
 *   - bảng Element:  ActivityP6_ID | Element_ID | Description | Status   (nhận ra nhờ có cột Element_ID)
 * Mỗi bảng đọc từ dưới dòng tiêu đề đến dòng trống hoặc dòng tiêu đề của bảng kia. Status: active (mặc định) | inactive.
 * Mã "---" (dòng "Chưa định nghĩa") là giá trị "chưa chọn", không đưa vào danh sách; cấu kiện có Activity "---" hoặc trống = chưa gán Activity.
 */

export interface CatalogActivity {
  id: string;
  description: string;
  active: boolean;
}

export interface CatalogElement {
  id: string;
  /** Mã Activity của cấu kiện ('' nếu chưa gán: để trống hoặc "---"). */
  activityId: string;
  description: string;
  active: boolean;
}

export interface ReportCatalog {
  unit: string;
  activities: CatalogActivity[];
  elements: CatalogElement[];
  /** Cảnh báo dữ liệu (trùng mã, cấu kiện trỏ tới Activity không có…) để admin sửa file. */
  warnings: string[];
}

/** Giá trị "chưa chọn" của cột mã. */
export const CATALOG_NONE = '---';

/** `select:@ActivityID` / `select:@Element` -> loại danh mục (không phân biệt hoa thường); không phải tham chiếu thì undefined. */
export function parseCatalogRef(option: string): CatalogRef | undefined {
  const m = /^@\s*(activity(?:id)?|element)$/i.exec(option.trim());
  if (!m) return undefined;
  return m[1].toLowerCase() === 'element' ? 'element' : 'activity';
}

function norm(v: unknown): string {
  return v === undefined || v === null ? '' : String(v).trim();
}

function isActive(status: string): boolean {
  return status.toLowerCase() !== 'inactive';
}

/** Đọc danh mục của `unit` từ nội dung file ActivityP6.xlsx; không có sheet của đơn vị thì trả danh mục rỗng kèm cảnh báo. */
export function parseCatalogFromWorkbook(XLSX: typeof XLSXTypes, wb: XLSXTypes.WorkBook, unit: string): ReportCatalog {
  const catalog: ReportCatalog = { unit, activities: [], elements: [], warnings: [] };
  const sheetName = wb.SheetNames.filter((n) => n.trim().toLowerCase() === unit.trim().toLowerCase())[0];
  if (!sheetName) {
    catalog.warnings.push(`Chưa có sheet "${unit}" trong file danh mục.`);
    return catalog;
  }
  const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sheetName], { header: 1, defval: '', blankrows: true });

  /** Chỉ số cột theo tên header (không phân biệt hoa thường). */
  const headerMap = (row: unknown[]): Record<string, number> => {
    const map: Record<string, number> = {};
    row.forEach((cell, i) => {
      const key = norm(cell).toLowerCase();
      if (key && map[key] === undefined) map[key] = i;
    });
    return map;
  };

  type Section = { kind: 'activity' | 'element'; start: number; cols: Record<string, number> };
  const sections: Section[] = [];
  rows.forEach((row, r) => {
    const cols = headerMap(row);
    if (cols.activityp6_id === undefined) return;
    if (cols.element_id !== undefined) sections.push({ kind: 'element', start: r, cols });
    else if (cols.description !== undefined) sections.push({ kind: 'activity', start: r, cols });
  });

  sections.forEach((sec, si) => {
    const limit = si + 1 < sections.length ? sections[si + 1].start : rows.length;
    for (let r = sec.start + 1; r < limit; r++) {
      const row = rows[r];
      if (row.every((c) => norm(c) === '')) break; // dòng trống = hết bảng
      // "---" = chưa gán Activity (dòng placeholder "Chưa định nghĩa" hoặc cấu kiện chưa có Activity)
      const rawActivity = norm(row[sec.cols.activityp6_id]);
      const activityId = rawActivity === CATALOG_NONE ? '' : rawActivity;
      const description = sec.cols.description !== undefined ? norm(row[sec.cols.description]) : '';
      const active = isActive(sec.cols.status !== undefined ? norm(row[sec.cols.status]) : '');
      if (sec.kind === 'activity') {
        if (activityId) catalog.activities.push({ id: activityId, description, active });
      } else {
        const id = norm(row[sec.cols.element_id]);
        if (id && id !== CATALOG_NONE) catalog.elements.push({ id, activityId, description, active });
      }
    }
  });

  const seenActivity = new Set<string>();
  catalog.activities.forEach((a) => {
    const k = a.id.toLowerCase();
    if (seenActivity.has(k)) catalog.warnings.push(`Trùng mã Activity ${a.id}.`);
    seenActivity.add(k);
  });
  const seenElement = new Set<string>();
  catalog.elements.forEach((e) => {
    const k = e.id.toLowerCase();
    if (seenElement.has(k)) catalog.warnings.push(`Trùng mã cấu kiện ${e.id}.`);
    seenElement.add(k);
    if (e.activityId && !seenActivity.has(e.activityId.toLowerCase())) {
      catalog.warnings.push(`Cấu kiện ${e.id} trỏ tới Activity ${e.activityId} không có trong danh mục.`);
    }
  });
  if (catalog.activities.length === 0) catalog.warnings.push(`Sheet "${sheetName}" chưa có bảng Activity (cột ActivityP6_ID, Description).`);
  return catalog;
}

/** Cấu kiện theo mã (không phân biệt hoa thường, kể cả đã inactive để báo cáo cũ vẫn tra được). */
export function findElement(catalog: ReportCatalog | undefined, id: string): CatalogElement | undefined {
  if (!catalog || !id || id === CATALOG_NONE) return undefined;
  const k = id.toLowerCase();
  return catalog.elements.filter((e) => e.id.toLowerCase() === k)[0];
}

/** Danh sách chọn của cột tham chiếu: "---" + mã đang dùng (active), kèm nhãn "mã – mô tả" để hiển thị. */
export function catalogOptions(
  catalog: ReportCatalog | undefined,
  ref: CatalogRef
): { options: string[]; labels: Record<string, string> } {
  const options = [CATALOG_NONE];
  const labels: Record<string, string> = {};
  if (catalog) {
    const items = ref === 'element' ? catalog.elements : catalog.activities;
    items.forEach((it) => {
      if (!it.active || options.indexOf(it.id) !== -1) return;
      options.push(it.id);
      if (it.description) labels[it.id] = `${it.id} – ${it.description}`;
    });
  }
  return { options, labels };
}

/** Activity theo mã (kể cả đã inactive). */
export function findActivity(catalog: ReportCatalog | undefined, id: string): CatalogActivity | undefined {
  if (!catalog || !id || id === CATALOG_NONE) return undefined;
  const k = id.toLowerCase();
  return catalog.activities.filter((a) => a.id.toLowerCase() === k)[0];
}

/** Mô tả của một mã trong danh mục; undefined nếu mã là "---", trống hoặc không có trong danh mục. */
export function catalogDescription(catalog: ReportCatalog | undefined, ref: CatalogRef, id: string): string | undefined {
  const item = ref === 'element' ? findElement(catalog, id) : findActivity(catalog, id);
  return item ? item.description : undefined;
}

function columnByCode(columns: TableColumn[], code: string): TableColumn | undefined {
  return columns.filter((c) => c.code.toLowerCase() === code.toLowerCase())[0];
}

/** Cột mã (select:@…) mà cột `desc=` đang lấy mô tả. */
function descSource(columns: TableColumn[], col: TableColumn): TableColumn | undefined {
  const src = col.desc ? columnByCode(columns, col.desc) : undefined;
  return src && src.optionsRef ? src : undefined;
}

/** Ô `desc=` bị khoá khi mã đang chọn có mô tả trong danh mục; mã "---" (hoặc chưa có mô tả) thì để người dùng nhập tay. */
export function isDescLocked(
  columns: TableColumn[],
  col: TableColumn,
  values: Record<string, CellValue>,
  catalog: ReportCatalog | undefined
): boolean {
  const src = descSource(columns, col);
  if (!src || !src.optionsRef) return false;
  return Boolean(catalogDescription(catalog, src.optionsRef, String(values[src.code] ?? '')));
}

/**
 * Sau khi cột `changed` đổi giá trị (sửa `values` tại chỗ):
 *  - cột `from=<changed>` (Activity theo cấu kiện) nhận Activity của cấu kiện, chưa gán thì "---" (cấu kiện "---" thì giữ nguyên);
 *  - cột `desc=<cột vừa đổi>` nhận mô tả của mã mới; mã "---" thì xoá để người dùng nhập tay.
 */
export function applyCatalogLinks(
  columns: TableColumn[],
  values: Record<string, CellValue>,
  changed: string,
  catalog: ReportCatalog | undefined
): void {
  const touched = [changed.toLowerCase()];
  columns.forEach((c) => {
    if (c.from && c.from.toLowerCase() === changed.toLowerCase()) {
      const el = findElement(catalog, String(values[changed] ?? ''));
      if (el) {
        values[c.code] = el.activityId || CATALOG_NONE;
        touched.push(c.code.toLowerCase());
      }
    }
  });
  columns.forEach((c) => {
    if (!c.desc || touched.indexOf(c.desc.toLowerCase()) === -1) return;
    const src = descSource(columns, c);
    if (!src || !src.optionsRef) return;
    values[c.code] = catalogDescription(catalog, src.optionsRef, String(values[src.code] ?? '')) || '';
  });
}

/** Khi mở báo cáo: ô `desc=` còn trống mà mã đã có mô tả trong danh mục (vd dòng cố định của template) thì điền mô tả. */
export function fillCatalogDescriptions(table: ReportTable, catalog: ReportCatalog | undefined): ReportTable {
  const descCols = table.columns.filter((c) => descSource(table.columns, c));
  if (!catalog || descCols.length === 0) return table;
  let changed = false;
  const rows = table.rows.map((row) => {
    let values = row.values;
    descCols.forEach((c) => {
      const src = descSource(table.columns, c);
      if (!src || !src.optionsRef || String(values[c.code] ?? '').trim() !== '') return;
      const d = catalogDescription(catalog, src.optionsRef, String(values[src.code] ?? ''));
      if (d) {
        if (values === row.values) values = { ...row.values };
        values[c.code] = d;
        changed = true;
      }
    });
    return values === row.values ? row : { ...row, values };
  });
  return changed ? { ...table, rows } : table;
}

/** Áp dụng `fillCatalogDescriptions` cho mọi mục bảng của báo cáo. */
export function fillDocCatalogDescriptions(doc: ReportDoc, catalog: ReportCatalog | undefined): ReportDoc {
  if (!catalog) return doc;
  return { ...doc, sections: doc.sections.map((s) => (s.kind === 'table' ? fillCatalogDescriptions(s, catalog) : s)) };
}
