import { ReportTable, TableColumn, TableRow } from '../types/report.types';

/** Một dòng ngày trước đang thiếu ở bảng hiện tại, kèm mô tả ngắn để người dùng nhận ra. */
export interface RestorableRow {
  id: string;
  label: string;
  row: TableRow;
}

/** Mô tả ngắn của dòng (tối đa `max` ô có giá trị thật, bỏ `---` và số 0) để người dùng nhận ra dòng cần khôi phục. */
export function describeRow(columns: TableColumn[], row: TableRow, max = 4): string {
  const parts: string[] = [];
  for (const col of columns) {
    const v = row.values[col.code];
    if (v === undefined || v === null) continue;
    const text = String(v).trim();
    if (text === '' || text === '---' || text === '0') continue;
    parts.push(text);
    if (parts.length >= max) break;
  }
  return parts.length > 0 ? parts.join(' · ') : '(dòng trống)';
}

/** Dòng có ở ngày trước (theo `id`) nhưng không còn ở bảng hiện tại, theo thứ tự ngày trước. */
export function missingRows(current: ReportTable, previous: ReportTable): RestorableRow[] {
  const have = new Set(current.rows.map((r) => r.id));
  return previous.rows
    .filter((r) => !have.has(r.id))
    .map((r) => ({ id: r.id, label: describeRow(previous.columns, r), row: r }));
}

/**
 * Chèn các dòng đã chọn vào bảng hiện tại, đặt ngay sau dòng đứng trước nó ở ngày trước nếu dòng đó còn
 * (không còn thì chèn sau dòng gần nhất còn lại, hoặc đầu bảng). Dòng đã có (cùng `id`) thì bỏ qua.
 */
export function insertRestoredRows(current: ReportTable, previous: ReportTable, selectedIds: string[]): ReportTable {
  const want = new Set(selectedIds);
  const rows = current.rows.slice();
  previous.rows.forEach((pr, pi) => {
    if (!want.has(pr.id) || rows.some((r) => r.id === pr.id)) return;
    let at = 0;
    for (let k = pi - 1; k >= 0; k--) {
      const found = rows.findIndex((r) => r.id === previous.rows[k].id);
      if (found >= 0) {
        at = found + 1;
        break;
      }
    }
    rows.splice(at, 0, pr);
  });
  return { ...current, rows };
}
