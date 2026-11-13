import { parseCatalogRef } from './reportCatalog';
import {
  CellValue,
  ColumnTone,
  ReportTable,
  TableColumn,
  TableRow
} from '../types/report.types';

/**
 * Bảng động: cấu trúc bảng nằm trong sheet của template (xem AGENTS.md mục 5.3).
 *
 *   Dòng 1: Title | <Tiêu đề VI> | <Tiêu đề EN> | <Nhãn ngắn> | [single]   (single = bảng một dòng cố định, không thêm/xoá dòng)
 *   Dòng 2: Summary | <Tóm tắt> | [off]                                (off = ẩn ô Tóm tắt)
 *   Dòng 3: mã cột (Org, Nos, …, IsLocked)
 *   Dòng 4: nhãn "Tiếng Việt | English"
 *   Dòng 5: spec cột: <loại>[;cờ…]
 *   Dòng 6+: dữ liệu
 *
 * Loại: text · number · date · select:A/B/C · calc:<công thức, ví dụ Nos*Hrs>
 * Cờ: key (chỉ đọc khi dòng bị khoá) · readonly (chỉ đọc ở mọi dòng; không cờ = cho sửa) · total (có dòng Tổng) · left|center|right · w=<%> ·
 *     default=<giá trị, có thể dùng {unit} {n3}> · tone=<giá trị>:danger|warning|success|info|muted,… ·
 *     carry=<công thức> (khi tạo từ báo cáo ngày trước: giá trị = công thức tính trên dòng của ngày trước, vd Daily+PrevCum) ·
 *     reset (khi tạo từ báo cáo ngày trước: về giá trị mặc định) ·
 *     from=<mã cột> (chọn mã ở cột đó thì ô này tự điền giá trị ràng buộc của mã trong danh mục và bị khoá, vd Activity theo Element;
 *       mã chưa có ràng buộc hoặc "---" thì ô cho nhập/chọn tay)
 * default có thể dùng {unit} và {n<k>}: số thứ tự tự tăng = số lớn nhất đang có trong cột + 1, đệm đủ k chữ số (vd default=ISS-{unit}-{n5} → ISS-ST01-00004).
 * Loại select có thể là select:@<tên bảng> (vd select:@ActivityP6_ID, select:@Element_ID): danh sách lấy từ ActivityP6.xlsx;
 * select:@Bảng.<cột ràng buộc> (vd select:@Material_ID.Material): dropdown liệt kê giá trị ràng buộc, chọn thì cột from= tự điền mã;
 * "Khác…" thêm giá trị vào spec (select:@Element_ID/Mới) vì file danh mục người dùng chỉ có quyền đọc.
 * Ô select luôn có mục cuối "Khác…" để nhập giá trị mới; giá trị mới được ghi lại vào spec (select:A/B/C/mới)
 * nên các ngày sau (tạo từ báo cáo ngày trước) có sẵn trong danh sách.
 */

export const LOCK_COLUMN_CODE = 'IsLocked';

export function parseColumnSpec(code: string, label: string, rawSpec: string): TableColumn {
  const spec = (rawSpec || '').trim();
  const tokens = spec.split(';').map((t) => t.trim()).filter(Boolean);
  const col: TableColumn = { code, label: label || code, spec, type: 'text' };

  tokens.forEach((tok, i) => {
    const lower = tok.toLowerCase();
    if (i === 0 && lower.startsWith('select:')) {
      col.type = 'select';
      col.options = tok.slice(7).split('/').map((o) => o.trim()).filter(Boolean);
      // select:@Element_ID[/giá trị thêm…] -> danh sách lấy từ file danh mục lúc chạy; các giá trị sau @… là giá trị người dùng
      // đã thêm bằng "Khác…" (ghi lại khi lưu để ngày sau có sẵn). Spec giữ nguyên `@…` khi lưu.
      const ref = col.options.length > 0 ? parseCatalogRef(col.options[0]) : undefined;
      if (ref) {
        col.optionsRef = ref.table;
        if (ref.byLink) col.optionsBy = 'link';
        col.options = col.options.slice(1);
      }
    } else if (i === 0 && lower.startsWith('calc:')) {
      col.type = 'calc';
      col.formula = tok.slice(5).trim();
    } else if (i === 0 && lower === 'number') {
      col.type = 'number';
    } else if (i === 0 && lower === 'date') {
      col.type = 'date';
    } else if (lower.startsWith('from=')) {
      col.from = tok.slice(5).trim();
    } else if (lower === 'reset') {
      col.reset = true;
    } else if (lower.startsWith('carry=')) {
      col.carry = tok.slice(6).trim();
    } else if (lower === 'key') {
      col.key = true;
    } else if (lower === 'readonly') {
      col.readonly = true;
    } else if (lower === 'total') {
      col.total = true;
    } else if (lower === 'left' || lower === 'center' || lower === 'right') {
      col.align = lower;
    } else if (lower.startsWith('w=')) {
      const w = parseFloat(tok.slice(2));
      if (!isNaN(w) && w > 0) col.width = w;
    } else if (lower.startsWith('default=')) {
      col.defaultValue = tok.slice(8);
    } else if (lower.startsWith('tone=')) {
      const tones: Record<string, ColumnTone> = {};
      tok.slice(5).split(',').forEach((pair) => {
        const idx = pair.lastIndexOf(':');
        if (idx < 1) return;
        const val = pair.slice(0, idx).trim();
        const tone = pair.slice(idx + 1).trim().toLowerCase();
        if (tone === 'danger' || tone === 'warning' || tone === 'success' || tone === 'info' || tone === 'muted') tones[val] = tone;
      });
      col.tones = tones;
    }
  });
  return col;
}

/** "Đơn vị | Org." -> { vi: 'Đơn vị', en: 'Org.' } */
export function splitLabel(label: string): { vi: string; en: string } {
  const idx = label.indexOf('|');
  if (idx === -1) return { vi: label.trim(), en: '' };
  return { vi: label.slice(0, idx).trim(), en: label.slice(idx + 1).trim() };
}

export function columnAlign(col: TableColumn): 'left' | 'center' | 'right' {
  if (col.align) return col.align;
  if (col.type === 'number' || col.type === 'calc') return 'right';
  if (col.type === 'select' || col.type === 'date') return 'center';
  return 'left';
}

/* ------------------------------ Ngày ------------------------------ */

/** "2026-11-12" -> "12/11/2026" (giá trị khác giữ nguyên). */
export function isoToVN(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso.trim());
  return m ? `${m[3]}/${m[2]}/${m[1]}` : iso;
}

/** "12/11/2026" (hoặc "12/11/26") -> "2026-11-12"; không đúng định dạng thì trả chuỗi rỗng (ô chọn ngày trống). */
export function vnToIso(vn: string): string {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(vn.trim());
  if (!m) return '';
  const y = m[3].length === 2 ? `20${m[3]}` : m[3];
  return `${y}-${('0' + m[2]).slice(-2)}-${('0' + m[1]).slice(-2)}`;
}

/** Số ngày của Excel (khi ô ngày bị Excel đổi thành số) -> "dd/MM/yyyy". */
export function excelSerialToVN(serial: number): string {
  const d = new Date(Math.round((serial - 25569) * 86400000));
  return `${('0' + d.getUTCDate()).slice(-2)}/${('0' + (d.getUTCMonth() + 1)).slice(-2)}/${d.getUTCFullYear()}`;
}

/** Thay {tomorrow} trong tiêu đề mục bằng ngày hôm sau của báo cáo (dd/MM/yyyy). */
export function resolveTitle(title: string, dateStr: string): string {
  if (!/\{tomorrow\}/i.test(title)) return title;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr);
  if (!m) return title.replace(/\s*[–-]?\s*\{tomorrow\}/gi, '');
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3] + 1));
  const vn = `${('0' + d.getUTCDate()).slice(-2)}/${('0' + (d.getUTCMonth() + 1)).slice(-2)}/${d.getUTCFullYear()}`;
  return title.replace(/\{tomorrow\}/gi, vn);
}

/* ------------------------ Danh sách chọn mở rộng ------------------------ */

/** Giá trị mới của ô select không được chứa dấu phân tách của spec. */
export function sanitizeOption(v: string): string {
  return v.replace(/[/;]/g, '-').trim();
}

/** Ghi lại spec với danh sách chọn mới: "select:A/B;center" -> "select:A/B/C;center". */
export function specWithOptions(spec: string, options: string[]): string {
  const tokens = spec.split(';');
  tokens[0] = `select:${options.join('/')}`;
  return tokens.join(';');
}

/** Thêm một giá trị mới vào danh sách chọn của cột (cập nhật cả `options` lẫn `spec`). */
export function columnWithOption(col: TableColumn, value: string): TableColumn {
  const options = col.options || [];
  if (options.some((o) => o.toLowerCase() === value.toLowerCase())) return col;
  const next = options.concat(value);
  // Cột danh mục: giữ `@bảng` ở đầu danh sách trong spec
  const refToken = col.optionsRef ? col.spec.split(';')[0].slice(7).split('/')[0] : undefined;
  return { ...col, options: next, spec: specWithOptions(col.spec, refToken ? [refToken].concat(next) : next) };
}

/* ---------------------------- Công thức ---------------------------- */

type Token = { t: 'num'; v: number } | { t: 'id'; v: string } | { t: 'op'; v: string };

function tokenize(src: string): Token[] {
  const out: Token[] = [];
  const re = /^\s*(?:(\d+(?:\.\d+)?)|([A-Za-z_][A-Za-z0-9_]*)|([()+\-*/,]))/;
  let rest = src;
  while (rest.trim() !== '') {
    const m = re.exec(rest);
    if (!m) throw new Error('Công thức không hợp lệ');
    if (m[1] !== undefined) out.push({ t: 'num', v: parseFloat(m[1]) });
    else if (m[2] !== undefined) out.push({ t: 'id', v: m[2] });
    else out.push({ t: 'op', v: m[3] });
    rest = rest.slice(m[0].length);
  }
  return out;
}

class FormulaParser {
  private pos = 0;

  constructor(private readonly tokens: Token[], private readonly getValue: (code: string) => number) {}

  public parse(): number {
    const v = this.expr();
    if (this.pos < this.tokens.length) throw new Error('Công thức không hợp lệ');
    return v;
  }

  private eatOp(v: string): boolean {
    const t = this.tokens[this.pos];
    if (t && t.t === 'op' && t.v === v) {
      this.pos += 1;
      return true;
    }
    return false;
  }

  private expr(): number {
    let v = this.term();
    for (;;) {
      if (this.eatOp('+')) v += this.term();
      else if (this.eatOp('-')) v -= this.term();
      else return v;
    }
  }

  private term(): number {
    let v = this.factor();
    for (;;) {
      if (this.eatOp('*')) v *= this.factor();
      else if (this.eatOp('/')) {
        const d = this.factor();
        v = d === 0 ? 0 : v / d;
      } else return v;
    }
  }

  private factor(): number {
    if (this.eatOp('-')) return -this.factor();
    if (this.eatOp('(')) {
      const v = this.expr();
      if (!this.eatOp(')')) throw new Error('Thiếu )');
      return v;
    }
    const t = this.tokens[this.pos];
    if (!t) throw new Error('Công thức thiếu toán hạng');
    this.pos += 1;
    if (t.t === 'num') return t.v;
    if (t.t === 'id') {
      if (this.eatOp('(')) {
        const arg = this.expr();
        if (!this.eatOp(')')) throw new Error('Thiếu )');
        if (t.v.toLowerCase() === 'round') return Math.round(arg);
        throw new Error('Hàm không hỗ trợ');
      }
      return this.getValue(t.v);
    }
    throw new Error('Công thức không hợp lệ');
  }
}

/** Tính công thức +,-,*,/, ngoặc, round(); chia cho 0 => 0. Mã cột không phân biệt hoa thường. */
export function evaluateFormula(formula: string, getValue: (code: string) => number): number {
  try {
    const result = new FormulaParser(tokenize(formula), getValue).parse();
    return isFinite(result) ? result : 0;
  } catch {
    return 0;
  }
}

function numberOf(v: CellValue): number {
  return typeof v === 'number' && isFinite(v) ? v : 0;
}

/** Tính lại các cột calc của một dòng (kết quả làm tròn 2 số lẻ). */
export function recalcRow(columns: TableColumn[], row: TableRow): TableRow {
  if (!columns.some((c) => c.type === 'calc')) return row;
  const values = { ...row.values };
  const byCode: Record<string, string> = {};
  columns.forEach((c) => {
    byCode[c.code.toLowerCase()] = c.code;
  });
  columns.forEach((c) => {
    if (c.type !== 'calc' || !c.formula) return;
    const v = evaluateFormula(c.formula, (code) => numberOf(values[byCode[code.toLowerCase()]]));
    values[c.code] = Math.round(v * 100) / 100;
  });
  return { ...row, values };
}

export function recalcTable(table: ReportTable): ReportTable {
  return { ...table, rows: table.rows.map((r) => recalcRow(table.columns, r)) };
}

/* ------------------------------ Dòng ------------------------------ */

/** Số thứ tự kế tiếp cho mẫu có {n<k>}: số lớn nhất đang có ở cột (khớp phần trước/sau {n<k>}) + 1, đệm k chữ số. */
function expandDefault(tpl: string, unit: string, table: ReportTable, code: string): string {
  const withUnit = tpl.replace(/\{unit\}/gi, unit);
  const m = /\{n(\d*)\}/i.exec(withUnit);
  if (!m) return withUnit;
  const prefix = withUnit.slice(0, m.index);
  const suffix = withUnit.slice(m.index + m[0].length);
  let max = 0;
  table.rows.forEach((r) => {
    const v = String(r.values[code] ?? '');
    if (v.length > prefix.length + suffix.length && v.startsWith(prefix) && v.endsWith(suffix)) {
      const n = parseInt(v.slice(prefix.length, v.length - suffix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  const width = m[1] ? parseInt(m[1], 10) : 0;
  return prefix + ('0'.repeat(width) + (max + 1)).slice(-Math.max(width, String(max + 1).length)) + suffix;
}

export function createRow(table: ReportTable, unit: string): TableRow {
  const values: Record<string, CellValue> = {};
  table.columns.forEach((c) => {
    if (c.type === 'calc') {
      values[c.code] = 0;
    } else if (c.type === 'number') {
      const n = c.defaultValue !== undefined ? parseFloat(c.defaultValue) : 0;
      values[c.code] = isNaN(n) ? 0 : n;
    } else if (c.type === 'date') {
      values[c.code] = c.defaultValue !== undefined ? c.defaultValue : '';
    } else if (c.type === 'select') {
      values[c.code] =
        c.defaultValue !== undefined ? c.defaultValue : c.options && c.options.length ? c.options[0] : '';
    } else {
      values[c.code] = c.defaultValue !== undefined ? expandDefault(c.defaultValue, unit, table, c.code) : '';
    }
  });
  return recalcRow(table.columns, {
    id: `${table.key}-${Date.now()}`,
    isLocked: false,
    values
  });
}

export function columnTotal(table: ReportTable, col: TableColumn): number {
  return table.rows.reduce((acc, r) => acc + numberOf(r.values[col.code]), 0);
}

export function hasTotals(table: ReportTable): boolean {
  return table.columns.some((c) => c.total);
}

/** Chỉ số cột đầu tiên có Tổng (ô nhãn "Tổng | Total" trải qua các cột trước nó). */
export function firstTotalIndex(table: ReportTable): number {
  const i = table.columns.findIndex((c) => c.total);
  return i < 1 ? 1 : i;
}

/** Dòng có coi là trống không (không có ô nhập nào có nội dung). */
export function isEmptyRow(columns: TableColumn[], values: Record<string, CellValue>): boolean {
  return columns.every((c) => {
    if (c.type === 'calc') return true;
    const v = values[c.code];
    return v === undefined || v === '';
  });
}

export function formatNumber(v: CellValue): string {
  return typeof v === 'number' ? v.toLocaleString('vi-VN') : '';
}

/** Màu tự động theo ý nghĩa của giá trị trạng thái (không cần khai báo trong template). */
const AUTO_TONES: Record<string, ColumnTone> = {
  xong: 'success', 'hoàn thành': 'success', 'hoàn tất': 'success', 'đạt': 'success', 'vượt': 'success',
  active: 'success', 'hoạt động': 'success', done: 'success',
  'đóng': 'muted', closed: 'muted', 'huỷ': 'muted', 'hủy': 'muted', cancelled: 'muted',
  'đang xử lý': 'info', 'đang thực hiện': 'info', 'in progress': 'info',
  'mở': 'warning', open: 'warning', pending: 'warning', 'chờ': 'warning', 'chờ việc': 'warning', 'bảo trì': 'warning',
  'hỏng': 'danger', 'chậm': 'danger', 'quá hạn': 'danger', overdue: 'danger',
  // Mức độ
  cao: 'danger', high: 'danger', 'trung bình': 'warning', medium: 'warning', 'thấp': 'success', low: 'success'
};

/** Màu của ô select: template khai báo (tone=…) được ưu tiên, nếu không thì theo từ khoá trạng thái chung. */
export function toneOf(col: TableColumn, value: CellValue): ColumnTone | undefined {
  if (typeof value !== 'string') return undefined;
  if (col.tones && col.tones[value]) return col.tones[value];
  if (col.type !== 'select') return undefined;
  return AUTO_TONES[value.trim().toLowerCase()];
}

/* ------------------- Tạo báo cáo từ báo cáo ngày trước ------------------- */

/**
 * Chuyển số liệu sang ngày mới: cột `carry=<công thức>` nhận kết quả tính trên dòng của ngày trước
 * (vd Luỹ kế hôm qua = Thực hiện hôm nay + Luỹ kế hôm qua); cột `reset` về giá trị mặc định.
 */
export function carryForwardTable(table: ReportTable): ReportTable {
  const carryCols = table.columns.filter((c) => c.carry);
  const resetCols = table.columns.filter((c) => c.reset);
  if (carryCols.length === 0 && resetCols.length === 0) return table;
  const byCode: Record<string, string> = {};
  table.columns.forEach((c) => {
    byCode[c.code.toLowerCase()] = c.code;
  });
  const rows = table.rows.map((row) => {
    const prev = row.values;
    const values = { ...prev };
    carryCols.forEach((c) => {
      const v = evaluateFormula(c.carry as string, (code) => numberOf(prev[byCode[code.toLowerCase()]]));
      values[c.code] = Math.round(v * 100) / 100;
    });
    resetCols.forEach((c) => {
      if (c.type === 'number') {
        const n = c.defaultValue !== undefined ? parseFloat(c.defaultValue) : 0;
        values[c.code] = isNaN(n) ? 0 : n;
      } else {
        values[c.code] = c.defaultValue !== undefined ? c.defaultValue : '';
      }
    });
    return recalcRow(table.columns, { ...row, values });
  });
  return { ...table, rows };
}

/** Bề rộng tối thiểu (px) của cột để không bị bóp chữ; bảng rộng hơn khung thì cuộn ngang. */
export function columnMinPx(col: TableColumn): number {
  // Cột mã danh mục (CV-ST01-00001, RES-HUM-00001…) cần đủ rộng để hiện trọn mã cùng mũi tên dropdown
  const floor = col.optionsRef
    ? 170
    : col.type === 'date'
    ? 140
    : col.type === 'select'
    ? 140
    : col.type === 'number' || col.type === 'calc'
    ? 110
    : 100;
  return Math.max(floor, Math.round((col.width || 14) * 9));
}
