import {
  CellValue,
  ColumnTone,
  ReportTable,
  TableColumn,
  TableRow
} from '../types/report.types';

/**
 * Bảng động: định nghĩa cột (kiểu, danh mục, công thức, mặc định, luỹ kế…) đến từ template JSON, ghép với nhãn của báo cáo
 * theo `id` cột (xem reportJson.ts). Ô select dùng danh mục (list CatalogItems) hoặc danh sách cố định `options`;
 * mọi ô select đều có mục cuối "Khác…" để thêm giá trị mới (danh mục thì ghi vào list dùng chung).
 * Mã tự tăng: cột text có default chứa {n<k>} (vd EVT-{unit}-{n5}) = số lớn nhất đang dùng + 1, đệm k chữ số.
 */

/** "Đơn vị | Org." -> { vi: 'Đơn vị', en: 'Org.' } */
export function splitLabel(label: string): { vi: string; en: string } {
  const idx = label.indexOf('|');
  if (idx === -1) return { vi: label.trim(), en: '' };
  return { vi: label.slice(0, idx).trim(), en: label.slice(idx + 1).trim() };
}

export function columnAlign(col: TableColumn): 'left' | 'center' | 'right' {
  if (col.align) return col.align;
  if (col.optionsBy) return 'left'; // chọn theo tên (vật tư, thiết bị…): chữ dài, canh trái cho dễ đọc
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

/** Giá trị mới của ô select: bỏ ký tự phân tách danh sách (/ và ;) và khoảng trắng thừa. */
export function sanitizeOption(v: string): string {
  return v.replace(/[/;]/g, '-').trim();
}

/** Thêm một giá trị mới vào danh sách chọn cố định của cột. */
export function columnWithOption(col: TableColumn, value: string): TableColumn {
  const options = col.options || [];
  if (options.some((o) => o.toLowerCase() === value.toLowerCase())) return col;
  return { ...col, options: options.concat(value) };
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

/** Tách mẫu mã tự tăng: phần trước, phần sau {n<k>} (đã thay {unit}) và độ rộng k; undefined nếu mẫu không có {n<k>}. */
function splitAutoTemplate(tpl: string, unit: string): { prefix: string; suffix: string; width: number } | undefined {
  const withUnit = tpl.replace(/\{unit\}/gi, unit);
  const m = /\{n(\d*)\}/i.exec(withUnit);
  if (!m) return undefined;
  return { prefix: withUnit.slice(0, m.index), suffix: withUnit.slice(m.index + m[0].length), width: m[1] ? parseInt(m[1], 10) : 0 };
}

/** Số lớn nhất đang có trong cột `code` theo mẫu `tpl` (0 nếu chưa có dòng nào khớp). */
function scanMaxAutoId(tpl: string, unit: string, table: ReportTable, code: string): number {
  const parts = splitAutoTemplate(tpl, unit);
  if (!parts) return 0;
  let max = 0;
  table.rows.forEach((r) => {
    const v = String(r.values[code] ?? '');
    if (v.length > parts.prefix.length + parts.suffix.length && v.startsWith(parts.prefix) && v.endsWith(parts.suffix)) {
      const n = parseInt(v.slice(parts.prefix.length, v.length - parts.suffix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return max;
}

/**
 * Số thứ tự kế tiếp cho mẫu có {n<k>}: lớn hơn cả số đang có ở cột lẫn `floor` (số đã dùng ở các báo cáo khác của đơn vị,
 * lấy từ sổ ReportRegistry), đệm k chữ số.
 */
function expandDefault(tpl: string, unit: string, table: ReportTable, code: string, floor = 0): string {
  const parts = splitAutoTemplate(tpl, unit);
  if (!parts) return tpl.replace(/\{unit\}/gi, unit);
  const next = Math.max(floor, scanMaxAutoId(tpl, unit, table, code)) + 1;
  return parts.prefix + ('0'.repeat(parts.width) + next).slice(-Math.max(parts.width, String(next).length)) + parts.suffix;
}

/** Cột mã tự tăng: cột text có default chứa {n<k>} (vd default=EVT-{unit}-{n5}). */
export function autoIdColumns(table: ReportTable): TableColumn[] {
  return table.columns.filter((c) => c.type === 'text' && c.defaultValue !== undefined && /\{n\d*\}/i.test(c.defaultValue));
}

/** Khoá bộ đếm của một cột mã tự tăng trong sổ ReportRegistry: "events.eventid". */
export function autoIdKey(table: ReportTable, col: TableColumn): string {
  return `${table.key}.${col.code}`.toLowerCase();
}

/** Bộ đếm cần ghi vào sổ: với mỗi cột mã tự tăng, số lớn nhất của báo cáo này và của các báo cáo trước (`floor`). */
export function collectAutoIdCounters(
  sections: { kind: string }[],
  unit: string,
  floor: Record<string, number> = {}
): Record<string, number> {
  const out: Record<string, number> = { ...floor };
  sections.forEach((s) => {
    if (s.kind !== 'table') return;
    const t = s as ReportTable;
    autoIdColumns(t).forEach((c) => {
      const k = autoIdKey(t, c);
      out[k] = Math.max(out[k] || 0, scanMaxAutoId(c.defaultValue as string, unit, t, c.code));
    });
  });
  return out;
}

/**
 * Điền mã tự tăng còn trống hoặc còn nguyên mẫu "{n5}" (dòng tạo bởi bản cũ): số tiếp theo lớn hơn cả số đã dùng ở
 * các báo cáo khác (`floor`) lẫn số đang có trong bảng.
 */
export function fillAutoIds(table: ReportTable, unit: string, floor: Record<string, number> = {}): ReportTable {
  let current = table;
  autoIdColumns(table).forEach((col) => {
    let used = floor[autoIdKey(table, col)] || 0;
    const rows = current.rows.map((r) => {
      const v = String(r.values[col.code] ?? '');
      if (v !== '' && !/\{n\d*\}/i.test(v)) return r;
      const id = expandDefault(col.defaultValue as string, unit, current, col.code, used);
      const parts = splitAutoTemplate(col.defaultValue as string, unit);
      const n = parts ? parseInt(id.slice(parts.prefix.length, id.length - parts.suffix.length), 10) : NaN;
      if (!isNaN(n)) used = Math.max(used, n);
      return { ...r, values: { ...r.values, [col.code]: id } };
    });
    current = { ...current, rows };
  });
  return current;
}

/** Giá trị mặc định của một ô theo cột (mã tự tăng {n<k>} do createRow/fillAutoIds điền). */
export function defaultCellValue(c: TableColumn): CellValue {
  if (c.type === 'calc') return 0;
  if (c.type === 'number') {
    const n = c.defaultValue !== undefined ? parseFloat(c.defaultValue) : 0;
    return isNaN(n) ? 0 : n;
  }
  if (c.type === 'select') return c.defaultValue !== undefined ? c.defaultValue : c.options && c.options.length ? c.options[0] : '';
  return c.defaultValue !== undefined ? c.defaultValue.replace(/\{unit\}/gi, '') : '';
}

export function createRow(table: ReportTable, unit: string, counters: Record<string, number> = {}): TableRow {
  const values: Record<string, CellValue> = {};
  table.columns.forEach((c) => {
    values[c.code] =
      c.type === 'text' && c.defaultValue !== undefined
        ? expandDefault(c.defaultValue, unit, table, c.code, counters[autoIdKey(table, c)] || 0)
        : defaultCellValue(c);
  });
  return recalcRow(table.columns, {
    id: nextRowId(table),
    isLocked: false,
    values
  });
}

/** Mã dòng mới: r<số lớn nhất đang có trong mục + 1>; không bao giờ trùng dòng đang có. */
export function nextRowId(table: Pick<ReportTable, 'rows'>): string {
  let max = 0;
  table.rows.forEach((r) => {
    const m = /^r(\d+)$/.exec(r.id);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return `r${max + 1}`;
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
  // (ô mã + mũi tên dropdown + lề: cần khoảng 210px để không bị cắt thành "MAT-ST01-000…")
  const floor = col.optionsRef && !col.optionsBy
    ? 210
    : col.optionsRef
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
