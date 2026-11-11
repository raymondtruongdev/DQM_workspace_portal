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
 *   Dòng 1: Title | <Tiêu đề VI> | <Tiêu đề EN> | <Nhãn ngắn>
 *   Dòng 2: Summary | <Tóm tắt>
 *   Dòng 3: mã cột (Org, Nos, …, IsLocked)
 *   Dòng 4: nhãn "Tiếng Việt | English"
 *   Dòng 5: spec cột: <loại>[;cờ…]
 *   Dòng 6+: dữ liệu
 *
 * Loại: text · number · select:A/B/C · calc:<công thức, ví dụ Nos*Hrs>
 * Cờ: key (chỉ đọc khi dòng bị khoá) · total (có dòng Tổng) · left|center|right · w=<%> ·
 *     default=<giá trị, có thể dùng {unit} {n3}> · tone=<giá trị>:danger|warning,…
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
    } else if (i === 0 && lower.startsWith('calc:')) {
      col.type = 'calc';
      col.formula = tok.slice(5).trim();
    } else if (i === 0 && lower === 'number') {
      col.type = 'number';
    } else if (lower === 'key') {
      col.key = true;
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
        if (tone === 'danger' || tone === 'warning') tones[val] = tone;
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
  if (col.type === 'select') return 'center';
  return 'left';
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

function expandDefault(tpl: string, unit: string, rowCount: number): string {
  return tpl
    .replace(/\{unit\}/gi, unit)
    .replace(/\{n3\}/gi, ('000' + (rowCount + 1)).slice(-3))
    .replace(/\{n\}/gi, String(rowCount + 1));
}

export function createRow(table: ReportTable, unit: string): TableRow {
  const values: Record<string, CellValue> = {};
  table.columns.forEach((c) => {
    if (c.type === 'calc') {
      values[c.code] = 0;
    } else if (c.type === 'number') {
      const n = c.defaultValue !== undefined ? parseFloat(c.defaultValue) : 0;
      values[c.code] = isNaN(n) ? 0 : n;
    } else if (c.type === 'select') {
      values[c.code] =
        c.defaultValue !== undefined ? c.defaultValue : c.options && c.options.length ? c.options[0] : '';
    } else {
      values[c.code] = c.defaultValue !== undefined ? expandDefault(c.defaultValue, unit, table.rows.length) : '';
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

export function toneOf(col: TableColumn, value: CellValue): ColumnTone | undefined {
  return col.tones && typeof value === 'string' ? col.tones[value] : undefined;
}
