import { CellValue, FormField, ReportDoc, ReportForm } from '../types/report.types';
import { specWithOptions } from './reportTable';

/**
 * Mục dạng form: một sheet có dòng 3 bắt đầu bằng `Layout` và dòng 4 bắt đầu bằng `Key`.
 *
 *   Dòng 1: Title | <Tiêu đề VI> | <Tiêu đề EN> | <Nhãn ngắn>
 *   Dòng 2: Summary | <tóm tắt> | [off]           (off = không hiện ô Tóm tắt)
 *   Dòng 3: Layout | <số cột lưới> | [sign]       (sign = in dạng khung chữ ký)
 *   Dòng 4: Key | Label | Spec | Value | Required
 *   Dòng 5+: mỗi dòng một trường
 *
 * Spec: text · textarea · number · select:A/B/C (có mục cuối "Khác…" để nhập giá trị mới, ghi lại vào spec)
 * Cờ (nối bằng `;`): full (trải hết hàng) · span=<n> · rows=<n> · ph=<gợi ý> ·
 *     default=<giá trị; {user} {siteManager} {unit} được điền tự động khi trống>
 */

export function parseFieldSpec(key: string, label: string, rawSpec: string, value: CellValue, required: boolean): FormField {
  const spec = (rawSpec || '').trim();
  const tokens = spec.split(';').map((t) => t.trim()).filter(Boolean);
  const field: FormField = { key, label: label || key, spec, type: 'text', value, required };
  tokens.forEach((tok, i) => {
    const lower = tok.toLowerCase();
    if (i === 0 && lower.startsWith('select:')) {
      field.type = 'select';
      field.options = tok.slice(7).split('/').map((o) => o.trim()).filter(Boolean);
    } else if (i === 0 && (lower === 'textarea' || lower === 'number')) {
      field.type = lower as FormField['type'];
    } else if (lower === 'full') {
      field.full = true;
    } else if (lower.startsWith('span=')) {
      const n = parseInt(tok.slice(5), 10);
      if (!isNaN(n) && n > 0) field.span = n;
    } else if (lower.startsWith('rows=')) {
      const n = parseInt(tok.slice(5), 10);
      if (!isNaN(n) && n > 0) field.rows = n;
    } else if (lower.startsWith('ph=')) {
      field.placeholder = tok.slice(3);
    } else if (lower.startsWith('default=')) {
      field.defaultValue = tok.slice(8);
    }
  });
  return field;
}

export interface FormContext {
  unit: string;
  email: string;
  siteManager?: string;
}

function expandTokens(tpl: string, ctx: FormContext): string {
  return tpl
    .replace(/\{unit\}/gi, ctx.unit)
    .replace(/\{user\}/gi, ctx.email)
    .replace(/\{siteManager\}/gi, ctx.siteManager || '');
}

/** Điền giá trị mặc định (có {user}/{siteManager}/{unit}) cho các trường đang trống. */
export function applyFormDefaults(doc: ReportDoc, ctx: FormContext): ReportDoc {
  let changed = false;
  const sections = doc.sections.map((s) => {
    if (s.kind !== 'form') return s;
    let formChanged = false;
    const fields = s.fields.map((f) => {
      if (f.defaultValue === undefined) return f;
      const empty = f.value === undefined || f.value === '';
      if (!empty) return f;
      const v = expandTokens(f.defaultValue, ctx);
      if (v === '') return f;
      formChanged = true;
      return { ...f, value: f.type === 'number' ? parseFloat(v) || 0 : v };
    });
    if (!formChanged) return s;
    changed = true;
    return { ...s, fields };
  });
  return changed ? { ...doc, sections } : doc;
}

export interface MissingFormField {
  sectionKey: string;
  fieldKey: string;
  label: string;
}

export function isFieldEmpty(f: FormField): boolean {
  return f.value === undefined || String(f.value).trim() === '';
}

/** Trường bắt buộc (cột Required của sheet form) đang để trống. */
export function getMissingFormFields(doc: Pick<ReportDoc, 'sections'>): MissingFormField[] {
  const out: MissingFormField[] = [];
  doc.sections.forEach((s) => {
    if (s.kind !== 'form') return;
    s.fields.forEach((f) => {
      if (f.required && isFieldEmpty(f)) {
        out.push({ sectionKey: s.key, fieldKey: f.key, label: f.label.replace(/\s*\|\s*/, ' (') + (f.label.indexOf('|') !== -1 ? ')' : '') });
      }
    });
  });
  return out;
}

export function formFieldDisplay(f: FormField): string {
  if (f.value === undefined || f.value === '') return '';
  return typeof f.value === 'number' ? f.value.toLocaleString('vi-VN') : String(f.value);
}

/**
 * Số cột thực tế của mục form. Mục ký duyệt (`sign`) tối đa `Layout` cột và tự co theo số người ký:
 * 2 trường ngắn → 2 cột, 3 → 3 cột (không để cột trống).
 */
export function effectiveColumns(form: ReportForm): number {
  if (!form.sign) return form.columns;
  const shortCount = form.fields.filter((f) => !f.full).length;
  return Math.max(1, Math.min(form.columns, shortCount));
}

/** Chia trường thành các khối: nhóm trường ngắn (lưới cột) xen kẽ trường trải hết hàng. */
export type FormBlock = { kind: 'short'; fields: FormField[] } | { kind: 'full'; field: FormField };

export function buildFormBlocks(form: ReportForm): FormBlock[] {
  const blocks: FormBlock[] = [];
  let cur: FormField[] = [];
  let used = 0;
  const flush = (): void => {
    if (cur.length) blocks.push({ kind: 'short', fields: cur });
    cur = [];
    used = 0;
  };
  form.fields.forEach((f) => {
    if (f.full) {
      flush();
      blocks.push({ kind: 'full', field: f });
      return;
    }
    const cols = effectiveColumns(form);
    const span = Math.min(f.span || 1, cols);
    if (used + span > cols) flush();
    cur.push(f);
    used += span;
  });
  flush();
  return blocks;
}

/** Thêm một giá trị mới vào danh sách chọn của trường (cập nhật cả `options` lẫn `spec`). */
export function fieldWithOption(f: FormField, value: string): FormField {
  const options = f.options || [];
  if (options.some((o) => o.toLowerCase() === value.toLowerCase())) return f;
  const next = options.concat(value);
  return { ...f, options: next, spec: specWithOptions(f.spec, next) };
}
