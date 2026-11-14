import { FormField, ReportDoc, ReportForm, ReportSection } from '../types/report.types';

/**
 * Trường văn bản của mục (`fields`): ghi chú HSE, khung ký duyệt, chỉ đạo… Định nghĩa (kiểu, full, span, rows, lines, mặc định)
 * đến từ template, giá trị đến từ file báo cáo (xem reportJson.ts).
 * Mặc định có thể dùng {user} {siteManager} {unit}, được điền tự động khi trường đang trống.
 */

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

function applyFieldDefaults(fields: FormField[], ctx: FormContext): FormField[] | undefined {
  let changed = false;
  const next = fields.map((f) => {
    if (f.defaultValue === undefined) return f;
    const empty = f.value === undefined || f.value === '';
    if (!empty) return f;
    const v = expandTokens(f.defaultValue, ctx);
    if (v === '') return f;
    changed = true;
    return { ...f, value: f.type === 'number' ? parseFloat(v) || 0 : v };
  });
  return changed ? next : undefined;
}

/** Điền giá trị mặc định (có {user}/{siteManager}/{unit}) cho các trường đang trống. */
export function applyFormDefaults(doc: ReportDoc, ctx: FormContext): ReportDoc {
  let changed = false;
  const sections = doc.sections.map((s): ReportSection => {
    if (s.kind === 'photo' || !s.fields) return s;
    const fields = applyFieldDefaults(s.fields, ctx);
    if (!fields) return s;
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

/** Trường văn bản bắt buộc đang để trống. */
export function getMissingFormFields(doc: Pick<ReportDoc, 'sections'>): MissingFormField[] {
  const out: MissingFormField[] = [];
  doc.sections.forEach((s) => {
    if (s.kind === 'photo' || !s.fields) return;
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

/** Thêm một giá trị mới vào danh sách chọn của trường. */
export function fieldWithOption(f: FormField, value: string): FormField {
  const options = f.options || [];
  if (options.some((o) => o.toLowerCase() === value.toLowerCase())) return f;
  return { ...f, options: options.concat(value) };
}
