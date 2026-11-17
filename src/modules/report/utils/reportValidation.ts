import { ReportDoc, ReportMetadata, RequirableField } from '../types/report.types';
import { splitLabel } from './reportTable';

/** Trường thông tin chung bắt buộc nhập trước khi lưu/in, kèm nhãn hiển thị. */
export const REQUIRABLE_FIELDS: Record<RequirableField, string> = {
  location: 'Vị trí / Trạm (Location)',
  reportNo: 'Số BCN (Report no.)',
  subcontractor: 'Đơn vị thi công (Subcontractor)'
};

export const REQUIRED_FIELDS: RequirableField[] = ['location', 'reportNo', 'subcontractor'];

export function getFieldLabel(field: RequirableField): string {
  return REQUIRABLE_FIELDS[field];
}

function isBlank(m: ReportMetadata, f: RequirableField): boolean {
  return f === 'subcontractor' ? m.subcontractors.length === 0 : !String(m[f] ?? '').trim();
}

/** Required fields that are currently blank. */
export function getMissingRequiredFields(m: ReportMetadata): RequirableField[] {
  return REQUIRED_FIELDS.filter((f) => isBlank(m, f));
}

export interface MissingSelectCell {
  sectionKey: string;
  rowId: string;
  colCode: string;
  /** Mô tả để báo người dùng: "<mục> – <cột> (dòng n)". */
  label: string;
}

/** Ô chọn (dropdown) của các mục dạng bảng còn để trống ("Chọn…"). `---` (vd Activity ID = không có Activity) là đã chọn. */
export function getMissingSelectCells(doc: Pick<ReportDoc, 'sections'>): MissingSelectCell[] {
  const out: MissingSelectCell[] = [];
  doc.sections.forEach((sec) => {
    if (sec.kind !== 'table') return;
    const selects = sec.columns.filter((c) => c.type === 'select');
    sec.rows.forEach((row, i) => {
      selects.forEach((c) => {
        const v = row.values[c.code];
        if (v === undefined || String(v).trim() === '') {
          out.push({
            sectionKey: sec.key,
            rowId: row.id,
            colCode: c.code,
            label: `${sec.shortTitle || sec.title} – ${splitLabel(c.label).vi} (dòng ${i + 1})`
          });
        }
      });
    });
  });
  return out;
}

export { getMissingFormFields } from './reportForm';
