import { ReportMetadata, RequirableField } from '../types/report.types';

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

export { getMissingFormFields } from './reportForm';
