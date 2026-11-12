import { ReportMetadata, RequirableField } from '../types/report.types';

/** Template key (01_General, lowercase) -> metadata field + display label. */
export const REQUIRABLE_FIELDS: Record<string, { field: RequirableField; label: string }> = {
  location: { field: 'location', label: 'Vị trí / Trạm (Location)' },
  reportno: { field: 'reportNo', label: 'Số BCN (Report no.)' },
  subcontractor: { field: 'subcontractor', label: 'Đơn vị thi công (Subcontractor)' }
};

/** Used when the template has no `Required` column. */
export const DEFAULT_REQUIRED_FIELDS: RequirableField[] = ['location', 'reportNo', 'subcontractor'];

export function getRequiredFields(m: ReportMetadata): RequirableField[] {
  return m.requiredFields ?? DEFAULT_REQUIRED_FIELDS;
}

export function getFieldLabel(field: RequirableField): string {
  const entry = Object.keys(REQUIRABLE_FIELDS)
    .map((k) => REQUIRABLE_FIELDS[k])
    .filter((e) => e.field === field)[0];
  return entry ? entry.label : field;
}

/** Required fields that are currently blank. */
export function getMissingRequiredFields(m: ReportMetadata): RequirableField[] {
  return getRequiredFields(m).filter((f) => !String(m[f] ?? '').trim());
}

export { getMissingFormFields } from './reportForm';
