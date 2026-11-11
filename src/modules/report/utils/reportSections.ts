import { ReportDoc, ReportSectionKey } from '../types/report.types';

/** Các bảng dữ liệu mà template có thể bật/tắt (mỗi bảng = một sheet 0N_<Tên>). */
export const TABLE_SECTIONS: ReportSectionKey[] = [
  'manpower',
  'equipment',
  'workqty',
  'materials',
  'issues'
];

/** Tên sheet (không kèm tiền tố số) ứng với từng bảng; so khớp không phân biệt hoa thường. */
export const SECTION_SHEET_NAME: Record<ReportSectionKey, string> = {
  manpower: 'Manpower',
  equipment: 'Equipment',
  workqty: 'WorkQty',
  materials: 'Materials',
  issues: 'Issues'
};

/** Bỏ tiền tố số của tên sheet: "05_Issues" / "5.Issues" -> "issues". */
export function sheetBaseName(sheetName: string): string {
  return sheetName.trim().replace(/^\d+\s*[_.\-\s]\s*/, '').toLowerCase();
}

/** Tìm tên sheet thực tế trong workbook cho một bảng (bất kể tiền tố số). */
export function findSectionSheet(sheetNames: string[], key: ReportSectionKey): string | undefined {
  const target = SECTION_SHEET_NAME[key].toLowerCase();
  return sheetNames.find((n) => sheetBaseName(n) === target);
}

/** Doc có bảng này không? `sections` không khai báo => đủ mọi bảng (tương thích dữ liệu cũ). */
export function hasSection(doc: Pick<ReportDoc, 'sections'>, key: ReportSectionKey): boolean {
  return !doc.sections || doc.sections.indexOf(key) !== -1;
}

const ROMAN = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV'];

export type ReportSectionId = ReportSectionKey | 'general' | 'worksplan' | 'hse' | 'signoff';

/** Thứ tự cố định các mục trên form/bản in. */
const ORDER: ReportSectionId[] = [
  'general',
  'manpower',
  'equipment',
  'workqty',
  'materials',
  'worksplan',
  'issues',
  'hse',
  'signoff'
];

export interface ReportSectionNumbering {
  /** Số La Mã của mục (rỗng nếu mục bị ẩn). */
  roman: (id: ReportSectionId) => string;
  /** Các mục đang hiển thị theo thứ tự, kèm id neo `report-sec-<n>`. */
  visible: { id: ReportSectionId; roman: string; anchor: string }[];
}

/** Đánh số La Mã liên tục cho các mục đang có (mục bị ẩn không chừa số). */
export function buildSectionNumbering(doc: Pick<ReportDoc, 'sections'>): ReportSectionNumbering {
  const map: Partial<Record<ReportSectionId, string>> = {};
  const visible: ReportSectionNumbering['visible'] = [];
  let n = 0;
  ORDER.forEach((id) => {
    if (TABLE_SECTIONS.indexOf(id as ReportSectionKey) !== -1 && !hasSection(doc, id as ReportSectionKey)) {
      return;
    }
    n += 1;
    const roman = ROMAN[n] || String(n);
    map[id] = roman;
    visible.push({ id, roman, anchor: `report-sec-${id}` });
  });
  return { roman: (id) => map[id] || '', visible };
}
