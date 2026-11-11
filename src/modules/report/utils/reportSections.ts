import { ReportDoc } from '../types/report.types';

/** Mã mục = khoá của sheet (chữ thường): "weather", "manpower", "issues", … */
export type ReportSectionId = string;

const ROMAN = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX'];

export interface ReportSectionNumbering {
  /** Số La Mã của mục (rỗng nếu mục không có). */
  roman: (id: ReportSectionId) => string;
  /** Các mục theo thứ tự, kèm id neo và nhãn thanh nhảy nhanh. */
  visible: { id: ReportSectionId; roman: string; anchor: string; label: string }[];
}

function capitalize(s: string): string {
  const t = s.trim().toLowerCase();
  return t.charAt(0).toUpperCase() + t.slice(1);
}

export function sectionAnchor(key: string): string {
  return `report-sec-${key}`;
}

/** Thứ tự mục = thứ tự sheet trong template; số La Mã liên tục. */
export function buildSectionNumbering(doc: Pick<ReportDoc, 'sections'>): ReportSectionNumbering {
  const map: Record<string, string> = {};
  const visible = doc.sections.map((sec, i) => {
    const roman = ROMAN[i + 1] || String(i + 1);
    map[sec.key] = roman;
    return {
      id: sec.key,
      roman,
      anchor: sectionAnchor(sec.key),
      label: sec.shortTitle || capitalize(sec.title)
    };
  });
  return { roman: (id) => map[id] || '', visible };
}
