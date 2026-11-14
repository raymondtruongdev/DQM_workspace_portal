import {
  sanitizeSharePointPath,
  formatDateToYYYYMMDDWithoutDash,
  parseDateString
} from '../../../core/utils';
import { TemplateVersionInfo } from '../types/report.types';

/**
 * Return formatted yyyyMMdd without dash from "YYYY-MM-DD" or Date
 */
export function toYYYYMMDD(date: string | Date): string {
  if (typeof date === 'string') {
    const clean = date.replace(/-/g, '').trim();
    if (clean.length === 8 && /^\d{8}$/.test(clean)) {
      return clean;
    }
    return formatDateToYYYYMMDDWithoutDash(parseDateString(date));
  }
  return formatDateToYYYYMMDDWithoutDash(date);
}

/**
 * Tên chung của một báo cáo: <DRP|WRP>_<yyyyMMdd>_<Unit>, e.g. DRP_20261114_ST01.
 * Dùng làm tên thư mục và tên file (.json, .html) của báo cáo đó.
 */
export function getReportFolderName(docId: string, unit: string, date: string | Date): string {
  return `${getDocShortPrefix(docId)}_${toYYYYMMDD(date)}_${sanitizeSharePointPath(unit)}`;
}

/**
 * File dữ liệu báo cáo trong thư mục: <DRP|WRP>_<yyyyMMdd>_<Unit>.json, e.g. DRP_20261114_ST01.json
 */
export function getReportFileName(docId: string, unit: string, date: string | Date): string {
  return `${getReportFolderName(docId, unit, date)}.json`;
}

/**
 * Bản xem/in độc lập (nhúng sẵn dữ liệu) cùng thư mục: <DRP|WRP>_<yyyyMMdd>_<Unit>.html
 */
export function getReportHtmlFileName(docId: string, unit: string, date: string | Date): string {
  return `${getReportFolderName(docId, unit, date)}.html`;
}

/**
 * Name of the signed PDF stored next to the report: <DocCode>_signed.pdf (cùng tên file PDF in ra + hậu tố _signed),
 * e.g. HCM2_1-23_CS-THAC_DCEM-CPZZ-ST01-CWS-DRP-00001_signed.pdf.
 * Không có Doc. code thì dùng tên theo mã báo cáo: <DRP|WRP>_<yyyyMMdd>_<Unit>_signed.pdf.
 */
export function getSignedPdfFileName(docId: string, unit: string, date: string | Date, docCode = ''): string {
  const code = docCode.replace(/[\\/:*?"<>|\s]+/g, '_').replace(/^_+|_+$/g, '');
  if (code) return `${code}_signed.pdf`;
  return `${getDocShortPrefix(docId)}_${toYYYYMMDD(date)}_${sanitizeSharePointPath(unit)}_signed.pdf`;
}

/** File PDF đã ký (mọi kiểu đặt tên đang/đã dùng) có hậu tố _signed.pdf. */
export function isSignedPdfName(name: string): boolean {
  return /_signed\.pdf$/i.test(name);
}

/**
 * Server relative path for unit folder:
 * <libraryPath>/data/<Unit>
 */
export function getUnitFolderRelativePath(
  libraryPath: string,
  unit: string
): string {
  const cleanLib = libraryPath.replace(/\/+$/, '');
  const cleanUnit = sanitizeSharePointPath(unit);
  return `${cleanLib}/data/${cleanUnit}`;
}

/**
 * Server relative path for the report folder:
 * <libraryPath>/data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>
 */
export function getReportFolderRelativePath(
  libraryPath: string,
  unit: string,
  docId: string,
  date: string | Date
): string {
  const cleanLib = libraryPath.replace(/\/+$/, '');
  const cleanUnit = sanitizeSharePointPath(unit);
  const folderName = getReportFolderName(docId, unit, date);
  return `${cleanLib}/data/${cleanUnit}/${folderName}`;
}

/**
 * Server relative path for the report file:
 * <libraryPath>/data/<Unit>/DRP_<yyyyMMdd>_<Unit>/DRP_<yyyyMMdd>_<Unit>.json
 */
export function getReportFileRelativePath(
  libraryPath: string,
  unit: string,
  docId: string,
  date: string | Date
): string {
  const folderPath = getReportFolderRelativePath(libraryPath, unit, docId, date);
  const fileName = getReportFileName(docId, unit, date);
  return `${folderPath}/${fileName}`;
}

/**
 * Server relative path for the report photos subfolder:
 * <libraryPath>/data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/photos
 */
export function getReportPhotosFolderRelativePath(
  libraryPath: string,
  unit: string,
  docId: string,
  date: string | Date
): string {
  const folderPath = getReportFolderRelativePath(libraryPath, unit, docId, date);
  return `${folderPath}/photos`;
}

/**
 * Extract 8-digit yyyyMMdd date from a template version string or filename
 * e.g. "template_ST01_20261109" -> "20261109"
 */
export function extractTemplateVersionDate(versionStr?: string): string {
  if (!versionStr) return '';
  const clean = versionStr.replace(/\.json$/i, '').trim();
  const matchCompact = clean.match(/(\d{8})$/) || clean.match(/_(\d{8})(?:_|$)/);
  if (matchCompact) {
    return matchCompact[1];
  }
  const matchIso = clean.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (matchIso) {
    return `${matchIso[1]}${matchIso[2]}${matchIso[3]}`;
  }
  return '';
}

/**
 * Format template version string for friendly UI & PDF display
 * e.g. "template_ST01_20261109" -> "template_ST01_20261109 (09/11/2026)"
 */
export function formatTemplateVersionDisplay(versionStr?: string): string {
  if (!versionStr || !versionStr.trim()) {
    return 'Chưa gắn phiên bản (Mẫu mặc định)';
  }
  const clean = versionStr.replace(/\.json$/i, '').trim();
  const d8 = extractTemplateVersionDate(clean);
  if (d8 && d8.length === 8) {
    const formattedDate = `${d8.slice(6, 8)}/${d8.slice(4, 6)}/${d8.slice(0, 4)}`;
    return `${clean} (${formattedDate})`;
  }
  return clean;
}

/**
 * Map DocID to short template prefix:
 * DailyReport -> DRP
 * WeeklyReport -> WRP
 */
export function getDocShortPrefix(docId?: string): string {
  const lower = (docId || 'DailyReport').trim().toLowerCase();
  return lower.includes('week') || lower === 'wrp' ? 'WRP' : 'DRP';
}

/**
 * Match a file name inside _template/ against a target unit and DocID.
 * Pattern: template_<DRP|WRP>_<Unit>_<yyyyMMdd>.json (e.g. template_DRP_ST01_20261114.json)
 */
export function matchUnitTemplateFileName(
  fileName: string,
  targetUnit: string,
  docId?: string
): { version: string; versionDate: string; fileName: string } | undefined {
  const trimmed = fileName.trim();
  if (!/\.json$/i.test(trimmed)) return undefined;

  const base = trimmed.replace(/\.json$/i, '');
  const match = base.match(/^template_([a-z]+)_([a-z0-9-]+)_(\d{8})$/i);
  if (!match) return undefined;
  if (match[1].toUpperCase() !== getDocShortPrefix(docId)) return undefined;
  if (match[2].toLowerCase() !== sanitizeSharePointPath(targetUnit).toLowerCase()) return undefined;
  return { fileName: trimmed, version: base, versionDate: match[3] };
}

/**
 * Check if a report's templateVersion is outdated compared to the latest TemplateVersionInfo in _template/
 */
export function isTemplateVersionOutdated(
  docTemplateVersion: string | undefined,
  latestTemplate: TemplateVersionInfo | undefined
): boolean {
  if (!latestTemplate || !latestTemplate.version) {
    return false;
  }

  // Only enforce template upgrade prompt when a unit-specific versioned template exists (or has versionDate)
  const latestVer = latestTemplate.version.replace(/\.json$/i, '').trim();
  const latestDate = latestTemplate.versionDate || extractTemplateVersionDate(latestVer);

  const docVer = (docTemplateVersion || '').replace(/\.json$/i, '').trim();
  if (!docVer) {
    // Report has no templateVersion recorded (legacy/old template)
    return Boolean(latestDate || /^template_/i.test(latestVer));
  }

  if (docVer.toLowerCase() === latestVer.toLowerCase()) {
    return false;
  }

  const docDate = extractTemplateVersionDate(docVer);
  if (latestDate && docDate) {
    return docDate < latestDate;
  }

  return docVer.toLowerCase() !== latestVer.toLowerCase();
}

