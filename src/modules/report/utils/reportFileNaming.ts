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
 * Folder name: <DocID>_<yyyyMMdd>, e.g. DailyReport_20261110 or WeeklyReport_20261110
 */
export function getReportFolderName(docId: string, date: string | Date): string {
  const cleanDoc = sanitizeSharePointPath(docId);
  const dateStr = toYYYYMMDD(date);
  return `${cleanDoc}_${dateStr}`;
}

/**
 * File name inside folder: <DocID>_<yyyyMMdd>.xlsx, e.g. DailyReport_20261110.xlsx
 */
export function getReportFileName(docId: string, date: string | Date): string {
  return `${getReportFolderName(docId, date)}.xlsx`;
}

/**
 * Name of the signed PDF stored next to the report: <DocID>_<yyyyMMdd>_signed.pdf
 */
export function getSignedPdfFileName(docId: string, date: string | Date): string {
  return `${getReportFolderName(docId, date)}_signed.pdf`;
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
 * <libraryPath>/data/<Unit>/<DocID>_<yyyyMMdd>
 */
export function getReportFolderRelativePath(
  libraryPath: string,
  unit: string,
  docId: string,
  date: string | Date
): string {
  const cleanLib = libraryPath.replace(/\/+$/, '');
  const cleanUnit = sanitizeSharePointPath(unit);
  const folderName = getReportFolderName(docId, date);
  return `${cleanLib}/data/${cleanUnit}/${folderName}`;
}

/**
 * Server relative path for the report file:
 * <libraryPath>/data/<Unit>/<DocID>_<yyyyMMdd>/<DocID>_<yyyyMMdd>.xlsx
 */
export function getReportFileRelativePath(
  libraryPath: string,
  unit: string,
  docId: string,
  date: string | Date
): string {
  const folderPath = getReportFolderRelativePath(libraryPath, unit, docId, date);
  const fileName = getReportFileName(docId, date);
  return `${folderPath}/${fileName}`;
}

/**
 * Server relative path for the report photos subfolder:
 * <libraryPath>/data/<Unit>/<DocID>_<yyyyMMdd>/photos
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
 * Server relative path for template file:
 * <libraryPath>/_template/<DocID>_template.xlsx
 */
export function getTemplateFileRelativePath(libraryPath: string, docId: string): string {
  const cleanLib = libraryPath.replace(/\/+$/, '');
  const cleanDoc = sanitizeSharePointPath(docId);
  return `${cleanLib}/_template/${cleanDoc}_template.xlsx`;
}

/**
 * Extract 8-digit yyyyMMdd date from a template version string or filename
 * e.g. "template_ST01_20261109" -> "20261109"
 */
export function extractTemplateVersionDate(versionStr?: string): string {
  if (!versionStr) return '';
  const clean = versionStr.replace(/\.xlsx$/i, '').trim();
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
  const clean = versionStr.replace(/\.xlsx$/i, '').trim();
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
function getDocShortPrefix(docId?: string): string {
  const lower = (docId || 'DailyReport').trim().toLowerCase();
  return lower.includes('week') || lower === 'wrp' ? 'WRP' : 'DRP';
}

/**
 * Match a file name inside _template/ against a target unit and DocID.
 * Pattern: template_<DRP|WRP>_<Unit>_<yyyyMMdd>.xlsx (e.g. template_DRP_ST01_20261109.xlsx)
 */
export function matchUnitTemplateFileName(
  fileName: string,
  targetUnit: string,
  docId?: string
): { version: string; versionDate: string; fileName: string } | undefined {
  const trimmed = fileName.trim();
  if (!/\.xlsx$/i.test(trimmed)) return undefined;

  const base = trimmed.replace(/\.xlsx$/i, '');
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
  const latestVer = latestTemplate.version.replace(/\.xlsx$/i, '').trim();
  const latestDate = latestTemplate.versionDate || extractTemplateVersionDate(latestVer);

  const docVer = (docTemplateVersion || '').replace(/\.xlsx$/i, '').trim();
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

