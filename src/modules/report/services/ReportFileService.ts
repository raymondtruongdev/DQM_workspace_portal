import { ISharePointFileInfo, SharePointClient } from '../../../core/services';
import { sanitizeSharePointPath } from '../../../core/utils';
import { REPORT_CONFIG } from '../constants/reportConfig';
import { TemplateVersionInfo } from '../types/report.types';
import {
  getReportFileName,
  getReportFolderRelativePath,
  getReportFileRelativePath,
  getReportPhotosFolderRelativePath,
  getTemplateFileRelativePath,
  getUnitFolderRelativePath,
  matchUnitTemplateFileName
} from '../utils/reportFileNaming';

function getPreviousCalendarDateStr(dateStr: string): string {
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const d = parseInt(parts[2], 10);
    const dt = new Date(y, m, d);
    if (!isNaN(dt.getTime())) {
      dt.setDate(dt.getDate() - 1);
      const yy = dt.getFullYear();
      const mm = ('0' + (dt.getMonth() + 1)).slice(-2);
      const dd = ('0' + dt.getDate()).slice(-2);
      return `${yy}-${mm}-${dd}`;
    }
  }
  return '';
}

export class ReportFileService {
  private _client: SharePointClient;
  private _siteUrl: string;
  private _libraryName: string;

  constructor(client: SharePointClient, siteUrl: string, libraryName = REPORT_CONFIG.LIBRARY_NAME) {
    this._client = client;
    this._siteUrl = siteUrl.replace(/\/+$/, '');
    this._libraryName = libraryName;
  }

  public getSiteRelativeUrl(): string {
    if (!this._siteUrl) return '';
    try {
      if (this._siteUrl.startsWith('http://') || this._siteUrl.startsWith('https://')) {
        const url = new URL(this._siteUrl);
        return url.pathname.replace(/\/+$/, '');
      }
    } catch {
      // ignore
    }
    return '/' + this._siteUrl.replace(/^\/+/, '').replace(/\/+$/, '');
  }

  public getLibraryPath(): string {
    const sitePath = this.getSiteRelativeUrl();
    return `${sitePath}/${this._libraryName}`;
  }

  /**
   * Path to Report_Role.xlsx
   */
  public getRoleFilePath(): string {
    return `${this.getLibraryPath()}/${REPORT_CONFIG.ROLE_FILE_NAME}`;
  }

  /**
   * Fetch Report_Role.xlsx buffer (searches library root and _template fallback)
   */
  public async getRoleFileBuffer(): Promise<ArrayBuffer | undefined> {
    const path = this.getRoleFilePath();
    try {
      const direct = await this._client.getFileBuffer(this._siteUrl, path);
      if (direct) return direct;
    } catch {
      // Fallback to folder listing below
    }

    const searchFolders = [this.getLibraryPath(), this.getTemplateFolderPath()];
    for (let i = 0; i < searchFolders.length; i++) {
      try {
        const files = await this._client.getFolderFiles(this._siteUrl, searchFolders[i]);
        const matched = files.find((f) => /^report[_-]?role.*\.xlsx$/i.test(f.name.trim()));
        if (matched) {
          const buf = await this._client.getFileBuffer(this._siteUrl, matched.serverRelativeUrl);
          if (buf) return buf;
        }
      } catch {
        // Ignore and continue
      }
    }

    return undefined;
  }

  /**
   * Path to _template folder
   */
  public getTemplateFolderPath(): string {
    return `${this.getLibraryPath()}/${REPORT_CONFIG.TEMPLATE_FOLDER}`;
  }

  /**
   * Path to template file for DocID (e.g. DailyReport_template.xlsx)
   */
  public getTemplateFilePath(docId: string): string {
    return getTemplateFileRelativePath(this.getLibraryPath(), docId);
  }

  /**
   * Find the latest template file for a specific Unit inside _template/
   * e.g. _template/template_ST01_20261109.xlsx
   * Falls back to _template/<DocID>_template.xlsx if no unit-specific template is found.
   */
  public async findLatestTemplateForUnit(
    unit: string,
    docId: string
  ): Promise<{ info?: TemplateVersionInfo; buffer?: ArrayBuffer }> {
    const primaryFolder = this.getTemplateFolderPath();
    const fallbackFolder = `${this.getLibraryPath()}/template`;

    let files: ISharePointFileInfo[] = [];
    try {
      files = await this._client.getFolderFiles(this._siteUrl, primaryFolder);
    } catch {
      files = [];
    }

    if (files.length === 0) {
      try {
        files = await this._client.getFolderFiles(this._siteUrl, fallbackFolder);
      } catch {
        files = [];
      }
    }

    const matchedCandidates: TemplateVersionInfo[] = [];
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      const match = matchUnitTemplateFileName(f.name, unit, docId);
      if (match) {
        matchedCandidates.push({
          fileName: match.fileName,
          version: match.version,
          versionDate: match.versionDate,
          serverRelativeUrl: f.serverRelativeUrl,
          timeLastModified: f.timeLastModified
        });
      }
    }

    if (matchedCandidates.length > 0) {
      matchedCandidates.sort((a, b) => {
        if (a.versionDate !== b.versionDate) {
          return b.versionDate.localeCompare(a.versionDate);
        }
        return (b.timeLastModified || '').localeCompare(a.timeLastModified || '');
      });

      const best = matchedCandidates[0];
      try {
        const buffer = await this._client.getFileBuffer(this._siteUrl, best.serverRelativeUrl);
        if (buffer) {
          return { info: best, buffer };
        }
      } catch {
        // Fallback below
      }
    }

    // Fallback to generic <DocID>_template.xlsx if no unit-specific template exists
    try {
      const genericBuffer = await this.getTemplateBuffer(docId);
      if (genericBuffer) {
        const genericFileName = `${sanitizeSharePointPath(docId)}_template.xlsx`;
        return {
          info: {
            fileName: genericFileName,
            version: `${sanitizeSharePointPath(docId)}_template`,
            versionDate: '',
            serverRelativeUrl: this.getTemplateFilePath(docId)
          },
          buffer: genericBuffer
        };
      }
    } catch {
      // Ignore
    }

    return {};
  }

  /**
   * Fetch template buffer for DocID
   */
  public async getTemplateBuffer(docId: string): Promise<ArrayBuffer | undefined> {
    const path = this.getTemplateFilePath(docId);
    let buf = await this._client.getFileBuffer(this._siteUrl, path);
    if (!buf) {
      // Fallback: without underscore e.g. template/DailyReport_template.xlsx
      const fallbackPath = `${this.getLibraryPath()}/template/${sanitizeSharePointPath(docId)}_template.xlsx`;
      buf = await this._client.getFileBuffer(this._siteUrl, fallbackPath);
    }
    return buf;
  }

  /**
   * Fetch existing report file with meta & ETag
   */
  public async getReportFileWithMeta(
    unit: string,
    docId: string,
    dateStr: string
  ): Promise<{ buffer: ArrayBuffer; eTag: string } | undefined> {
    const path = getReportFileRelativePath(this.getLibraryPath(), unit, docId, dateStr);
    return await this._client.getFileWithMeta(this._siteUrl, path);
  }

  /**
   * Find and load the closest previous report file before dateStr
   * (Checks immediate previous day first, then scans unit subfolders for the closest prior date)
   */
  public async getPreviousReportFile(
    unit: string,
    docId: string,
    dateStr: string
  ): Promise<{ buffer: ArrayBuffer; sourceDateStr: string } | undefined> {
    // 1. Try the immediate adjacent previous calendar day (dateStr - 1 day)
    const prevDateStr = getPreviousCalendarDateStr(dateStr);
    if (prevDateStr) {
      try {
        const adjFile = await this.getReportFileWithMeta(unit, docId, prevDateStr);
        if (adjFile && adjFile.buffer) {
          return { buffer: adjFile.buffer, sourceDateStr: prevDateStr };
        }
      } catch {
        // Ignore and fallback to folder scan
      }
    }

    // 2. Scan subfolders in data/<Unit> for the most recent <DocID>_<yyyyMMdd> prior to dateStr
    try {
      const unitFolder = this.getUnitFolderPath(unit);
      const subfolders = await this._client.getFolderSubfolderNames(this._siteUrl, unitFolder);
      if (!subfolders || subfolders.length === 0) return undefined;

      const cleanDocId = sanitizeSharePointPath(docId).toLowerCase();
      const targetCompact = dateStr.replace(/-/g, '');
      const prefix = `${cleanDocId}_`;

      const priorDates: string[] = [];
      for (let i = 0; i < subfolders.length; i++) {
        const name = subfolders[i];
        const lower = name.toLowerCase();
        if (lower.indexOf(prefix) === 0) {
          const compactDate = name.slice(prefix.length).trim();
          if (/^\d{8}$/.test(compactDate) && compactDate < targetCompact) {
            const isoDate = `${compactDate.slice(0, 4)}-${compactDate.slice(4, 6)}-${compactDate.slice(6, 8)}`;
            if (isoDate !== prevDateStr) {
              priorDates.push(isoDate);
            }
          }
        }
      }

      // Sort descending so closest date comes first
      priorDates.sort((a, b) => b.localeCompare(a));

      for (let i = 0; i < Math.min(priorDates.length, 3); i++) {
        const candidateDate = priorDates[i];
        const found = await this.getReportFileWithMeta(unit, docId, candidateDate);
        if (found && found.buffer) {
          return { buffer: found.buffer, sourceDateStr: candidateDate };
        }
      }
    } catch {
      // Ignore folder scan errors
    }

    return undefined;
  }

  /**
   * Path to the Unit folder:
   * <libraryPath>/data/<Unit>
   */
  public getUnitFolderPath(unit: string): string {
    return getUnitFolderRelativePath(this.getLibraryPath(), unit);
  }

  /**
   * Explicitly check and ensure the Unit folder (e.g. data/ST01) exists on SharePoint
   */
  public async ensureUnitFolder(unit: string): Promise<string> {
    const unitFolderPath = this.getUnitFolderPath(unit);
    await this._client.ensureFolder(this._siteUrl, unitFolderPath);
    return unitFolderPath;
  }

  /**
   * Save report file to SharePoint: ensures folder data/<Unit> and data/<Unit>/<DocID>_<yyyyMMdd> exist, then uploads file
   */
  public async saveReportFile(
    unit: string,
    docId: string,
    dateStr: string,
    buffer: ArrayBuffer,
    eTag?: string
  ): Promise<{ serverRelativeUrl: string; eTag?: string }> {
    // 1. Explicitly check and ensure the Unit folder (e.g. data/ST01) exists
    await this.ensureUnitFolder(unit);

    // 2. Ensure the specific report folder exists (e.g. data/ST01/DailyReport_20261110)
    const folderPath = getReportFolderRelativePath(this.getLibraryPath(), unit, docId, dateStr);
    await this._client.ensureFolder(this._siteUrl, folderPath);

    // 3. Upload file
    const fileName = getReportFileName(docId, dateStr);
    return await this._client.uploadFile(this._siteUrl, folderPath, fileName, buffer, true, eTag);
  }

  /**
   * Path to photos subfolder:
   * <libraryPath>/data/<Unit>/<DocID>_<yyyyMMdd>/photos
   */
  public getReportPhotosFolderPath(unit: string, docId: string, dateStr: string): string {
    return getReportPhotosFolderRelativePath(this.getLibraryPath(), unit, docId, dateStr);
  }

  /**
   * Upload a photo into data/<Unit>/<DocID>_<yyyyMMdd>/photos/<groupId>/
   * Returns the stored file name.
   */
  public async uploadReportPhoto(
    unit: string,
    docId: string,
    dateStr: string,
    groupId: string,
    fileName: string,
    buffer: ArrayBuffer
  ): Promise<string> {
    await this.ensureUnitFolder(unit);
    const reportFolder = getReportFolderRelativePath(this.getLibraryPath(), unit, docId, dateStr);
    await this._client.ensureFolder(this._siteUrl, reportFolder);

    const groupFolder = `${this.getReportPhotosFolderPath(unit, docId, dateStr)}/${groupId}`;
    await this._client.ensureFolder(this._siteUrl, groupFolder);

    const safeName = sanitizeSharePointPath(fileName) || `photo_${Date.now()}.jpg`;
    await this._client.uploadFile(this._siteUrl, groupFolder, safeName, buffer, true);
    return safeName;
  }

  /**
   * Delete an uploaded photo
   */
  public async deleteReportPhoto(serverRelativeUrl: string): Promise<void> {
    await this._client.deleteFile(this._siteUrl, serverRelativeUrl);
  }
}

