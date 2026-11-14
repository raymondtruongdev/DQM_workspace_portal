import { ISharePointClient, ISharePointFileInfo } from '../../../core/services';
import { sanitizeSharePointPath } from '../../../core/utils';
import { REPORT_CONFIG } from '../constants/reportConfig';
import { TemplateVersionInfo, ReportDayInfo, SignedDocInfo } from '../types/report.types';
import {
  getReportFileName,
  getReportHtmlFileName,
  getReportFolderRelativePath,
  getReportFileRelativePath,
  getReportPhotosFolderRelativePath,
  getUnitFolderRelativePath,
  matchUnitTemplateFileName,
  getReportFolderName,
  getSignedPdfFileName,
  isSignedPdfName
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
  private _client: ISharePointClient;
  private _siteUrl: string;
  private _libraryName: string;

  constructor(client: ISharePointClient, siteUrl: string, libraryName = REPORT_CONFIG.LIBRARY_NAME) {
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
   * Fetch ActivityP6.xlsx (danh mục Activity P6 + cấu kiện) from the library root; undefined if missing
   */
  public async getActivityFileBuffer(): Promise<ArrayBuffer | undefined> {
    return await this._client.getFileBuffer(
      this._siteUrl,
      `${this.getLibraryPath()}/${REPORT_CONFIG.ACTIVITY_FILE_NAME}`
    );
  }

  /**
   * Path to _template folder
   */
  public getTemplateFolderPath(): string {
    return `${this.getLibraryPath()}/${REPORT_CONFIG.TEMPLATE_FOLDER}`;
  }


  /**
   * Template mới nhất của đơn vị trong _template/ (template_<DRP|WRP>_<Unit>_<yyyyMMdd>.json): ngày trong tên file lớn nhất thắng.
   * Không có thì trả {}.
   */
  public async findLatestTemplateForUnit(
    unit: string,
    docId: string
  ): Promise<{ info?: TemplateVersionInfo; text?: string }> {
    let files: ISharePointFileInfo[] = [];
    try {
      files = await this._client.getFolderFiles(this._siteUrl, this.getTemplateFolderPath());
    } catch {
      files = [];
    }
    const candidates: TemplateVersionInfo[] = [];
    files.forEach((f) => {
      const match = matchUnitTemplateFileName(f.name, unit, docId);
      if (match) {
        candidates.push({
          fileName: match.fileName,
          version: match.version,
          versionDate: match.versionDate,
          serverRelativeUrl: f.serverRelativeUrl,
          timeLastModified: f.timeLastModified
        });
      }
    });
    candidates.sort((a, b) =>
      a.versionDate !== b.versionDate
        ? b.versionDate.localeCompare(a.versionDate)
        : (b.timeLastModified || '').localeCompare(a.timeLastModified || '')
    );
    const best = candidates[0];
    if (!best) return {};
    const buffer = await this._client.getFileBuffer(this._siteUrl, best.serverRelativeUrl);
    return buffer ? { info: best, text: new TextDecoder('utf-8').decode(buffer) } : {};
  }

  /**
   * File dữ liệu của một ngày (.data.json) kèm ETag; undefined nếu chưa có.
   */
  public async getReportJsonWithMeta(
    unit: string,
    docId: string,
    dateStr: string
  ): Promise<{ text: string; eTag: string } | undefined> {
    const path = getReportFileRelativePath(this.getLibraryPath(), unit, docId, dateStr);
    const file = await this._client.getFileWithMeta(this._siteUrl, path);
    return file && file.buffer ? { text: new TextDecoder('utf-8').decode(file.buffer), eTag: file.eTag } : undefined;
  }

  /**
   * Báo cáo gần nhất trước `dateStr` (ngày liền trước, rồi quét thư mục data/<Unit> lấy ngày trước gần nhất).
   */
  public async getPreviousReportJson(
    unit: string,
    docId: string,
    dateStr: string
  ): Promise<{ text: string; sourceDateStr: string } | undefined> {
    // 1. Ngày liền trước (dateStr - 1)
    const prevDateStr = getPreviousCalendarDateStr(dateStr);
    if (prevDateStr) {
      try {
        const adj = await this.getReportJsonWithMeta(unit, docId, prevDateStr);
        if (adj) return { text: adj.text, sourceDateStr: prevDateStr };
      } catch {
        // thử quét thư mục
      }
    }

    // 2. Quét các thư mục con của data/<Unit> tìm <DocID>_<yyyyMMdd> gần nhất trước dateStr
    try {
      const subfolders = await this._client.getFolderSubfolderNames(this._siteUrl, this.getUnitFolderPath(unit));
      if (!subfolders || subfolders.length === 0) return undefined;
      const prefix = `${sanitizeSharePointPath(docId).toLowerCase()}_`;
      const targetCompact = dateStr.replace(/-/g, '');
      const priorDates: string[] = [];
      subfolders.forEach((name) => {
        if (name.toLowerCase().indexOf(prefix) !== 0) return;
        const compact = name.slice(prefix.length).trim();
        if (/^\d{8}$/.test(compact) && compact < targetCompact) {
          const iso = `${compact.slice(0, 4)}-${compact.slice(4, 6)}-${compact.slice(6, 8)}`;
          if (iso !== prevDateStr) priorDates.push(iso);
        }
      });
      priorDates.sort((a, b) => b.localeCompare(a));
      for (let i = 0; i < Math.min(priorDates.length, 3); i++) {
        const found = await this.getReportJsonWithMeta(unit, docId, priorDates[i]);
        if (found) return { text: found.text, sourceDateStr: priorDates[i] };
      }
    } catch {
      // bỏ qua lỗi quét thư mục
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

  /** Đảm bảo data/<Unit> và data/<Unit>/<DocID>_<yyyyMMdd> tồn tại; trả đường dẫn thư mục báo cáo. */
  private async ensureReportFolder(unit: string, docId: string, dateStr: string): Promise<string> {
    await this.ensureUnitFolder(unit);
    const folderPath = getReportFolderRelativePath(this.getLibraryPath(), unit, docId, dateStr);
    await this._client.ensureFolder(this._siteUrl, folderPath);
    return folderPath;
  }

  private static encode(text: string): ArrayBuffer {
    const bytes = new TextEncoder().encode(text);
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  }

  /**
   * Ghi file dữ liệu báo cáo (.data.json) vào data/<Unit>/<DocID>_<yyyyMMdd>/ (tự tạo thư mục), có kiểm tra ETag khi ghi đè.
   */
  public async saveReportJson(
    unit: string,
    docId: string,
    dateStr: string,
    text: string,
    eTag?: string
  ): Promise<{ serverRelativeUrl: string; eTag?: string }> {
    const folderPath = await this.ensureReportFolder(unit, docId, dateStr);
    return await this._client.uploadFile(this._siteUrl, folderPath, getReportFileName(docId, dateStr), ReportFileService.encode(text), true, eTag);
  }

  /** Ghi bản xem/in độc lập (.html) cạnh file dữ liệu; luôn ghi đè. */
  public async saveReportHtml(unit: string, docId: string, dateStr: string, html: string): Promise<void> {
    const folderPath = await this.ensureReportFolder(unit, docId, dateStr);
    await this._client.uploadFile(this._siteUrl, folderPath, getReportHtmlFileName(docId, dateStr), ReportFileService.encode(html), true);
  }

  /** Nội dung một ảnh hiện trường (để nhúng vào bản HTML độc lập). */
  public async getReportPhotoBuffer(unit: string, docId: string, dateStr: string, groupId: string, name: string): Promise<ArrayBuffer | undefined> {
    return await this._client.getFileBuffer(this._siteUrl, `${this.getReportPhotosFolderPath(unit, docId, dateStr)}/${groupId}/${name}`);
  }

  /** Nội dung một file trong thư mục resource (logo bản in). */
  public async getResourceFileBuffer(name: string): Promise<ArrayBuffer | undefined> {
    return await this._client.getFileBuffer(this._siteUrl, `${this.getLibraryPath()}/resource/${name}`);
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

  /** File PDF trong thư mục báo cáo (ưu tiên file có hậu tố _signed.pdf, nếu không thì PDF mới nhất). */
  private pickSignedPdf(files: ISharePointFileInfo[]): SignedDocInfo | undefined {
    const pdfs = files.filter((f) => /\.pdf$/i.test(f.name));
    if (pdfs.length === 0) return undefined;
    const byTime = (a: ISharePointFileInfo, b: ISharePointFileInfo): number =>
      (b.timeLastModified || '').localeCompare(a.timeLastModified || '');
    const best =
      pdfs.filter((f) => isSignedPdfName(f.name)).sort(byTime)[0] ||
      pdfs.sort(byTime)[0];
    return {
      name: best.name,
      serverRelativeUrl: best.serverRelativeUrl,
      length: best.length,
      timeLastModified: best.timeLastModified
    };
  }

  /**
   * Trạng thái báo cáo của các ngày (dateStrs: YYYY-MM-DD) trong data/<Unit>:
   * - missing: chưa có file báo cáo (.data.json) của ngày đó
   * - unsigned: có file báo cáo nhưng thư mục chưa có file PDF
   * - done: có file báo cáo và có file PDF đã ký
   */
  public async getReportDayInfos(
    unit: string,
    docId: string,
    dateStrs: string[],
    concurrency = 14
  ): Promise<ReportDayInfo[]> {
    let subfolders: string[] = [];
    try {
      subfolders = await this._client.getFolderSubfolderNames(this._siteUrl, this.getUnitFolderPath(unit));
    } catch {
      subfolders = [];
    }
    const existing = new Set(subfolders.map((n) => n.toLowerCase()));

    const readDay = async (dateStr: string): Promise<ReportDayInfo> => {
      if (!existing.has(getReportFolderName(docId, dateStr).toLowerCase())) {
        return { dateStr, status: 'missing' };
      }
      const folder = getReportFolderRelativePath(this.getLibraryPath(), unit, docId, dateStr);
      let files: ISharePointFileInfo[] = [];
      try {
        files = await this._client.getFolderFiles(this._siteUrl, folder);
      } catch {
        files = [];
      }
      const reportName = getReportFileName(docId, dateStr).toLowerCase();
      const report = files.filter((f) => f.name.toLowerCase() === reportName)[0];
      if (!report) return { dateStr, status: 'missing' };
      const signed = this.pickSignedPdf(files);
      return {
        dateStr,
        status: signed ? 'done' : 'unsigned',
        savedAt: report.timeLastModified,
        signed
      };
    };

    // Đọc theo từng nhóm `concurrency` ngày để không gửi quá nhiều yêu cầu cùng lúc
    const out: ReportDayInfo[] = [];
    for (let i = 0; i < dateStrs.length; i += concurrency) {
      out.push(...(await Promise.all(dateStrs.slice(i, i + concurrency).map(readDay))));
    }
    return out;
  }

  /** File PDF đã ký của một ngày (undefined nếu chưa có). */
  public async getSignedPdf(unit: string, docId: string, dateStr: string): Promise<SignedDocInfo | undefined> {
    const folder = getReportFolderRelativePath(this.getLibraryPath(), unit, docId, dateStr);
    const files = await this._client.getFolderFiles(this._siteUrl, folder);
    return this.pickSignedPdf(files);
  }

  /**
   * Upload file PDF đã ký vào thư mục báo cáo, luôn đổi tên theo quy ước <DocCode>_signed.pdf
   * (ghi đè nếu đã có). Mọi file `*_signed.pdf` khác (tên cũ của hệ thống) được xoá để không còn hai bản.
   */
  public async uploadSignedPdf(
    unit: string,
    docId: string,
    dateStr: string,
    buffer: ArrayBuffer,
    docCode = ''
  ): Promise<SignedDocInfo> {
    await this.ensureUnitFolder(unit);
    const folder = getReportFolderRelativePath(this.getLibraryPath(), unit, docId, dateStr);
    await this._client.ensureFolder(this._siteUrl, folder);
    const name = getSignedPdfFileName(docId, unit, dateStr, docCode);
    const res = await this._client.uploadFile(this._siteUrl, folder, name, buffer, true);

    try {
      const files = await this._client.getFolderFiles(this._siteUrl, folder);
      const olds = files.filter((f) => isSignedPdfName(f.name) && f.name.toLowerCase() !== name.toLowerCase());
      for (const old of olds) await this._client.deleteFile(this._siteUrl, old.serverRelativeUrl);
    } catch (err: unknown) {
      console.warn('Could not remove legacy signed file:', err);
    }

    return {
      name,
      serverRelativeUrl: res.serverRelativeUrl,
      length: buffer.byteLength,
      timeLastModified: new Date().toISOString()
    };
  }

  /**
   * Delete an uploaded photo
   */
  public async deleteReportPhoto(serverRelativeUrl: string): Promise<void> {
    await this._client.deleteFile(this._siteUrl, serverRelativeUrl);
  }
}

