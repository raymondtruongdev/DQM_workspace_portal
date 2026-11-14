import { ReportDayInfo, ReportDoc, ReportMetadata, ReportSection, SignedDocInfo, TemplateVersionInfo } from '../types/report.types';
import { ReportJson } from '../types/reportJson.types';
import { applyTemplate, carryForwardDoc, docFromTemplate, docToJson, jsonToDoc, parseReportJson, serializeReportJson } from '../utils/reportJson';
import { applyFormDefaults } from '../utils/reportForm';
import { formatReportNo, nextReportNo, syncMetadata } from '../utils/reportMeta';
import { collectAutoIdCounters, fillAutoIds } from '../utils/reportTable';
import { isTemplateVersionOutdated } from '../utils/reportFileNaming';
import { CatalogRow, ReportCatalog, emptyCatalog, mergeCatalogRows, parseCatalogFromWorkbook } from '../utils/reportCatalog';
import { buildStandaloneHtml, ReportHtmlAssets } from '../utils/reportHtml';
import { bufferToDataUrl, embeddableImage, mimeOfFile } from '../utils/reportImage';
import { ReportFileService } from './ReportFileService';
import { LoadedTemplate, ReportTemplateService } from './ReportTemplateService';
import { ReportRegistryService } from './ReportRegistryService';
import { ReportCatalogService } from './ReportCatalogService';

export interface LoadedReportResult {
  doc: ReportDoc;
  isExisting: boolean;
}

/** Chi tiết lỗi SharePoint (rút gọn) để người dùng/admin biết cột nào sai hoặc thiếu quyền. */
function errorDetail(err: unknown): string {
  const raw = err instanceof Error ? err.message : String(err);
  const msg = /"message"\s*:\s*\{[^}]*"value"\s*:\s*"([^"]+)"/.exec(raw);
  return `Chi tiết: ${(msg ? msg[1] : raw).slice(0, 220)}`;
}

export class ReportDataService {
  private _fileService: ReportFileService;
  private _templateService: ReportTemplateService;
  private _registry: ReportRegistryService;
  private _catalogSource: ReportCatalogService;

  constructor(
    fileService: ReportFileService,
    templateService: ReportTemplateService,
    registry: ReportRegistryService,
    catalogSource: ReportCatalogService
  ) {
    this._fileService = fileService;
    this._templateService = templateService;
    this._registry = registry;
    this._catalogSource = catalogSource;
  }

  private _catalogFile?: { at: number; wb: Promise<{ XLSX: typeof import('xlsx'); wb: import('xlsx').WorkBook } | undefined> };

  /**
   * Danh mục (Activity P6, cấu kiện, tài nguyên, vật tư, thời tiết, trạng thái, đơn vị thi công…) của đơn vị cho các dropdown.
   * Nguồn là list CatalogItems + file ActivityP6.xlsx (Activity, chỉ admin ghi). Không đọc được thì trả
   * danh mục rỗng/một phần kèm cảnh báo (form vẫn dùng được, chỉ giữ giá trị đã lưu).
   */
  public async getCatalog(unit: string, force = false): Promise<ReportCatalog> {
    const base = await this.getCatalogFromFile(unit, force);
    try {
      return mergeCatalogRows(base, await this._catalogSource.getRows(unit));
    } catch (err: unknown) {
      console.warn('[Report] Could not read CatalogItems:', err);
      return { ...base, warnings: base.warnings.concat('Chưa đọc được danh mục dùng chung (kiểm tra list CatalogItems và quyền đọc).') };
    }
  }

  /** Thêm một mục vào danh mục dùng chung (người dùng thêm bằng "Khác…"); trả mục lưu được và `created` = false nếu đã có. */
  public addCatalogItem(unit: string, row: CatalogRow): Promise<{ row: CatalogRow; created: boolean }> {
    return this._catalogSource.addRow(unit, row);
  }

  private async getCatalogFromFile(unit: string, force: boolean): Promise<ReportCatalog> {
    const now = Date.now();
    if (force || !this._catalogFile || now - this._catalogFile.at > 5 * 60 * 1000) {
      this._catalogFile = {
        at: now,
        wb: (async () => {
          const buffer = await this._fileService.getActivityFileBuffer();
          if (!buffer) return undefined;
          const XLSX = await import(/* webpackChunkName: 'xlsx' */ 'xlsx');
          return { XLSX, wb: XLSX.read(buffer, { type: 'array' }) };
        })()
      };
    }
    try {
      const parsed = await this._catalogFile.wb;
      if (!parsed) return emptyCatalog(unit, 'Chưa có file danh mục ActivityP6.xlsx ở gốc thư viện.');
      return parseCatalogFromWorkbook(parsed.XLSX, parsed.wb, unit);
    } catch (err: unknown) {
      this._catalogFile = undefined;
      console.warn('[Report] Could not read ActivityP6.xlsx:', err);
      return emptyCatalog(unit, 'Không đọc được file danh mục ActivityP6.xlsx.');
    }
  }

  /** Báo cáo ngày mới có số báo cáo theo sổ ReportRegistry (báo cáo tuần là placeholder nên bỏ qua). */
  private isDaily(docId: string): boolean {
    return !docId.toLowerCase().includes('week');
  }

  /**
   * Gán số báo cáo từ sổ ReportRegistry (số lớn nhất của đơn vị + 1, hoặc số đã có của ngày này).
   * Không đọc được sổ (chưa tạo list, không có quyền…) thì giữ số suy ra từ báo cáo ngày trước.
   */
  private async applyRegistryNo(doc: ReportDoc): Promise<ReportDoc> {
    if (!this.isDaily(doc.docId)) return doc;
    try {
      const no = await this._registry.getReportNoFor(doc.unit, doc.docId, doc.dateStr);
      return { ...doc, metadata: syncMetadata({ ...doc.metadata, reportNo: no }, doc.unit) };
    } catch (err: unknown) {
      console.warn('[ReportRegistry] Could not read next report no., using previous report:', err);
      return doc;
    }
  }

  /**
   * Gắn bộ đếm mã tự tăng (Event ID, Issue ID…) từ sổ ReportRegistry vào báo cáo để dòng mới nhận số kế tiếp liên tục giữa các ngày,
   * và điền các mã còn trống/còn nguyên mẫu "{n5}". Không đọc được sổ thì chỉ dựa vào các dòng đang có trong báo cáo.
   */
  private async applyIdCounters(doc: ReportDoc): Promise<ReportDoc> {
    if (!this.isDaily(doc.docId)) return doc;
    let counters: Record<string, number> = {};
    try {
      counters = await this._registry.getCounters(doc.unit, doc.docId);
    } catch (err: unknown) {
      console.warn('[ReportRegistry] Could not read ID counters:', err);
    }
    return {
      ...doc,
      idCounters: counters,
      sections: doc.sections.map((s): ReportSection => (s.kind === 'table' ? fillAutoIds(s, doc.unit, counters) : s))
    };
  }

  /** Cờ cập nhật template: báo cáo (đã lưu hoặc kế thừa) dùng template cũ hơn bản mới nhất trong _template/. */
  private flagOutdated(doc: ReportDoc, latest: LoadedTemplate | undefined, source: 'existing' | 'previous_day'): ReportDoc {
    const info: TemplateVersionInfo | undefined = latest ? latest.info : undefined;
    if (isTemplateVersionOutdated(doc.metadata.templateVersion, info)) {
      return { ...doc, latestTemplateInfo: info, templateUpdateRequired: true, outdatedSourceType: source, outdatedTemplateVersion: doc.metadata.templateVersion || '' };
    }
    const metadata: ReportMetadata =
      info && info.version && (source === 'previous_day' || !doc.metadata.templateVersion)
        ? { ...doc.metadata, templateVersion: info.version }
        : doc.metadata;
    return { ...doc, metadata, latestTemplateInfo: info, templateUpdateRequired: false };
  }

  /**
   * Mở báo cáo của đơn vị/ngày:
   * 1. Đã có data/<Unit>/<DocID>_<yyyyMMdd>/<DocID>_<yyyyMMdd>.data.json: đọc đúng file đó (báo cáo cũ hơn template mới nhất thì
   *    báo cần cập nhật template).
   * 2. Chưa có: lấy báo cáo ngày trước gần nhất làm nền (luỹ kế/reset theo template, số báo cáo + 1, bỏ ảnh).
   * 3. Đơn vị chưa có báo cáo nào: tạo từ template mới nhất trong _template/.
   */
  public async loadReport(
    docId: string,
    unit: string,
    dateStr: string,
    currentUserEmail: string,
    siteManager?: string
  ): Promise<LoadedReportResult> {
    const result = await this.loadReportCore(docId, unit, dateStr, currentUserEmail, siteManager);
    return { ...result, doc: await this.applyIdCounters(result.doc) };
  }

  private async loadReportCore(
    docId: string,
    unit: string,
    dateStr: string,
    currentUserEmail: string,
    siteManager?: string
  ): Promise<LoadedReportResult> {
    const ctx = { unit, email: currentUserEmail, siteManager };
    const [existing, latest] = await Promise.all([
      this._fileService.getReportJsonWithMeta(unit, docId, dateStr),
      this._templateService.getLatest(unit, docId)
    ]);

    if (existing) {
      let doc = jsonToDoc(parseReportJson(existing.text), latest ? latest.template : undefined);
      doc = { ...doc, eTag: existing.eTag, isExisting: true };
      doc = applyFormDefaults(doc, ctx);
      return { doc: this.flagOutdated(doc, latest, 'existing'), isExisting: true };
    }

    const prev = await this._fileService.getPreviousReportJson(unit, docId, dateStr);
    if (prev) {
      const prevDoc = jsonToDoc(parseReportJson(prev.text), latest ? latest.template : undefined);
      let doc = carryForwardDoc({ ...prevDoc, dateStr: prevDoc.dateStr }, dateStr, currentUserEmail);
      doc = { ...doc, copiedFromDate: prev.sourceDateStr };
      // Số báo cáo tự nhảy theo từng đơn vị: số của báo cáo ngày trước + 1 (rồi theo sổ ReportRegistry)
      doc = { ...doc, metadata: syncMetadata({ ...doc.metadata, reportNo: nextReportNo(doc.metadata.reportNo) }, unit) };
      doc = applyFormDefaults(doc, ctx);
      doc = await this.applyRegistryNo(doc);
      return { doc: this.flagOutdated(doc, latest, 'previous_day'), isExisting: false };
    }

    if (!latest) {
      if (!this.isDaily(docId)) return { doc: this.emptyWeeklyDoc(docId, unit, dateStr, currentUserEmail), isExisting: false };
      throw new Error(`Chưa có template cho ${unit} trong thư mục _template/ (template_DRP_${unit}_<yyyyMMdd>.json). Vui lòng tải template lên SharePoint.`);
    }
    let doc = docFromTemplate(latest.template, { docId, unit, dateStr, email: currentUserEmail, siteManager, version: latest.info.version });
    doc = { ...doc, latestTemplateInfo: latest.info };
    return { doc: await this.applyRegistryNo(doc), isExisting: false };
  }

  /** Báo cáo tuần (placeholder): chưa có template JSON nên chỉ là khung trống. */
  private emptyWeeklyDoc(docId: string, unit: string, dateStr: string, email: string): ReportDoc {
    return {
      docId,
      unit,
      dateStr,
      metadata: {
        docTitle: 'BÁO CÁO TUẦN',
        docTitleEn: 'WEEKLY REPORT',
        subTitle: '',
        docCode: '',
        templateVersion: `default_${unit}`,
        location: unit,
        reportNo: formatReportNo(1),
        subcontractors: [],
        email,
        updateTime: new Date().toISOString()
      },
      sections: [],
      tasks: []
    };
  }

  /**
   * Áp dụng template mới nhất lên báo cáo đang mở khi người dùng xác nhận (ghép theo `id` mục/cột/trường, giữ dữ liệu cũ).
   */
  public async applyLatestTemplate(doc: ReportDoc, email: string, siteManager?: string): Promise<ReportDoc> {
    const latest = await this._templateService.getLatest(doc.unit, doc.docId);
    if (!latest) throw new Error(`Chưa có template cho ${doc.unit} trong thư mục _template/.`);
    const next = applyTemplate(doc, latest.template, latest.info.version, { email, siteManager });
    return await this.applyIdCounters({ ...next, latestTemplateInfo: latest.info });
  }

  /**
   * Lưu báo cáo: ghi `.data.json` (nguồn dữ liệu, kiểm tra ETag) rồi `.html` (bản xem/in độc lập nhúng dữ liệu) cạnh nhau,
   * sau đó ghi sổ ReportRegistry. Lỗi ghi HTML hoặc sổ chỉ báo cảnh báo, không chặn việc lưu dữ liệu.
   */
  public async saveReport(
    doc: ReportDoc,
    userEmail: string
  ): Promise<{ serverRelativeUrl: string; eTag?: string; metadata: ReportMetadata; registryWarning?: string; htmlWarning?: string }> {
    const base = doc.isExisting ? doc : await this.applyRegistryNo(doc);
    const now = new Date().toISOString();
    // updated = người đang lưu; created chỉ ghi ở lần lưu đầu tiên và giữ nguyên các lần sau
    const metadata: ReportMetadata = {
      ...base.metadata,
      email: userEmail || base.metadata.email,
      updateTime: now,
      created: base.metadata.created && base.isExisting ? base.metadata.created : { by: userEmail || base.metadata.email, at: now }
    };
    const toSave: ReportDoc = { ...base, metadata };
    const json = docToJson(toSave);
    const text = serializeReportJson(json);

    const res = await this._fileService.saveReportJson(doc.unit, doc.docId, doc.dateStr, text, doc.eTag);

    let htmlWarning: string | undefined;
    try {
      const html = await buildStandaloneHtml(toSave, text, await this.collectHtmlAssets(toSave, json));
      await this._fileService.saveReportHtml(doc.unit, doc.docId, doc.dateStr, html);
    } catch (err: unknown) {
      console.warn('[Report] Could not write standalone HTML:', err);
      htmlWarning = `Đã lưu dữ liệu báo cáo nhưng chưa tạo được file HTML xem độc lập. ${errorDetail(err)}`;
    }

    let registryWarning: string | undefined;
    if (this.isDaily(doc.docId)) {
      try {
        // Bộ đếm cộng dồn: số lớn nhất của báo cáo này và của các báo cáo trước (để ngày/báo cáo sau không dùng lại số)
        const counters = collectAutoIdCounters(toSave.sections, doc.unit, toSave.idCounters);
        await this._registry.recordSave(doc.unit, doc.docId, doc.dateStr, metadata.reportNo, metadata.email, counters);
      } catch (err: unknown) {
        console.warn('[ReportRegistry] Could not record save:', err);
        registryWarning = `Đã lưu báo cáo nhưng chưa ghi được vào sổ ReportRegistry (kiểm tra list và quyền ghi). ${errorDetail(err)}`;
      }
    }
    return { ...res, metadata, registryWarning, htmlWarning };
  }

  /** Logo và ảnh hiện trường (đã thu nhỏ) để nhúng vào file HTML; ảnh/logo không đọc được thì bỏ qua (HTML hiện chữ/khung trống). */
  private async collectHtmlAssets(doc: ReportDoc, json: ReportJson): Promise<ReportHtmlAssets> {
    const logos: ReportHtmlAssets['logos'] = {};
    const readLogo = async (file: string): Promise<string | undefined> => {
      try {
        const buf = await this._fileService.getResourceFileBuffer(file);
        return buf ? bufferToDataUrl(buf, mimeOfFile(file)) : undefined;
      } catch {
        return undefined;
      }
    };
    [logos.thaco, logos.crec] = await Promise.all([readLogo('thaco_logo.jpg'), readLogo('crec_logo.jpg')]);

    const photos: Record<string, string> = {};
    for (const section of json.sections) {
      if (section.type !== 'photo') continue;
      for (const row of section.rows || []) {
        const files = (row as { files?: string[] }).files || [];
        for (const name of files) {
          try {
            const buf = await this._fileService.getReportPhotoBuffer(doc.unit, doc.docId, doc.dateStr, row.id, name);
            if (buf) photos[`${row.id}/${name}`] = await embeddableImage(buf, name);
          } catch (err: unknown) {
            console.warn(`[Report] Could not embed photo ${row.id}/${name}:`, err);
          }
        }
      }
    }
    return { logos, photos };
  }

  /**
   * Server-relative folder holding shared resources (logos): <library>/resource
   */
  public getResourceFolderPath(): string {
    return `${this._fileService.getLibraryPath()}/resource`;
  }

  /**
   * Server-relative folder holding the report's photos: data/<Unit>/<DocID>_<yyyyMMdd>/photos
   */
  public getPhotosFolderPath(unit: string, docId: string, dateStr: string): string {
    return this._fileService.getReportPhotosFolderPath(unit, docId, dateStr);
  }

  /**
   * Upload a photo into data/<Unit>/<DocID>_<yyyyMMdd>/photos/<groupId>/; returns the stored file name
   */
  public async uploadPhoto(
    unit: string,
    docId: string,
    dateStr: string,
    groupId: string,
    fileName: string,
    buffer: ArrayBuffer
  ): Promise<string> {
    return await this._fileService.uploadReportPhoto(unit, docId, dateStr, groupId, fileName, buffer);
  }

  /**
   * Delete an uploaded photo by its server-relative URL
   */
  public async deletePhoto(serverRelativeUrl: string): Promise<void> {
    await this._fileService.deleteReportPhoto(serverRelativeUrl);
  }

  /**
   * Trạng thái các ngày (YYYY-MM-DD) của một đơn vị: chưa nhập / chưa upload bản ký / hoàn tất.
   * Đọc từ sổ ReportRegistry (cùng nguồn với tab Tổng quan nên hai màn hình luôn khớp nhau);
   * chưa tạo được/không đọc được sổ thì tạm quét thư mục data/ (có thể lệch với Tổng quan).
   */
  public async getDayInfos(
    unit: string,
    docId: string,
    dateStrs: string[],
    concurrency?: number
  ): Promise<ReportDayInfo[]> {
    if (this.isDaily(docId) && dateStrs.length > 0) {
      try {
        const from = dateStrs.slice().sort()[0];
        const byKey = await this._registry.getDayInfosSince(docId, from, unit);
        const days = byKey[`${unit}|${docId}`.toLowerCase()] || {};
        return dateStrs.map((d): ReportDayInfo => days[d] || { dateStr: d, status: 'missing' });
      } catch (err: unknown) {
        console.warn('[ReportRegistry] Could not read day infos, scanning folders instead:', err);
      }
    }
    return await this._fileService.getReportDayInfos(unit, docId, dateStrs, concurrency);
  }

  /** Doc. code ghi trong file báo cáo đã lưu ('' nếu chưa có file). */
  public async getSavedDocCode(unit: string, docId: string, dateStr: string): Promise<string> {
    const file = await this._fileService.getReportJsonWithMeta(unit, docId, dateStr);
    return file ? parseReportJson(file.text).docCode || '' : '';
  }

  /** Mở file PDF đã ký của một ngày trong tab mới (tên file chỉ biết khi đọc thư mục). */
  public async openSignedDocument(unit: string, docId: string, dateStr: string): Promise<boolean> {
    const signed = await this._fileService.getSignedPdf(unit, docId, dateStr);
    if (!signed) return false;
    window.open(signed.serverRelativeUrl, '_blank', 'noopener');
    return true;
  }

  /** File PDF đã ký của một ngày (undefined nếu chưa có) */
  public async getSignedDocument(unit: string, docId: string, dateStr: string): Promise<SignedDocInfo | undefined> {
    return await this._fileService.getSignedPdf(unit, docId, dateStr);
  }

  /**
   * Upload file PDF đã ký vào thư mục báo cáo của ngày
   */
  public async uploadSignedDocument(
    unit: string,
    docId: string,
    dateStr: string,
    buffer: ArrayBuffer,
    reportNo = '',
    reporter = '',
    docCode = ''
  ): Promise<{ info: SignedDocInfo; registryWarning?: string }> {
    // Upload từ lịch không biết Doc. code: đọc từ file báo cáo đã lưu
    let code = docCode;
    if (!code) {
      try {
        code = await this.getSavedDocCode(unit, docId, dateStr);
      } catch (err: unknown) {
        console.warn('Could not read doc code for signed file name:', err);
      }
    }
    const info = await this._fileService.uploadSignedPdf(unit, docId, dateStr, buffer, code);
    let registryWarning: string | undefined;
    if (this.isDaily(docId)) {
      try {
        await this._registry.recordSigned(unit, docId, dateStr, reportNo, reporter);
      } catch (err: unknown) {
        console.warn('[ReportRegistry] Could not record signed:', err);
        registryWarning = `Đã upload bản ký nhưng chưa cập nhật được sổ ReportRegistry (kiểm tra list và quyền ghi). ${errorDetail(err)}`;
      }
    }
    return { info, registryWarning };
  }

  /**
   * Trạng thái báo cáo ngày của mọi đơn vị từ ngày `fromDate` (YYYY-MM-DD), đọc một lần từ sổ ReportRegistry
   */
  public async getRegistryDayInfos(
    docId: string,
    fromDate: string
  ): Promise<Record<string, Record<string, ReportDayInfo>>> {
    return await this._registry.getDayInfosSince(docId, fromDate);
  }
}
