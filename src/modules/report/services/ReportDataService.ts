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
import { Clock, LOCAL_CLOCK } from '../../../core/services/ServerClock';
import { DEFAULT_REPORT_SETTINGS, ReportSettings, isPastLocked, vnDateAsLocal, vnDateStr } from '../utils/reportSettings';
import { ReportFileService } from './ReportFileService';
import { LoadedTemplate, ReportTemplateService } from './ReportTemplateService';
import { IReportLock, ISignedTrail, ReportLockedError, ReportRegistryService, canUploadSignedWhileLocked } from './ReportRegistryService';
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

  private readonly _settings: () => Promise<ReportSettings>;
  private readonly _clock: Clock;

  constructor(
    fileService: ReportFileService,
    templateService: ReportTemplateService,
    registry: ReportRegistryService,
    catalogSource: ReportCatalogService,
    settings: () => Promise<ReportSettings> = () => Promise.resolve(DEFAULT_REPORT_SETTINGS),
    clock: Clock = LOCAL_CLOCK
  ) {
    this._settings = settings;
    this._clock = clock;
    this._fileService = fileService;
    this._templateService = templateService;
    this._registry = registry;
    this._catalogSource = catalogSource;
  }

  private _catalogFile?: { at: number; wb: Promise<{ XLSX: typeof import('xlsx'); wb: import('xlsx').WorkBook } | undefined> };

  /**
   * Danh mục (Activity P6, cấu kiện, tài nguyên, vật tư, thời tiết, trạng thái, đơn vị thi công…) của đơn vị cho các dropdown.
   * Nguồn là list DRPCatalogItems + file ActivityP6.xlsx (Activity, chỉ admin ghi). Không đọc được thì trả
   * danh mục rỗng/một phần kèm cảnh báo (form vẫn dùng được, chỉ giữ giá trị đã lưu).
   */
  public async getCatalog(unit: string, force = false): Promise<ReportCatalog> {
    const base = await this.getCatalogFromFile(unit, force);
    try {
      return mergeCatalogRows(base, await this._catalogSource.getRows(unit));
    } catch (err: unknown) {
      console.warn('[Report] Could not read DRPCatalogItems:', err);
      return { ...base, warnings: base.warnings.concat('Chưa đọc được danh mục dùng chung (kiểm tra list DRPCatalogItems và quyền đọc).') };
    }
  }

  /** Thêm một mục vào danh mục dùng chung (người dùng thêm bằng "Thêm mới…"); trả mục lưu được và `created` = false nếu đã có. */
  public updateCatalogLink(unit: string, table: string, code: string, link: string): Promise<boolean> {
    return this._catalogSource.updateLink(unit, table, code, link);
  }

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
      return emptyCatalog(unit, `Không đọc được file danh mục ActivityP6.xlsx. ${errorDetail(err)}`);
    }
  }

  /** Báo cáo ngày mới có số báo cáo theo sổ DRPRegistry (báo cáo tuần là placeholder nên bỏ qua). */
  /**
   * Trạng thái khoá của báo cáo một ngày (chỉ báo cáo ngày). Báo cáo tuần hoặc không đọc được sổ thì undefined
   * (coi là không khoá, form vẫn dùng được; việc ghi vẫn được kiểm tra lại ở `assertUnlocked`).
   */
  public async getLockState(unit: string, docId: string, dateStr: string): Promise<IReportLock | undefined> {
    if (!this.isDaily(docId)) return undefined;
    try {
      return await this._registry.getLock(unit, docId, dateStr);
    } catch (err: unknown) {
      console.warn('[DRPRegistry] Could not read lock state:', err);
      return undefined;
    }
  }

  /** Admin bật/tắt khoá báo cáo một ngày. Nơi gọi phải `checkMaintenance()` trước. Lỗi (chưa có dòng, thiếu cột Locked, thiếu quyền) thì ném ra. */
  public async setReportLocked(unit: string, docId: string, dateStr: string, locked: boolean): Promise<void> {
    const settings = await this._settings();
    await this._registry.setLocked(unit, docId, dateStr, locked, settings.unlockHours);
  }

  /** Mọi thao tác ghi vào báo cáo của một ngày phải qua đây: báo cáo đang khoá thì ném `ReportLockedError`. Không đọc được sổ thì cho qua. */
  private async assertUnlocked(unit: string, docId: string, dateStr: string): Promise<void> {
    const lock = await this.getLockState(unit, docId, dateStr);
    if (lock && lock.locked) throw new ReportLockedError();
  }

  private isDaily(docId: string): boolean {
    return !docId.toLowerCase().includes('week');
  }

  /**
   * Gán số báo cáo từ sổ DRPRegistry (số lớn nhất của đơn vị + 1, hoặc số đã có của ngày này).
   * Không đọc được sổ (chưa tạo list, không có quyền…) thì giữ số suy ra từ báo cáo ngày trước.
   */
  private async applyRegistryNo(doc: ReportDoc): Promise<ReportDoc> {
    if (!this.isDaily(doc.docId)) return doc;
    try {
      const no = await this._registry.getReportNoFor(doc.unit, doc.docId, doc.dateStr);
      return { ...doc, metadata: syncMetadata({ ...doc.metadata, reportNo: no }, doc.unit) };
    } catch (err: unknown) {
      console.warn('[DRPRegistry] Could not read next report no., using previous report:', err);
      return doc;
    }
  }

  /**
   * Gắn bộ đếm mã tự tăng (Event ID, Issue ID…) từ sổ DRPRegistry vào báo cáo để dòng mới nhận số kế tiếp liên tục giữa các ngày,
   * và điền các mã còn trống/còn nguyên mẫu "{n5}". Không đọc được sổ thì chỉ dựa vào các dòng đang có trong báo cáo.
   */
  private async applyIdCounters(doc: ReportDoc): Promise<ReportDoc> {
    if (!this.isDaily(doc.docId)) return doc;
    let counters: Record<string, number> = {};
    try {
      counters = await this._registry.getCounters(doc.unit, doc.docId);
    } catch (err: unknown) {
      console.warn('[DRPRegistry] Could not read ID counters:', err);
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
   * 1. Đã có data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>.json: đọc đúng file đó (báo cáo cũ hơn template mới nhất thì
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
      // Số báo cáo tự nhảy theo từng đơn vị: số của báo cáo ngày trước + 1 (rồi theo sổ DRPRegistry)
      doc = { ...doc, metadata: syncMetadata({ ...doc.metadata, reportNo: nextReportNo(doc.metadata.reportNo) }, unit) };
      doc = applyFormDefaults(doc, ctx);
      doc = await this.applyRegistryNo(doc);
      return { doc: this.flagOutdated(doc, latest, 'previous_day'), isExisting: false };
    }

    if (!latest) {
      if (!this.isDaily(docId)) return { doc: this.emptyWeeklyDoc(docId, unit, dateStr, currentUserEmail), isExisting: false };
      throw new Error(`Chưa có template cho ${unit} trong thư mục _template/ (template_DRP_${unit}_<yyyyMMdd>.json) và cũng chưa có template chung template_DRP_default.json. Vui lòng tải template lên SharePoint.`);
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
        geoTag: '',
        subcontractors: [],
        attachedDocs: '',
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
   * Lưu báo cáo: ghi `.json` (nguồn dữ liệu, kiểm tra ETag) rồi `.html` (bản xem/in độc lập nhúng dữ liệu) cạnh nhau,
   * sau đó ghi sổ DRPRegistry. Lỗi ghi HTML hoặc sổ chỉ báo cảnh báo, không chặn việc lưu dữ liệu.
   */
  public async saveReport(
    doc: ReportDoc,
    userEmail: string
  ): Promise<{ serverRelativeUrl: string; eTag?: string; metadata: ReportMetadata; registryWarning?: string; htmlWarning?: string }> {
    await this.assertUnlocked(doc.unit, doc.docId, doc.dateStr);
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
        console.warn('[DRPRegistry] Could not record save:', err);
        registryWarning = `Đã lưu báo cáo nhưng chưa ghi được vào sổ DRPRegistry (kiểm tra list và quyền ghi). ${errorDetail(err)}`;
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
   * Server-relative folder holding the report's photos: data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/photos
   */
  public getPhotosFolderPath(unit: string, docId: string, dateStr: string): string {
    return this._fileService.getReportPhotosFolderPath(unit, docId, dateStr);
  }

  /**
   * Upload a photo into data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/photos/<groupId>/; returns the stored file name
   */
  public async uploadPhoto(
    unit: string,
    docId: string,
    dateStr: string,
    groupId: string,
    fileName: string,
    buffer: ArrayBuffer
  ): Promise<string> {
    await this.assertUnlocked(unit, docId, dateStr);
    return await this._fileService.uploadReportPhoto(unit, docId, dateStr, groupId, fileName, buffer);
  }

  /** Ảnh cùng tên đã có trong thư mục ảnh của nhóm chưa (không đọc được thư mục thì coi là chưa có). */
  public async checkPhoto(unit: string, docId: string, dateStr: string, groupId: string, fileName: string): Promise<{ name: string; exists: boolean }> {
    try {
      return await this._fileService.photoExists(unit, docId, dateStr, groupId, fileName);
    } catch (err: unknown) {
      console.warn('Could not check existing photo:', err);
      return { name: fileName, exists: false };
    }
  }

  /** Thư mục file PDF đính kèm sự kiện (server-relative): <thư mục báo cáo>/pdf/eventlog */
  public getEventPdfFolderPath(unit: string, docId: string, dateStr: string): string {
    return this._fileService.getReportEventPdfFolderPath(unit, docId, dateStr);
  }

  /** Tên file PDF sẽ lưu (giữ nguyên tên gốc) và nó đã tồn tại trong pdf/eventlog của báo cáo chưa. Không đọc được thư mục thì coi là chưa có. */
  public async checkEventPdf(unit: string, docId: string, dateStr: string, fileName: string): Promise<{ name: string; exists: boolean }> {
    const name = ReportFileService.eventPdfStoredName(fileName);
    try {
      return { name, exists: await this._fileService.eventPdfExists(unit, docId, dateStr, name) };
    } catch (err: unknown) {
      console.warn('Could not check existing event PDF:', err);
      return { name, exists: false };
    }
  }

  /** Upload PDF đính kèm của một dòng sự kiện (giữ nguyên tên gốc, cùng tên thì ghi đè); trả về tên file đã lưu. Báo cáo đang khoá thì ném `ReportLockedError`. */
  public async uploadEventPdf(unit: string, docId: string, dateStr: string, fileName: string, buffer: ArrayBuffer): Promise<string> {
    await this.assertUnlocked(unit, docId, dateStr);
    return await this._fileService.uploadEventPdf(unit, docId, dateStr, fileName, buffer);
  }

  /** Xoá một PDF đính kèm sự kiện theo tên file. Báo cáo đang khoá thì ném `ReportLockedError`. */
  public async deleteEventPdf(unit: string, docId: string, dateStr: string, fileName: string): Promise<void> {
    await this.assertUnlocked(unit, docId, dateStr);
    await this._fileService.deleteReportPhoto(`${this._fileService.getReportEventPdfFolderPath(unit, docId, dateStr)}/${fileName}`);
  }

  /**
   * Delete an uploaded photo by its server-relative URL
   */
  public async deletePhoto(serverRelativeUrl: string): Promise<void> {
    await this._fileService.deleteReportPhoto(serverRelativeUrl);
  }

  /**
   * Trạng thái các ngày (YYYY-MM-DD) của một đơn vị: chưa nhập / chưa upload bản ký / hoàn tất.
   * Đọc từ sổ DRPRegistry (cùng nguồn với tab Tổng quan nên hai màn hình luôn khớp nhau);
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
        const settings = await this._settings();
        await this._clock.ensureFresh();
        const today = vnDateStr(this._clock.now());
        // Ngày chưa có dòng trong sổ: ngày cũ vẫn bị khoá khi bật chế độ "chỉ nhập ngày hiện tại"
        return dateStrs.map((d): ReportDayInfo => days[d] || { dateStr: d, status: 'missing', locked: isPastLocked(settings, d, today) || undefined });
      } catch (err: unknown) {
        console.warn('[DRPRegistry] Could not read day infos, scanning folders instead:', err);
      }
    }
    // Quét thư mục (sổ không đọc được): khoá ngày cũ theo chế độ "chỉ nhập ngày hiện tại" (không biết khoá riêng từng ngày)
    const infos = await this._fileService.getReportDayInfos(unit, docId, dateStrs, concurrency);
    if (!this.isDaily(docId)) return infos;
    const settings = await this._settings();
    await this._clock.ensureFresh();
    const today = vnDateStr(this._clock.now());
    return infos.map((i) => (isPastLocked(settings, i.dateStr, today) ? { ...i, locked: true } : i));
  }

  /** Thời điểm hiện tại theo giờ máy chủ SharePoint (ms UTC). */
  public now(): number {
    return this._clock.now();
  }

  /** Ngày hôm nay (YYYY-MM-DD) theo giờ Việt Nam, tính từ giờ máy chủ. */
  public today(): string {
    return vnDateStr(this._clock.now());
  }

  /** Ngày hôm nay theo giờ Việt Nam dạng Date (0h giờ máy) để dựng lịch (`buildRecentDates`). */
  public todayDate(): Date {
    return vnDateAsLocal(this._clock.now());
  }

  /** Cài đặt báo cáo của admin (số ngày thống kê, chế độ chỉ nhập ngày hiện tại). */
  public getSettings(): Promise<ReportSettings> {
    return this._settings();
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
    docCode = '',
    uploader = ''
  ): Promise<{ info: SignedDocInfo; registryWarning?: string }> {
    // Báo cáo khoá: vẫn upload được bản ký lần đầu khi báo cáo đã Lưu (không sửa được nội dung); đã có bản ký thì admin mở khoá mới được thay
    const lock = await this.getLockState(unit, docId, dateStr);
    if (lock && lock.locked && !canUploadSignedWhileLocked(lock)) throw new ReportLockedError();
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
        await this._registry.recordSigned(unit, docId, dateStr, reportNo, reporter, uploader);
      } catch (err: unknown) {
        console.warn('[DRPRegistry] Could not record signed:', err);
        registryWarning = `Đã upload bản ký nhưng chưa cập nhật được sổ DRPRegistry (kiểm tra list và quyền ghi). ${errorDetail(err)}`;
      }
    }
    return { info, registryWarning };
  }

  /**
   * Xoá bản đã ký của một ngày: xoá file PDF rồi đưa sổ DRPRegistry về "Chưa upload bản ký", ghi thời điểm + người xoá
   * GẦN NHẤT (cột SignedRemovedAt/SignedRemovedBy). Xoá file lỗi thì ném ra; ghi sổ lỗi chỉ trả cảnh báo.
   * Nơi gọi phải `checkMaintenance()` trước.
   */
  public async removeSignedDocument(unit: string, docId: string, dateStr: string, actor: string): Promise<{ warning?: string }> {
    await this.assertUnlocked(unit, docId, dateStr); // đang khoá thì admin phải mở khoá trước khi xoá bản ký
    await this._fileService.deleteSignedPdfs(unit, docId, dateStr);
    if (!this.isDaily(docId)) return {};
    try {
      const logged = await this._registry.recordUnsigned(unit, docId, dateStr, actor);
      if (!logged) {
        return { warning: 'Đã xoá bản đã ký nhưng chưa ghi được người xoá: list DRPRegistry chưa có cột SignedRemovedAt và SignedRemovedBy.' };
      }
    } catch (err: unknown) {
      console.warn('[DRPRegistry] Could not record unsigned:', err);
      return { warning: `Đã xoá bản đã ký nhưng chưa cập nhật được sổ DRPRegistry (kiểm tra list và quyền ghi). ${errorDetail(err)}` };
    }
    return {};
  }

  /** Dấu vết bản đã ký của một ngày (người upload, lần xoá gần nhất); không có hoặc không đọc được thì undefined. */
  public async getSignedTrail(unit: string, docId: string, dateStr: string): Promise<ISignedTrail | undefined> {
    if (!this.isDaily(docId)) return undefined;
    try {
      return await this._registry.getSignedTrail(unit, docId, dateStr);
    } catch {
      return undefined;
    }
  }

  /**
   * Trạng thái báo cáo ngày của mọi đơn vị từ ngày `fromDate` (YYYY-MM-DD), đọc một lần từ sổ DRPRegistry
   */
  public async getRegistryDayInfos(
    docId: string,
    fromDate: string
  ): Promise<Record<string, Record<string, ReportDayInfo>>> {
    return await this._registry.getDayInfosSince(docId, fromDate);
  }
}
