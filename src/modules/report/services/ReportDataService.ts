import { ReportDayInfo, ReportDoc, ReportMetadata, SignedDocInfo, TemplateVersionInfo } from '../types/report.types';
import { parseReportDocFromBuffer, encodeReportDocToBuffer } from '../utils/reportCodec';
import { applyFormDefaults } from '../utils/reportForm';
import { nextReportNo, syncMetadata } from '../utils/reportMeta';
import { carryForwardTable, collectAutoIdCounters, fillAutoIds } from '../utils/reportTable';
import { isTemplateVersionOutdated } from '../utils/reportFileNaming';
import { CatalogRow, ReportCatalog, emptyCatalog, mergeCatalogRows, parseCatalogFromWorkbook } from '../utils/reportCatalog';
import { ReportFileService } from './ReportFileService';
import { ReportTemplateService } from './ReportTemplateService';
import { ReportRegistryService } from './ReportRegistryService';
import { ReportCatalogService } from './ReportCatalogService';

export interface LoadedReportResult {
  doc: ReportDoc;
  templateBuffer?: ArrayBuffer;
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
  private _catalogList: ReportCatalogService;

  constructor(
    fileService: ReportFileService,
    templateService: ReportTemplateService,
    registry: ReportRegistryService,
    catalogList: ReportCatalogService
  ) {
    this._fileService = fileService;
    this._templateService = templateService;
    this._registry = registry;
    this._catalogList = catalogList;
  }

  private _catalogFile?: { at: number; wb: Promise<{ XLSX: typeof import('xlsx'); wb: import('xlsx').WorkBook } | undefined> };

  /**
   * Danh mục (Activity P6, cấu kiện, tài nguyên, vật tư…) của đơn vị (file ActivityP6.xlsx ở gốc thư viện). File nhớ 5 phút trong bộ nhớ trang;
   * `force` bỏ qua bộ nhớ. Không đọc được file thì trả danh mục rỗng kèm cảnh báo (form vẫn dùng được, chỉ giữ giá trị cũ).
   */
  public async getCatalog(unit: string, force = false): Promise<ReportCatalog> {
    const fromFile = await this.getCatalogFromFile(unit, force);
    // Bảng người dùng được tự thêm nằm ở list CatalogItems; không đọc được thì chỉ cảnh báo, vẫn dùng phần từ file
    try {
      return mergeCatalogRows(fromFile, await this._catalogList.getRows(unit));
    } catch (err: unknown) {
      console.warn('[Report] Could not read CatalogItems list:', err);
      return { ...fromFile, warnings: fromFile.warnings.concat('Chưa đọc được list CatalogItems (kiểm tra list và quyền đọc).') };
    }
  }

  /** Thêm một mục vào list CatalogItems (người dùng thêm bằng "Khác…"); trả về mục lưu được và `created` = false nếu đã có. */
  public addCatalogItem(unit: string, row: CatalogRow): Promise<{ row: CatalogRow; created: boolean }> {
    return this._catalogList.addRow(unit, row);
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
      if (!parsed) {
        return emptyCatalog(unit, 'Chưa có file danh mục ActivityP6.xlsx ở gốc thư viện.');
      }
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
      sections: doc.sections.map((s) => (s.kind === 'table' ? fillAutoIds(s, doc.unit, counters) : s))
    };
  }

  /**
   * Load report for unit, docId and dateStr.
   * 1. Checks latest template in _template/ (e.g. template_ST01_20261109.xlsx).
   * 2. If existing file is found in data/<Unit>/<DocID>_<yyyyMMdd>/, loads it and checks if its templateVersion is outdated.
   * 3. Otherwise checks adjacent / closest previous day's report:
   *    - If previous report's templateVersion matches the latest template -> loads previous day's data directly without prompting.
   *    - If previous report's templateVersion is older than the latest template -> sets templateUpdateRequired = true so the UI prompts the user to load the new template first.
   * 4. If no previous report exists, loads fresh from the latest unit template.
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
    const [existingFile, latestTemplate] = await Promise.all([
      this._fileService.getReportFileWithMeta(unit, docId, dateStr),
      this._fileService
        .findLatestTemplateForUnit(unit, docId)
        .catch((): { info?: TemplateVersionInfo; buffer?: ArrayBuffer } => ({}))
    ]);

    const templateBuffer = latestTemplate.buffer;
    const latestTemplateInfo = latestTemplate.info;

    if (existingFile && existingFile.buffer) {
      let doc = await parseReportDocFromBuffer(existingFile.buffer, docId, unit, dateStr);
      doc.eTag = existingFile.eTag;
      doc.isExisting = true;
      doc.latestTemplateInfo = latestTemplateInfo;
      doc = applyFormDefaults(doc, { unit, email: currentUserEmail, siteManager });

      const isOutdated = isTemplateVersionOutdated(doc.metadata.templateVersion, latestTemplateInfo);
      if (isOutdated) {
        doc.templateUpdateRequired = true;
        doc.outdatedSourceType = 'existing';
        doc.outdatedTemplateVersion = doc.metadata.templateVersion || '';
      } else {
        doc.templateUpdateRequired = false;
        if (latestTemplateInfo?.version && !doc.metadata.templateVersion) {
          doc.metadata.templateVersion = latestTemplateInfo.version;
        }
      }

      return {
        doc,
        templateBuffer,
        isExisting: true
      };
    }

    // Try loading from the adjacent / closest previous day's report first
    const prevReport = await this._fileService.getPreviousReportFile(unit, docId, dateStr);
    if (prevReport && prevReport.buffer) {
      let doc = await parseReportDocFromBuffer(prevReport.buffer, docId, unit, dateStr);
      doc.dateStr = dateStr;
      doc.isExisting = false;
      doc.eTag = undefined;
      doc.copiedFromDate = prevReport.sourceDateStr;
      // Ảnh thuộc về từng ngày: không sao chép sang báo cáo mới
      // Luỹ kế: số liệu của ngày trước thành "luỹ kế hôm qua"; ô thực hiện trong ngày về 0 (cột carry/reset của template)
      doc.sections = doc.sections.map((sec) =>
        sec.kind === 'photo' ? { ...sec, groups: [] } : sec.kind === 'table' ? carryForwardTable(sec) : sec
      );
      // Số báo cáo tự nhảy theo từng đơn vị: số của báo cáo ngày trước + 1
      doc.metadata = syncMetadata({ ...doc.metadata, reportNo: nextReportNo(doc.metadata.reportNo) }, unit);
      doc.latestTemplateInfo = latestTemplateInfo;
      doc.metadata.email = currentUserEmail;
      doc = applyFormDefaults(doc, { unit, email: currentUserEmail, siteManager });
      doc.metadata.updateTime = new Date().toISOString();
      doc = await this.applyRegistryNo(doc);

      const isOutdated = isTemplateVersionOutdated(doc.metadata.templateVersion, latestTemplateInfo);
      if (isOutdated) {
        doc.templateUpdateRequired = true;
        doc.outdatedSourceType = 'previous_day';
        doc.outdatedTemplateVersion = doc.metadata.templateVersion || '';
      } else {
        doc.templateUpdateRequired = false;
        if (latestTemplateInfo?.version) {
          doc.metadata.templateVersion = latestTemplateInfo.version;
        }
      }

      return {
        doc,
        templateBuffer,
        isExisting: false
      };
    }

    const templateResult = await this._templateService.getTemplate(
      docId,
      unit,
      dateStr,
      currentUserEmail,
      siteManager,
      latestTemplate
    );


    return {
      doc: await this.applyRegistryNo(templateResult.doc),
      templateBuffer: templateResult.templateBuffer,
      isExisting: false
    };
  }

  /**
   * Explicitly load the latest unit template from _template/ when the user confirms upgrading from an outdated template
   */
  public async loadLatestTemplateForReport(
    docId: string,
    unit: string,
    dateStr: string,
    currentUserEmail: string,
    siteManager?: string,
    currentDoc?: ReportDoc
  ): Promise<LoadedReportResult> {
    const templateResult = await this._templateService.getTemplate(
      docId,
      unit,
      dateStr,
      currentUserEmail,
      siteManager
    );

    const newDoc = templateResult.doc;
    if (currentDoc) {
      newDoc.eTag = currentDoc.eTag;
      newDoc.isExisting = currentDoc.isExisting;
      // Giữ số báo cáo của báo cáo hiện tại (không quay về 00001)
      newDoc.metadata = syncMetadata({ ...newDoc.metadata, reportNo: currentDoc.metadata.reportNo }, currentDoc.unit);
      // Giữ lại ảnh đã tải của báo cáo hiện tại khi chuyển sang template mới
      newDoc.sections = newDoc.sections.map((sec) => {
        if (sec.kind !== 'photo') return sec;
        const old = currentDoc.sections.filter((x) => x.kind === 'photo' && x.key === sec.key)[0];
        return old && old.kind === 'photo' ? { ...sec, groups: old.groups } : sec;
      });
    }
    newDoc.templateUpdateRequired = false;
    newDoc.copiedFromDate = undefined;

    return {
      doc: await this.applyIdCounters(currentDoc && currentDoc.isExisting ? newDoc : await this.applyRegistryNo(newDoc)),
      templateBuffer: templateResult.templateBuffer,
      isExisting: Boolean(currentDoc?.isExisting)
    };
  }

  /**
   * Save report: encodes the workbook, uploads it, then records it in the ReportRegistry list.
   * New reports re-read the registry first so the report no. is the latest one at the moment of saving.
   * `registryWarning` is set when the file was saved but the registry could not be updated.
   */
  public async saveReport(
    doc: ReportDoc,
    templateBuffer?: ArrayBuffer
  ): Promise<{ serverRelativeUrl: string; eTag?: string; metadata: ReportMetadata; registryWarning?: string }> {
    const base = doc.isExisting ? doc : await this.applyRegistryNo(doc);
    const toSave: ReportDoc = { ...base, metadata: { ...base.metadata, updateTime: new Date().toISOString() } };

    const buffer = await encodeReportDocToBuffer(toSave, templateBuffer);
    const res = await this._fileService.saveReportFile(
      doc.unit,
      doc.docId,
      doc.dateStr,
      buffer,
      doc.eTag
    );

    let registryWarning: string | undefined;
    if (this.isDaily(doc.docId)) {
      try {
        // Bộ đếm cộng dồn: số lớn nhất của báo cáo này và của các báo cáo trước (để ngày/báo cáo sau không dùng lại số)
        const counters = collectAutoIdCounters(toSave.sections, doc.unit, toSave.idCounters);
        await this._registry.recordSave(
          doc.unit,
          doc.docId,
          doc.dateStr,
          toSave.metadata.reportNo,
          toSave.metadata.email,
          counters
        );
      } catch (err: unknown) {
        console.warn('[ReportRegistry] Could not record save:', err);
        registryWarning = `Đã lưu báo cáo nhưng chưa ghi được vào sổ ReportRegistry (kiểm tra list và quyền ghi). ${errorDetail(err)}`;
      }
    }
    return { ...res, metadata: toSave.metadata, registryWarning };
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

  /**
   * File PDF đã ký của một ngày (undefined nếu chưa có)
   */
  /** Doc. code ghi trong file báo cáo đã lưu ('' nếu chưa có file). */
  public async getSavedDocCode(unit: string, docId: string, dateStr: string): Promise<string> {
    const file = await this._fileService.getReportFileWithMeta(unit, docId, dateStr);
    if (!file || !file.buffer) return '';
    const parsed = await parseReportDocFromBuffer(file.buffer, docId, unit, dateStr);
    return parsed.metadata.docCode || '';
  }

  /** Mở file PDF đã ký của một ngày trong tab mới (tên file chỉ biết khi đọc thư mục). */
  public async openSignedDocument(unit: string, docId: string, dateStr: string): Promise<boolean> {
    const signed = await this._fileService.getSignedPdf(unit, docId, dateStr);
    if (!signed) return false;
    window.open(signed.serverRelativeUrl, '_blank', 'noopener');
    return true;
  }

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
