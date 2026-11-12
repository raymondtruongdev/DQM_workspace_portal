import { ReportDayInfo, ReportDoc, SignedDocInfo, TemplateVersionInfo } from '../types/report.types';
import { parseReportDocFromBuffer, encodeReportDocToBuffer } from '../utils/reportCodec';
import { applyFormDefaults } from '../utils/reportForm';
import { nextReportNo, syncMetadata } from '../utils/reportMeta';
import { carryForwardTable } from '../utils/reportTable';
import { isTemplateVersionOutdated } from '../utils/reportFileNaming';
import { ReportFileService } from './ReportFileService';
import { ReportTemplateService } from './ReportTemplateService';

export interface LoadedReportResult {
  doc: ReportDoc;
  templateBuffer?: ArrayBuffer;
  isExisting: boolean;
}

export class ReportDataService {
  private _fileService: ReportFileService;
  private _templateService: ReportTemplateService;

  constructor(fileService: ReportFileService, templateService: ReportTemplateService) {
    this._fileService = fileService;
    this._templateService = templateService;
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
      doc: templateResult.doc,
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
      doc: newDoc,
      templateBuffer: templateResult.templateBuffer,
      isExisting: Boolean(currentDoc?.isExisting)
    };
  }

  /**
   * Save report: encodes 6-sheet buffer and uploads to SharePoint library
   */
  public async saveReport(
    doc: ReportDoc,
    templateBuffer?: ArrayBuffer
  ): Promise<{ serverRelativeUrl: string; eTag?: string }> {
    doc.metadata.updateTime = new Date().toISOString();

    const buffer = await encodeReportDocToBuffer(doc, templateBuffer);
    return await this._fileService.saveReportFile(
      doc.unit,
      doc.docId,
      doc.dateStr,
      buffer,
      doc.eTag
    );
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
   * Trạng thái các ngày (YYYY-MM-DD) của một đơn vị: chưa nhập / chưa upload bản ký / hoàn tất
   */
  public async getDayInfos(
    unit: string,
    docId: string,
    dateStrs: string[],
    concurrency?: number
  ): Promise<ReportDayInfo[]> {
    return await this._fileService.getReportDayInfos(unit, docId, dateStrs, concurrency);
  }

  /**
   * File PDF đã ký của một ngày (undefined nếu chưa có)
   */
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
    buffer: ArrayBuffer
  ): Promise<SignedDocInfo> {
    return await this._fileService.uploadSignedPdf(unit, docId, dateStr, buffer);
  }
}
