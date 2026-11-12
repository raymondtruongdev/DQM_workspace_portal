import { ReportDoc, TemplateVersionInfo } from '../types/report.types';
import { parseReportDocFromBuffer } from '../utils/reportCodec';
import { applyFormDefaults } from '../utils/reportForm';
import { formatReportNo, syncMetadata } from '../utils/reportMeta';
import { ReportFileService } from './ReportFileService';

export class ReportTemplateService {
  private _fileService: ReportFileService;

  constructor(fileService: ReportFileService) {
    this._fileService = fileService;
  }

  /**
   * Load latest versioned template for Unit & DocID (e.g. _template/template_ST01_20261109.xlsx)
   * and initialize a ReportDoc
   */
  public async getTemplate(
    docId: string,
    unit: string,
    dateStr: string,
    currentUserEmail: string,
    siteManager?: string,
    preloadedLatest?: { info?: TemplateVersionInfo; buffer?: ArrayBuffer }
  ): Promise<{ doc: ReportDoc; templateBuffer?: ArrayBuffer; latestTemplateInfo?: TemplateVersionInfo }> {
    const latest = preloadedLatest || (await this._fileService.findLatestTemplateForUnit(unit, docId));
    const buffer = latest.buffer;
    const latestInfo = latest.info;

    let doc: ReportDoc;
    if (buffer) {
      doc = await parseReportDocFromBuffer(buffer, docId, unit, dateStr);
    } else if (docId !== 'WeeklyReport') {
      throw new Error(`Chưa có template cho ${unit} trong thư mục _template/. Vui lòng tải template lên SharePoint.`);
    } else {
      doc = {
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
          reportNo: '00001',
          subcontractor: '',
          email: currentUserEmail,
          updateTime: new Date().toISOString()
        },
        sections: [],
        tasks: []
      };
    }

    // Apply unit-specific metadata and template version
    if (latestInfo && latestInfo.version) {
      doc.metadata.templateVersion = latestInfo.version;
      doc.latestTemplateInfo = latestInfo;
    } else if (!doc.metadata.templateVersion) {
      doc.metadata.templateVersion = `default_${unit}`;
    }

    // Báo cáo đầu tiên của đơn vị (chưa có báo cáo trước): số báo cáo bắt đầu từ 00001
    if (!doc.metadata.reportNo) doc.metadata.reportNo = formatReportNo(1);
    doc.metadata = syncMetadata(doc.metadata, unit);
    doc.metadata.email = currentUserEmail;
    doc.metadata.updateTime = new Date().toISOString();
    doc.dateStr = dateStr;
    doc.templateUpdateRequired = false;
    doc = applyFormDefaults(doc, { unit, email: currentUserEmail, siteManager });

    return {
      doc,
      templateBuffer: buffer,
      latestTemplateInfo: latestInfo
    };
  }
}
