import { TemplateVersionInfo } from '../types/report.types';
import { ReportTemplateJson } from '../types/reportJson.types';
import { parseTemplateJson } from '../utils/reportJson';
import { ReportFileService } from './ReportFileService';

export interface LoadedTemplate {
  info: TemplateVersionInfo;
  template: ReportTemplateJson;
}

/**
 * Template báo cáo (JSON) trong thư viện: `_template/template_<DRP|WRP>_<Unit>_<yyyyMMdd>.json`, ngày trong tên file càng mới càng ưu tiên.
 * Template quyết định cách nhập (kiểu ô, danh mục, công thức, mặc định, luỹ kế…) và các dòng mẫu của báo cáo mới.
 */
export class ReportTemplateService {
  private _fileService: ReportFileService;

  constructor(fileService: ReportFileService) {
    this._fileService = fileService;
  }

  /** Template mới nhất của đơn vị; undefined nếu chưa có. File sai cấu trúc thì ném lỗi nêu tên file. */
  public async getLatest(unit: string, docId: string): Promise<LoadedTemplate | undefined> {
    const found = await this._fileService.findLatestTemplateForUnit(unit, docId);
    if (!found.info || found.text === undefined) return undefined;
    try {
      return { info: found.info, template: parseTemplateJson(found.text) };
    } catch (err: unknown) {
      throw new Error(`${found.info.fileName}: ${err instanceof Error ? err.message : 'không đọc được template'}`);
    }
  }
}
