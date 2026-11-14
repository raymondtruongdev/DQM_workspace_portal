import { UserRoleResolution } from '../../../core/types';
import { ISharePointClient, RoleResolutionService } from '../../../core/services';
import { REPORT_CONFIG } from '../constants/reportConfig';
import { RoleListNames, parseRoleRows } from '../utils/reportRole';

/** Tên SharePoint List phân quyền báo cáo (cột: xem docs/SHAREPOINT_SETUP.md và Report_Role_list.xlsx). */
export const REPORT_ROLE_LIST = 'Report_Role';

export class ReportRoleService {
  private _client: ISharePointClient;
  private _siteUrl: string;

  constructor(client: ISharePointClient, siteUrl: string) {
    this._client = client;
    this._siteUrl = siteUrl;
  }

  /** Tên nội bộ của các cột (tra theo tên hiển thị; DisplayName có thể là cột Title của list). */
  private async names(): Promise<RoleListNames> {
    const map = await this._client.getListFieldMap(this._siteUrl, REPORT_ROLE_LIST);
    const pick = (display: string, fallback = display): string => map[display.toLowerCase()] || fallback;
    return {
      displayName: pick('DisplayName', 'Title'),
      email: pick('Email'),
      unit: pick('Unit'),
      docId: pick('DocID'),
      docTitle: pick('DocTitle'),
      active: pick('Active')
    };
  }

  /**
   * Phân quyền báo cáo của người dùng, đọc từ list Report_Role (đọc lại mỗi lần mở module). Admin thấy mọi phân quyền đang hiệu lực.
   * Không đọc được list (chưa tạo, thiếu quyền đọc…) thì ném lỗi để giao diện báo cho người dùng.
   */
  public async resolveUserRoles(rawUserEmail: string, isAdmin: boolean): Promise<UserRoleResolution> {
    const names = await this.names();
    const rows = await this._client.getItems<Record<string, unknown>>(this._siteUrl, REPORT_ROLE_LIST, 5000);
    const valid = new Set<string>(REPORT_CONFIG.VALID_DOC_IDS.map((d) => d.toLowerCase()));
    const all = parseRoleRows(rows, names).filter((r) => valid.has(r.docId.toLowerCase()));
    return RoleResolutionService.resolveRoles(all, rawUserEmail, isAdmin);
  }
}
