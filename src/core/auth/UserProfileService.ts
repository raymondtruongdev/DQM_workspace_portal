import { SharePointClient } from '../services/SharePointClient';
import type { UserProfileRecord } from './auth.types';
import { mapSharePointUserProfile } from './authUtils';

export const USER_PROFILES_LIST = 'UserProfiles';

export class UserProfileService {
  private _client: SharePointClient;

  constructor(client: SharePointClient) {
    this._client = client;
  }

  /**
   * Đọc hồ sơ từ list UserProfiles. Dòng Active = No bị loại (coi như không có trong list: không vào được portal,
   * không hiện trong bộ chuyển vai trò của admin).
   */
  public async getUserProfiles(siteUrl: string, listTitle = USER_PROFILES_LIST): Promise<UserProfileRecord[]> {
    const rawList = await this._client.getItems<Record<string, unknown>>(siteUrl, listTitle, 1000);
    // Tên nội bộ của cột có thể khác tên hiển thị (vd Email0, field_1 khi list tạo từ Excel hoặc đổi tên cột): tra theo tên hiển thị.
    let map: Record<string, string> = {};
    try {
      map = await this._client.getListFieldMap(siteUrl, listTitle);
    } catch (e) {
      console.warn('[UserProfileService] Không tra được tên cột, dùng tên mặc định:', e);
    }
    const wanted = ['Email', 'DisplayName', 'Department', 'DepartmentName', 'Role', 'Active'];
    return rawList
      .map((raw) => {
        const norm: Record<string, unknown> = { ...raw };
        wanted.forEach((name) => {
          const internal = map[name.toLowerCase()];
          if (internal && raw[internal] !== undefined && raw[internal] !== null) norm[name] = raw[internal];
        });
        // Dự phòng CHỈ khi list không có cột tên Email: lấy cột văn bản đầu tiên chứa "@" (tránh lấy nhầm cột khi đã có cột Email).
        if (map.email === undefined && !String(norm.Email ?? '').trim()) {
          const found = Object.keys(raw).find((k) => typeof raw[k] === 'string' && String(raw[k]).indexOf('@') > 0);
          if (found) norm.Email = raw[found];
        }
        // Import từ Excel: SharePoint lấy cột chữ đầu tiên (DisplayName) làm cột Title nên list không có cột DisplayName.
        // Khi đã có cột Email riêng và không có cột DisplayName thì Title chính là tên hiển thị.
        if (map.displayname === undefined && map.email !== undefined && !String(norm.DisplayName ?? '').trim()) {
          const title = raw.Title;
          if (title !== undefined && title !== null && String(title).trim() !== '') norm.DisplayName = title;
        }
        return mapSharePointUserProfile(norm);
      })
      .filter((p) => p.active);
  }
}
