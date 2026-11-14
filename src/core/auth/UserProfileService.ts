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
    if (rawList.length > 0) {
      console.warn('[UserProfileService] Tên cột nội bộ của dòng đầu:', Object.keys(rawList[0]), 'bản đồ tên cột:', map);
    }
    return rawList
      .map((raw) => {
        const norm: Record<string, unknown> = { ...raw };
        wanted.forEach((name) => {
          const internal = map[name.toLowerCase()];
          if (internal && raw[internal] !== undefined && raw[internal] !== null) norm[name] = raw[internal];
        });
        // Dự phòng: không tìm được cột email theo tên thì lấy cột văn bản đầu tiên chứa "@".
        if (!String(norm.Email ?? '').trim()) {
          const found = Object.keys(raw).find((k) => typeof raw[k] === 'string' && String(raw[k]).indexOf('@') > 0);
          if (found) norm.Email = raw[found];
        }
        return mapSharePointUserProfile(norm);
      })
      .filter((p) => p.active);
  }
}
