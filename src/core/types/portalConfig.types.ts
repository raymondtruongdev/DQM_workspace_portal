/** Cấu hình dùng chung của portal, lưu trong portal_config.json trên SharePoint. */
export interface PortalConfig {
  /** Chế độ bảo trì: người dùng thường chỉ thấy màn hình bảo trì, admin vẫn vào được. */
  maintenance: boolean;
  /** Nội dung hiển thị trên màn hình bảo trì. */
  message: string;
  updatedBy: string;
  updatedAt: string;
}

export const DEFAULT_PORTAL_CONFIG: PortalConfig = {
  maintenance: false,
  message: '',
  updatedBy: '',
  updatedAt: ''
};
