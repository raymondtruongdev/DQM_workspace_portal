import { SharePointClient } from './SharePointClient';
import { DEFAULT_PORTAL_CONFIG, PortalConfig } from '../types/portalConfig.types';

const CONFIG_FOLDER = 'SiteAssets';
const CONFIG_FILE = 'portal_config.json';

/** Đọc/ghi cấu hình portal (chế độ bảo trì) trong SiteAssets/portal_config.json của site. */
export class PortalConfigService {
  constructor(private readonly _client: SharePointClient, private readonly _siteUrl: string) {}

  private _folderPath(): string {
    let rel = '';
    try {
      rel = new URL(this._siteUrl).pathname;
    } catch {
      rel = this._siteUrl;
    }
    return `${rel.replace(/\/+$/, '')}/${CONFIG_FOLDER}`;
  }

  /** Lấy cấu hình; chưa có file hoặc lỗi đọc thì coi như không bảo trì (không khoá nhầm người dùng). */
  public async getConfig(): Promise<PortalConfig> {
    try {
      const buf = await this._client.getFileBuffer(this._siteUrl, `${this._folderPath()}/${CONFIG_FILE}`);
      if (!buf) return DEFAULT_PORTAL_CONFIG;
      const raw = JSON.parse(new TextDecoder('utf-8').decode(buf)) as Partial<PortalConfig>;
      return {
        maintenance: raw.maintenance === true,
        message: typeof raw.message === 'string' ? raw.message : '',
        updatedBy: typeof raw.updatedBy === 'string' ? raw.updatedBy : '',
        updatedAt: typeof raw.updatedAt === 'string' ? raw.updatedAt : ''
      };
    } catch (err: unknown) {
      console.warn('[PortalConfigService] Could not read portal config:', err);
      return DEFAULT_PORTAL_CONFIG;
    }
  }

  /** Bật/tắt chế độ bảo trì và ghi lại ai đổi, khi nào. */
  public async setMaintenance(on: boolean, message: string, updatedBy: string): Promise<PortalConfig> {
    const config: PortalConfig = {
      maintenance: on,
      message: message.trim(),
      updatedBy,
      updatedAt: new Date().toISOString()
    };
    const bytes = new TextEncoder().encode(JSON.stringify(config, null, 2));
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    await this._client.uploadFile(this._siteUrl, this._folderPath(), CONFIG_FILE, buffer, true);
    return config;
  }
}
