import { SharePointClient } from './SharePointClient';
import { DEFAULT_PORTAL_CONFIG, PortalConfig } from '../types/portalConfig.types';

const CONFIG_FOLDER = 'SiteAssets';
const CONFIG_FILE = 'portal_config.json';

function parseModuleSettings(raw: unknown): Record<string, Record<string, unknown>> {
  const out: Record<string, Record<string, unknown>> = {};
  if (!raw || typeof raw !== 'object') return out;
  Object.keys(raw as Record<string, unknown>).forEach((id) => {
    const v = (raw as Record<string, unknown>)[id];
    if (v && typeof v === 'object' && !Array.isArray(v)) out[id] = v as Record<string, unknown>;
  });
  return out;
}

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
        updatedAt: typeof raw.updatedAt === 'string' ? raw.updatedAt : '',
        moduleSettings: parseModuleSettings(raw.moduleSettings)
      };
    } catch (err: unknown) {
      console.warn('[PortalConfigService] Could not read portal config:', err);
      return DEFAULT_PORTAL_CONFIG;
    }
  }

  private async _write(config: PortalConfig): Promise<void> {
    const bytes = new TextEncoder().encode(JSON.stringify(config, null, 2));
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    await this._client.uploadFile(this._siteUrl, this._folderPath(), CONFIG_FILE, buffer, true);
  }

  /** Bật/tắt chế độ bảo trì và ghi lại ai đổi, khi nào (giữ nguyên cài đặt của các module). */
  public async setMaintenance(on: boolean, message: string, updatedBy: string): Promise<PortalConfig> {
    const current = await this.getConfig();
    const config: PortalConfig = {
      maintenance: on,
      message: message.trim(),
      updatedBy,
      updatedAt: new Date().toISOString(),
      moduleSettings: current.moduleSettings
    };
    await this._write(config);
    return config;
  }

  /** Ghi cài đặt của một module (thay toàn bộ phần của module đó), giữ nguyên bảo trì và cài đặt module khác. */
  public async setModuleSettings(moduleId: string, settings: Record<string, unknown>, updatedBy: string): Promise<PortalConfig> {
    const current = await this.getConfig();
    const config: PortalConfig = {
      ...current,
      updatedBy,
      updatedAt: new Date().toISOString(),
      moduleSettings: { ...current.moduleSettings, [moduleId]: settings }
    };
    await this._write(config);
    return config;
  }
}
