import { ISharePointClient } from '../../../core/services';
import { parseActive } from '../utils/reportRole';
import { parseStartDate } from '../utils/reportMeta';

/** Tên SharePoint List cấu hình từng nhà ga (một dòng mỗi nhà ga; cột: xem docs/SHAREPOINT_SETUP.md và DRPUnitConfig_list.xlsx). */
export const REPORT_UNIT_CONFIG_LIST = 'DRPUnitConfig';

/** Cấu hình của một nhà ga. `startDate` rỗng = chưa cấu hình hoặc sai định dạng (không có số báo cáo). */
export interface IUnitConfig {
  unit: string;
  /** `yyyy-MM-dd`: ngày được đánh số báo cáo 00001. */
  startDate?: string;
  /** Tọa độ `lat, long` tự điền vào báo cáo; rỗng = chưa có. */
  geoTag: string;
}

const CACHE_MS = 3 * 60 * 1000;

const clean = (v: unknown): string => (v === undefined || v === null ? '' : String(v).trim());

/**
 * Đọc list DRPUnitConfig (người báo cáo chỉ đọc, chỉ admin ghi). Đọc cả list một lần rồi nhớ 3 phút trong bộ nhớ trang.
 * Không đọc được list (chưa tạo, thiếu quyền…) thì ném lỗi để nơi gọi báo rõ, không im lặng coi như chưa cấu hình.
 */
export class ReportUnitConfigService {
  private _client: ISharePointClient;
  private _siteUrl: string;
  private _cache?: { at: number; promise: Promise<Record<string, IUnitConfig>> };

  constructor(client: ISharePointClient, siteUrl: string) {
    this._client = client;
    this._siteUrl = siteUrl;
  }

  private async load(): Promise<Record<string, IUnitConfig>> {
    const map = await this._client.getListFieldMap(this._siteUrl, REPORT_UNIT_CONFIG_LIST);
    const col = (display: string, fallback = display): string => map[display.toLowerCase()] || fallback;
    const [cUnit, cStart, cGeo, cActive] = [col('Unit', 'Title'), col('StartDate'), col('GeoTag'), col('Active')];
    const rows = await this._client.getItems<Record<string, unknown>>(this._siteUrl, REPORT_UNIT_CONFIG_LIST, 5000);
    const out: Record<string, IUnitConfig> = {};
    rows.forEach((row) => {
      const unit = clean(row[cUnit]) || clean(row.Title);
      if (!unit || !parseActive(row[cActive])) return;
      const key = unit.toUpperCase();
      if (out[key]) return; // trùng nhà ga: dòng đầu tiên thắng
      out[key] = { unit, startDate: parseStartDate(row[cStart]), geoTag: clean(row[cGeo]) };
    });
    return out;
  }

  /** Cấu hình của mọi nhà ga (khoá = mã nhà ga viết HOA). `force` bỏ qua bộ nhớ đệm. */
  public getAll(force = false): Promise<Record<string, IUnitConfig>> {
    if (force || !this._cache || Date.now() - this._cache.at > CACHE_MS) {
      const promise = this.load();
      this._cache = { at: Date.now(), promise };
      promise.catch(() => {
        if (this._cache && this._cache.promise === promise) this._cache = undefined; // lỗi thì lần sau đọc lại
      });
    }
    return this._cache.promise;
  }

  /** Cấu hình của một nhà ga; chưa có dòng (hoặc `Active = 0`) thì undefined. */
  public async get(unit: string): Promise<IUnitConfig | undefined> {
    return (await this.getAll())[unit.toUpperCase()];
  }
}
