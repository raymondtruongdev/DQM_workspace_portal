import { SharePointClient } from '../../../core/services/SharePointClient';
import { CatalogRow, normKey } from '../utils/reportCatalog';

/** Tên SharePoint List chứa danh mục người dùng được tự thêm (cột cần tạo: xem AGENTS.md mục 5.3). */
export const REPORT_CATALOG_LIST = 'CatalogItems';

/** Tên (tiêu đề) các cột của list; tên nội bộ thật được tra từ list vì có thể khác. */
const COLS = ['Title', 'Unit', 'Catalog', 'Code', 'Link'] as const;
type Col = (typeof COLS)[number];

const esc = (v: string): string => v.replace(/'/g, "''");

/** Mục có phải đã tồn tại không (SharePoint từ chối vì khoá duy nhất `Title` bị trùng). */
function isDuplicateError(err: unknown): boolean {
  const m = err instanceof Error ? err.message : '';
  return /duplicate|unique|already exists|\(409\)|-2130575169|-2147024809/i.test(m);
}

/**
 * Danh mục người dùng được tự thêm (SharePoint List `CatalogItems`, dùng chung cho mọi nhà ga và mọi loại danh mục):
 * Element, Resource nhân lực/thiết bị, Vật tư… Mỗi dòng = một mục. Chống trùng hai lớp: kiểm tra trước khi thêm
 * và khoá duy nhất `Title` của list (`<Unit>|<Catalog>|code|<mã chuẩn hoá>` hoặc `…|link|<giá trị chuẩn hoá>`).
 */
export class ReportCatalogService {
  private _client: SharePointClient;
  private _siteUrl: string;
  private _names: Promise<Record<Col, string>> | undefined;

  constructor(client: SharePointClient, siteUrl: string) {
    this._client = client;
    this._siteUrl = siteUrl;
  }

  /** Khoá duy nhất của một mục; `kind` = code khi mục có mã, link khi mục chỉ có giá trị ràng buộc (vd vật tư theo tên). */
  public static keyOf(unit: string, table: string, row: Pick<CatalogRow, 'code' | 'link'>): string {
    return row.code
      ? `${unit}|${table.toLowerCase()}|code|${normKey(row.code)}`
      : `${unit}|${table.toLowerCase()}|link|${normKey(row.link)}`;
  }

  private names(): Promise<Record<Col, string>> {
    if (!this._names) {
      this._names = this._client
        .getListFieldMap(this._siteUrl, REPORT_CATALOG_LIST)
        .then((map) => {
          const out = {} as Record<Col, string>;
          COLS.forEach((c) => {
            out[c] = map[c.toLowerCase()] || c;
          });
          out.Title = 'Title';
          return out;
        })
        .catch((err: unknown) => {
          this._names = undefined;
          throw err;
        });
    }
    return this._names;
  }

  private rowOf(n: Record<Col, string>, raw: Record<string, unknown>): CatalogRow {
    const str = (c: Col): string => (raw[n[c]] === undefined || raw[n[c]] === null ? '' : String(raw[n[c]]).trim());
    return { table: str('Catalog').toLowerCase(), code: str('Code'), link: str('Link') };
  }

  /** Mọi mục của một nhà ga. Lỗi (chưa tạo list, thiếu quyền…) thì ném ra để nơi gọi báo cảnh báo. */
  public async getRows(unit: string): Promise<CatalogRow[]> {
    const n = await this.names();
    const raw = await this._client.queryItems<Record<string, unknown>>(
      this._siteUrl,
      REPORT_CATALOG_LIST,
      `$filter=${n.Unit} eq '${esc(unit)}'&$top=5000`
    );
    return raw.map((r) => this.rowOf(n, r)).filter((r) => r.table && (r.code || r.link));
  }

  /**
   * Thêm một mục. Đã có mục cùng khoá (do người khác vừa thêm cùng lúc) thì không báo lỗi:
   * trả về mục sẵn có với `created = false`.
   */
  public async addRow(unit: string, row: CatalogRow): Promise<{ row: CatalogRow; created: boolean }> {
    const n = await this.names();
    const key = ReportCatalogService.keyOf(unit, row.table, row);
    const find = async (): Promise<CatalogRow | undefined> => {
      const found = await this._client.queryItems<Record<string, unknown>>(
        this._siteUrl,
        REPORT_CATALOG_LIST,
        `$filter=${n.Title} eq '${esc(key)}'&$top=1`
      );
      return found[0] ? this.rowOf(n, found[0]) : undefined;
    };
    const existing = await find();
    if (existing) return { row: existing, created: false };
    try {
      await this._client.createItem<Record<string, unknown>>(this._siteUrl, REPORT_CATALOG_LIST, {
        [n.Title]: key,
        [n.Unit]: unit,
        [n.Catalog]: row.table.toLowerCase(),
        [n.Code]: row.code,
        [n.Link]: row.link
      });
      return { row, created: true };
    } catch (err: unknown) {
      if (!isDuplicateError(err)) throw err;
      const again = await find();
      return { row: again || row, created: false };
    }
  }
}
