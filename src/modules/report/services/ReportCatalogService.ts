import { ISharePointClient } from '../../../core/services/SharePointClient';
import { CatalogRow, normKey } from '../utils/reportCatalog';
import { parseActive } from '../utils/reportRole';

/** Tên SharePoint List chứa danh mục người dùng được tự thêm (cột cần tạo: xem AGENTS.md mục 5.3). */
export const REPORT_CATALOG_LIST = 'DRPCatalogItems';

/** Tên (tiêu đề) các cột của list; tên nội bộ thật được tra từ list vì có thể khác. */
const COLS = ['Title', 'Unit', 'Catalog', 'Code', 'Link', 'NameVN', 'NameEN', 'Note', 'Active', 'AddedByAdmin'] as const;
type Col = (typeof COLS)[number];

/** `Unit` của các mục dùng chung cho mọi nhà ga (vd mã nhân lực, thiết bị): nhập một dòng, mọi nhà ga đều thấy. */
export const CATALOG_SHARED_UNIT = 'ALL';

/** Danh sách chọn dùng chung mọi nhà ga: giá trị người dùng thêm bằng "Khác…" ghi với `Unit = ALL` để nhà ga khác cũng thấy. */
export const SHARED_LIST_TABLES = ['weather', 'event_type', 'event_status', 'supplier', 'equipment_status', 'issue_severity', 'issue_status', 'subcontractor', 'material_id'];

const esc = (v: string): string => v.replace(/'/g, "''");

/** Mục có phải đã tồn tại không (SharePoint từ chối vì khoá duy nhất `Title` bị trùng). */
function isDuplicateError(err: unknown): boolean {
  const m = err instanceof Error ? err.message : '';
  return /duplicate|unique|already exists|\(409\)|-2130575169|-2147024809/i.test(m);
}

/**
 * Danh mục người dùng được tự thêm (SharePoint List `DRPCatalogItems`, dùng chung cho mọi nhà ga và mọi loại danh mục):
 * Element, Resource nhân lực/thiết bị, Vật tư… Mỗi dòng = một mục. Chống trùng hai lớp: kiểm tra trước khi thêm
 * và khoá duy nhất `Title` của list (`<Unit>|<Catalog>|code|<mã chuẩn hoá>` hoặc `…|link|<giá trị chuẩn hoá>`).
 */
export class ReportCatalogService {
  private _client: ISharePointClient;
  private _siteUrl: string;
  private _names: Promise<Record<Col, string>> | undefined;
  /** List đã có cột Active chưa (có thì dòng mới ghi Active = 1; thiếu cột thì không ghi, tránh lỗi trường không tồn tại). */
  private _hasActive = false;
  /** List đã có cột NameVN chưa (tuỳ chọn; chỉ lưu trong danh mục, thiếu cột thì bỏ qua). */
  private _hasNameVn = false;
  /** List đã có cột NameEN chưa (tuỳ chọn). */
  private _hasNameEn = false;
  /** List đã có cột Note chưa (tuỳ chọn, như NameEN). */
  private _hasNote = false;
  /** List đã có cột AddedByAdmin chưa (tuỳ chọn): có thì mục người dùng thêm bằng "Khác…" ghi AddedByAdmin = 0; thiếu cột thì bỏ qua. */
  private _hasAddedByAdmin = false;

  constructor(client: ISharePointClient, siteUrl: string) {
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
          this._hasActive = map.active !== undefined;
          this._hasNameVn = map.namevn !== undefined;
          this._hasNameEn = map.nameen !== undefined;
          this._hasNote = map.note !== undefined;
          this._hasAddedByAdmin = map.addedbyadmin !== undefined;
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
    const nameVn = str('NameVN');
    const nameEn = str('NameEN');
    const note = str('Note');
    const row: CatalogRow = { table: str('Catalog').toLowerCase(), code: str('Code'), link: str('Link') };
    // AddedByAdmin = 0 (hoặc false/no): người dùng thêm; trống hoặc 1: admin tạo sẵn
    const byUser = /^(0|false|no)$/i.test(str('AddedByAdmin'));
    return { ...row, ...(byUser ? { addedByUser: true } : {}), ...(nameVn ? { nameVn } : {}), ...(nameEn ? { nameEn } : {}), ...(note ? { note } : {}) };
  }

  /** Mọi mục của một nhà ga cộng các mục dùng chung (`Unit = ALL`). Lỗi (chưa tạo list, thiếu quyền…) thì ném ra để nơi gọi báo cảnh báo. */
  public async getRows(unit: string): Promise<CatalogRow[]> {
    const n = await this.names();
    const raw = await this._client.queryItems<Record<string, unknown>>(
      this._siteUrl,
      REPORT_CATALOG_LIST,
      `$filter=(${n.Unit} eq '${esc(unit)}' or ${n.Unit} eq '${CATALOG_SHARED_UNIT}')&$top=5000`
    );
    // Active = 0 (cột tuỳ chọn, trống/không có cột = hiện): ẩn khỏi dropdown
    return raw
      .filter((r) => parseActive(r[n.Active]))
      .map((r) => this.rowOf(n, r))
      .filter((r) => r.table && (r.code || r.link));
  }

  /**
   * Cập nhật giá trị liên kết (`Link`, vd Activity của cấu kiện) của mục có mã `code`. CHỈ mục do người dùng thêm
   * (`AddedByAdmin = 0`) mới được cập nhật; mục admin tạo sẵn (trống/1) bị từ chối (trả false). Người dùng cần quyền sửa dòng do chính họ tạo
   * (List settings > Advanced > Item-level Permissions: "Create items and edit items that were created by the user").
   */
  public async updateLink(unit: string, table: string, code: string, link: string): Promise<boolean> {
    const n = await this.names();
    if (!this._hasAddedByAdmin) return false; // không có cột đánh dấu thì không phân biệt được mục người dùng thêm
    const found = await this._client.queryItems<Record<string, unknown>>(
      this._siteUrl,
      REPORT_CATALOG_LIST,
      `$filter=${n.Catalog} eq '${esc(table.toLowerCase())}' and ${n.Code} eq '${esc(code)}' and (${n.Unit} eq '${esc(unit)}' or ${n.Unit} eq '${CATALOG_SHARED_UNIT}')&$top=1`
    );
    const raw = found[0];
    if (!raw || !this.rowOf(n, raw).addedByUser) return false;
    const id = Number(raw.Id ?? raw.ID);
    if (!id) return false;
    await this._client.updateItem<Record<string, unknown>>(this._siteUrl, REPORT_CATALOG_LIST, id, { [n.Link]: link });
    return true;
  }

  /**
   * Thêm một mục. Đã có mục cùng khoá (do người khác vừa thêm cùng lúc) thì không báo lỗi:
   * trả về mục sẵn có với `created = false`.
   */
  public async addRow(unit: string, row: CatalogRow): Promise<{ row: CatalogRow; created: boolean }> {
    if (SHARED_LIST_TABLES.indexOf(row.table.toLowerCase()) !== -1) unit = CATALOG_SHARED_UNIT;
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
        [n.Link]: row.link,
        ...(this._hasNameVn && row.nameVn ? { [n.NameVN]: row.nameVn } : {}),
        ...(this._hasNameEn && row.nameEn ? { [n.NameEN]: row.nameEn } : {}),
        ...(this._hasNote && row.note ? { [n.Note]: row.note } : {}),
        ...(this._hasAddedByAdmin ? { [n.AddedByAdmin]: 0 } : {}),
        ...(this._hasActive ? { [n.Active]: 1 } : {})
      });
      return { row: { ...row, addedByUser: true }, created: true };
    } catch (err: unknown) {
      if (!isDuplicateError(err)) throw err;
      const again = await find();
      return { row: again || row, created: false };
    }
  }
}
