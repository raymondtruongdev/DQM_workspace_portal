import { ISharePointClient } from '../../../core/services';
import { getDocShortPrefix } from '../utils/reportFileNaming';

/** Tên SharePoint List bộ đếm mã tự tăng (một dòng mỗi đơn vị × loại báo cáo; cột: xem docs/SHAREPOINT_SETUP.md). */
export const REPORT_COUNTERS_LIST = 'DRPCounters';

const MAX_ATTEMPTS = 5;

const esc = (v: string): string => v.replace(/'/g, "''");

/** JSON -> bộ đếm; chuỗi hỏng hoặc rỗng thì trả rỗng. */
export function parseCounters(raw: unknown): Record<string, number> {
  try {
    const obj = raw ? (JSON.parse(String(raw)) as Record<string, unknown>) : {};
    const out: Record<string, number> = {};
    Object.keys(obj).forEach((k) => {
      const v = Number(obj[k]);
      if (isFinite(v) && v > 0) out[k] = Math.floor(v);
    });
    return out;
  } catch {
    return {};
  }
}

/** Gộp hai bộ đếm: lấy số lớn hơn theo từng khoá (bộ đếm chỉ tăng, không bao giờ giảm). */
export function mergeCounters(a: Record<string, number>, b: Record<string, number>): Record<string, number> {
  const out: Record<string, number> = { ...a };
  Object.keys(b).forEach((k) => {
    out[k] = Math.max(out[k] || 0, b[k]);
  });
  return out;
}

const sameCounters = (a: Record<string, number>, b: Record<string, number>): boolean => {
  const keys = Object.keys(a);
  return keys.length === Object.keys(b).length && keys.every((k) => a[k] === b[k]);
};

/**
 * Bộ đếm mã tự tăng (Event ID, Issue ID…) của từng đơn vị, lưu ở list `DRPCounters`: một dòng mỗi `<Unit>|<DRP|WRP>`
 * (cột `Title`, duy nhất), cột `Counters` là JSON `{"events.eventid":8,"issues.issueid":4}`. Đây là nguồn DUY NHẤT để tính
 * mã kế tiếp, nên không phụ thuộc thứ tự ngày nhập báo cáo; sổ DRPRegistry không còn dùng để tính.
 *
 * Người báo cáo cần quyền thêm/sửa list này (bộ đếm chỉ làm lệch số mã, không ảnh hưởng số báo cáo hay cấu hình).
 * Ghi dùng ETag (`If-Match`): hai người lưu cùng lúc thì người sau nhận 412, đọc lại, gộp bằng max rồi ghi lại.
 * Không nhớ đệm: mỗi lần mở/lưu báo cáo đọc lại để không dùng số cũ.
 */
export class ReportCounterService {
  private _client: ISharePointClient;
  private _siteUrl: string;
  private _names: Promise<{ title: string; counters: string }> | undefined;

  constructor(client: ISharePointClient, siteUrl: string) {
    this._client = client;
    this._siteUrl = siteUrl;
  }

  /** Khoá dòng: "ST01|DRP". */
  public static keyOf(unit: string, docId: string): string {
    return `${unit}|${getDocShortPrefix(docId)}`;
  }

  private names(): Promise<{ title: string; counters: string }> {
    if (!this._names) {
      this._names = this._client
        .getListFieldMap(this._siteUrl, REPORT_COUNTERS_LIST)
        .then((map) => ({ title: 'Title', counters: map.counters || 'Counters' }))
        .catch((err: unknown) => {
          this._names = undefined; // lần sau thử lại
          throw err;
        });
    }
    return this._names;
  }

  private async read(unit: string, docId: string): Promise<{ id?: number; eTag: string; counters: Record<string, number> }> {
    const n = await this.names();
    const rows = await this._client.queryItemsWithETag<Record<string, unknown>>(
      this._siteUrl,
      REPORT_COUNTERS_LIST,
      `$filter=${n.title} eq '${esc(ReportCounterService.keyOf(unit, docId))}'&$top=1`
    );
    if (!rows[0]) return { eTag: '', counters: {} };
    return { id: Number(rows[0].item.Id), eTag: rows[0].eTag, counters: parseCounters(rows[0].item[n.counters]) };
  }

  /** Bộ đếm hiện tại của đơn vị; chưa có dòng thì rỗng. Ném lỗi nếu không đọc được list. */
  public async get(unit: string, docId: string): Promise<Record<string, number>> {
    return (await this.read(unit, docId)).counters;
  }

  /**
   * ĐẶT CHỖ một số mã tự tăng (khi người dùng bấm thêm hàng): tăng bộ đếm `key` lên 1 (không thấp hơn `atLeast` + 1) và ghi
   * ngay vào list, trả về số vừa đặt. Hai người nhập hai báo cáo khác nhau của cùng đơn vị vì vậy không bao giờ nhận cùng một số.
   * Hàng bị xoá hoặc báo cáo bỏ dở thì số đã đặt bị bỏ qua (dãy mã có thể có chỗ trống, không bao giờ trùng).
   * Xung đột (412 / trùng khoá tạo dòng đầu) thì đọc lại và thử tiếp; ném lỗi nếu list không đọc/ghi được (nơi gọi tự dự phòng).
   */
  public async reserve(unit: string, docId: string, key: string, atLeast = 0): Promise<number> {
    const n = await this.names();
    let lastError: unknown;
    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
      const current = await this.read(unit, docId);
      const next = Math.max(current.counters[key] || 0, atLeast) + 1;
      const merged = { ...current.counters, [key]: next };
      try {
        if (current.id === undefined) {
          await this._client.createItem<Record<string, unknown>>(this._siteUrl, REPORT_COUNTERS_LIST, {
            [n.title]: ReportCounterService.keyOf(unit, docId),
            [n.counters]: JSON.stringify(merged)
          });
        } else {
          await this._client.updateItem<Record<string, unknown>>(
            this._siteUrl,
            REPORT_COUNTERS_LIST,
            current.id,
            { [n.counters]: JSON.stringify(merged) },
            current.eTag
          );
        }
        return next;
      } catch (err: unknown) {
        lastError = err;
        const msg = err instanceof Error ? err.message : String(err);
        if (!/412|409|duplicate|unique|precondition/i.test(msg)) throw err;
      }
    }
    throw lastError instanceof Error ? lastError : new Error('Không đặt chỗ được mã tự tăng (xung đột liên tục).');
  }

  /**
   * Nâng bộ đếm của đơn vị lên ít nhất bằng `counters` (gộp theo max với giá trị đang lưu, không bao giờ giảm).
   * Tạo dòng nếu chưa có. Xung đột (412 hoặc trùng khoá khi hai người cùng tạo dòng đầu tiên) thì đọc lại và thử tiếp.
   */
  public async raise(unit: string, docId: string, counters: Record<string, number>): Promise<void> {
    const n = await this.names();
    let lastError: unknown;
    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
      const current = await this.read(unit, docId);
      const merged = mergeCounters(current.counters, counters);
      if (current.id !== undefined && sameCounters(merged, current.counters)) return; // không có gì cần nâng
      try {
        if (current.id === undefined) {
          await this._client.createItem<Record<string, unknown>>(this._siteUrl, REPORT_COUNTERS_LIST, {
            [n.title]: ReportCounterService.keyOf(unit, docId),
            [n.counters]: JSON.stringify(merged)
          });
        } else {
          await this._client.updateItem<Record<string, unknown>>(
            this._siteUrl,
            REPORT_COUNTERS_LIST,
            current.id,
            { [n.counters]: JSON.stringify(merged) },
            current.eTag
          );
        }
        return;
      } catch (err: unknown) {
        lastError = err;
        const msg = err instanceof Error ? err.message : String(err);
        if (!/412|409|duplicate|unique|precondition/i.test(msg)) throw err; // lỗi khác (quyền, mạng…): báo ngay
      }
    }
    throw lastError instanceof Error ? lastError : new Error('Không ghi được bộ đếm mã tự tăng (xung đột liên tục).');
  }
}
