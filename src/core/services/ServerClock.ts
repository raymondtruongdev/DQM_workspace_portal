import type { ISharePointClient } from './SharePointClient';

/** Nguồn thời gian: giờ máy chủ SharePoint (đồng bộ qua `ServerClock`) hoặc giờ máy người dùng (`LOCAL_CLOCK`). */
export interface Clock {
  /** Thời điểm hiện tại (ms kể từ 1970, UTC). */
  now(): number;
  /** Đồng bộ lại với máy chủ nếu lần đồng bộ cuối đã cũ hơn `maxAgeMs` (mặc định 60 giây). Không bao giờ ném lỗi. */
  ensureFresh(maxAgeMs?: number): Promise<void>;
}

/** Giờ máy người dùng (dùng khi không có máy chủ, vd test). */
export const LOCAL_CLOCK: Clock = {
  now: () => Date.now(),
  ensureFresh: () => Promise.resolve()
};

/**
 * Đồng hồ theo GIỜ MÁY CHỦ SharePoint (header `Date` của một phản hồi REST), để người dùng chỉnh giờ máy
 * không lách được các kiểm tra theo ngày/giờ (khoá báo cáo ngày cũ, hạn mở khoá).
 * Sau mỗi lần đồng bộ, giờ chạy tiếp bằng `performance.now()` (đồng hồ đơn điệu, KHÔNG đổi khi người dùng chỉnh giờ hệ thống).
 * Không đồng bộ được (mạng, thiếu header) thì dùng tạm giờ máy người dùng. Vẫn chỉ là kiểm tra ở trình duyệt, không phải bảo mật thật.
 */
export class ServerClock implements Clock {
  private _serverMs = 0;
  private _perfAtSync = 0;
  private _synced = false;
  private _pending: Promise<void> | undefined;

  constructor(private readonly _client: Pick<ISharePointClient, 'getServerTime'>, private readonly _siteUrl: string) {}

  public now(): number {
    return this._synced ? this._serverMs + (performance.now() - this._perfAtSync) : Date.now();
  }

  /** Đã đồng bộ được với máy chủ ít nhất một lần. */
  public get isSynced(): boolean {
    return this._synced;
  }

  public ensureFresh(maxAgeMs = 60000): Promise<void> {
    if (this._synced && performance.now() - this._perfAtSync < maxAgeMs) return Promise.resolve();
    if (!this._pending) {
      const request = this.sync().then(
        () => {
          this._pending = undefined;
        },
        () => {
          this._pending = undefined;
        }
      );
      this._pending = request;
    }
    return this._pending;
  }

  /** Hỏi giờ máy chủ; lỗi thì giữ lần đồng bộ trước (hoặc giờ máy). */
  public async sync(): Promise<void> {
    if (typeof this._client.getServerTime !== 'function') return;
    try {
      const before = performance.now();
      const serverMs = await this._client.getServerTime(this._siteUrl);
      const after = performance.now();
      if (!Number.isFinite(serverMs)) return;
      this._serverMs = serverMs;
      this._perfAtSync = (before + after) / 2; // header Date do máy chủ đặt khoảng giữa chuyến đi
      this._synced = true;
    } catch (err: unknown) {
      console.warn('[ServerClock] Could not read server time, using local clock:', err);
    }
  }
}
