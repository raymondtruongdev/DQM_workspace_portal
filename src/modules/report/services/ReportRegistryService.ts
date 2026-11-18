import { ISharePointClient } from '../../../core/services/SharePointClient';
import { Clock, LOCAL_CLOCK } from '../../../core/services/ServerClock';
import { ReportDayInfo } from '../types/report.types';
import { getDocShortPrefix, getReportFolderRelativePath } from '../utils/reportFileNaming';
import { formatReportNo } from '../utils/reportMeta';
import { normalizeDocId } from '../utils/reportRole';
import { DEFAULT_REPORT_SETTINGS, ReportSettings, isPastLocked, vnDateStr } from '../utils/reportSettings';

/** Tên SharePoint List quản lý trạng thái và số báo cáo (cột cần tạo: xem AGENTS.md mục 5.3). */
export const REPORT_REGISTRY_LIST = 'DRPRegistry';

export type RegistryStatus = 'unsigned' | 'done';

/** Báo cáo đang bị khoá: người dùng chỉ được xem, không sửa/lưu/upload/xoá. Admin mở khoá để cho sửa lại. */
export class ReportLockedError extends Error {
  constructor(message = 'Báo cáo này đã bị khoá. Vui lòng liên hệ admin để được mở khoá.') {
    super(message);
    this.name = 'ReportLockedError';
  }
}

/**
 * Dòng "giữ chỗ": admin khoá một ngày CHƯA nhập báo cáo nên sổ có dòng nhưng chưa có báo cáo (chưa lưu lần nào, số báo cáo 0).
 * Dòng này vẫn tính là "Chưa nhập"; khi báo cáo được lưu (sau khi admin mở khoá) dòng được điền tiếp như dòng mới.
 */
export function isPlaceholderRow(savedAt: string, reportNo: number, status: RegistryStatus): boolean {
  return !savedAt && reportNo === 0 && status === 'unsigned';
}

/** Trạng thái khoá của một báo cáo trong sổ. */
export interface IReportLock {
  /** Đã có dòng trong sổ (báo cáo chưa lưu lần nào thì chưa có dòng, không khoá được). */
  exists: boolean;
  locked: boolean;
  /** Báo cáo đã Lưu (có dòng thật trong sổ, không phải dòng giữ chỗ). Chỉ có khi `exists`. */
  saved?: boolean;
  /** Đã upload bản ký (`Status = done`). Chỉ có khi `exists`. */
  signed?: boolean;
  /** Đang mở khoá có hạn: thời điểm (ISO) báo cáo tự khoá lại; không có hạn hoặc đã hết hạn thì undefined. */
  unlockUntil?: string;
  /** List có cột `Locked` (chưa có thì không bật/tắt khoá được). */
  canToggle: boolean;
}

/**
 * Khoá hiệu lực của một dòng sổ: cột `Locked` = 1 → khoá, = 0 → mở (admin đã mở khoá); cột trống hoặc không có
 * cột → khoá khi đã upload bản ký (`Status = done`), nên báo cáo đã ký từ trước khi có tính năng này cũng được khoá,
 * hoặc khi `pastLocked` (ngày cũ trong chế độ "chỉ nhập ngày hiện tại"; admin đã mở khoá riêng với `Locked = 0` thì vẫn mở).
 */
export function isRegistryRowLocked(
  locked: boolean | undefined,
  status: RegistryStatus,
  pastLocked = false,
  unlockUntil = '',
  now: number = Date.now()
): boolean {
  // Mở khoá có hạn (`UnlockUntil`): quá hạn thì coi như không còn giá trị mở khoá riêng, quay về quy tắc mặc định bên dưới
  if (locked === false && isUnlockExpired(unlockUntil, now)) locked = undefined;
  if (locked !== undefined) return locked;
  return pastLocked || status === 'done';
}

/**
 * Báo cáo đang khoá vẫn cho NGƯỜI DÙNG upload bản đã ký khi báo cáo đã Lưu mà chưa có bản ký (không sửa được nội dung).
 * Đã có bản ký thì vẫn khoá hẳn (thay bản ký phải nhờ admin mở khoá).
 */
export function canUploadSignedWhileLocked(lock: Pick<IReportLock, 'saved' | 'signed'> | undefined): boolean {
  return Boolean(lock && lock.saved && !lock.signed);
}

/** Thời hạn mở khoá (ISO) đã qua. Trống hoặc không đọc được = không có hạn. */
export function isUnlockExpired(unlockUntil: string, now: number = Date.now()): boolean {
  if (!unlockUntil) return false;
  const t = Date.parse(unlockUntil);
  return Number.isFinite(t) && t <= now;
}

/** Hạn mở khoá còn hiệu lực của dòng sổ (đang mở khoá riêng và chưa quá hạn), không thì undefined. */
function activeUnlockUntil(item: { Locked: boolean | undefined; UnlockUntil: string }, now: number = Date.now()): string | undefined {
  return item.Locked === false && item.UnlockUntil && !isUnlockExpired(item.UnlockUntil, now) ? item.UnlockUntil : undefined;
}

/** Dấu vết bản đã ký: ai upload bản hiện tại, ai xoá và lúc nào (lần xoá gần nhất). */
export interface ISignedTrail {
  uploadBy: string;
  removedAt: string;
  removedBy: string;
}

/** Tên (tiêu đề) các cột của list; tên nội bộ thật được tra từ list vì có thể khác (vd Unit -> Unit0). */
const COLS = ['No', 'Title', 'Unit', 'DocID', 'ReportDate', 'ReportNo', 'Reporter', 'Status', 'SavedAt', 'SignedAt', 'Counters', 'SignedUploadBy', 'SignedRemovedAt', 'SignedRemovedBy', 'Locked', 'UnlockUntil'] as const;
type Col = (typeof COLS)[number];

/** Một dòng của list (khoá = tên cột chuẩn ở trên, đã đổi từ tên nội bộ). */
interface IRegistryRow {
  id: number;
  /** Số thứ tự dòng (cột tuỳ chọn, code tự đánh khi tạo dòng mới). */
  No: number;
  Title: string;
  Unit: string;
  DocID: string;
  ReportDate: string;
  ReportNo: number;
  Reporter: string;
  Status: RegistryStatus;
  SavedAt: string;
  SignedAt: string;
  /** JSON bộ đếm mã tự tăng tại thời điểm lưu, vd {"events.eventid":5,"issues.issueid":4}. */
  Counters: string;
  /** Email người upload bản đã ký hiện tại (cột không bắt buộc; trống khi chưa có bản ký). */
  SignedUploadBy: string;
  /** Lần xoá bản đã ký gần nhất: thời điểm ISO và người xoá (cột không bắt buộc). */
  SignedRemovedAt: string;
  SignedRemovedBy: string;
  /** Cột `Locked` (Number 1/0, tuỳ chọn): 1 = khoá, 0 = admin đã mở khoá, trống = theo trạng thái (xem isRegistryRowLocked). */
  Locked: boolean | undefined;
  /** Cột `UnlockUntil` (Text ISO, tuỳ chọn): admin mở khoá có hạn, quá giờ này báo cáo tự khoá lại (xem isRegistryRowLocked). */
  UnlockUntil: string;
}

const esc = (v: string): string => v.replace(/'/g, "''");

/**
 * Sổ quản lý báo cáo (SharePoint List `DRPRegistry`): số báo cáo hiện tại của từng đơn vị, người lập,
 * trạng thái, thời điểm lưu/ký. Nguồn dữ liệu của số báo cáo, lịch 14 ngày và tab Tổng quan.
 */
export class ReportRegistryService {
  private _client: ISharePointClient;
  private _siteUrl: string;
  private _libraryPath: () => string;
  private _names: Promise<Record<Col, string>> | undefined;
  /** List đã có cột Counters chưa (chưa thêm thì bỏ qua bộ đếm, mã tự tăng chỉ dựa vào báo cáo đang mở). */
  private _hasCounters = false;
  /** List đã có cột SignedRemovedAt/SignedRemovedBy chưa (chưa thêm thì không ghi/đọc lần xoá bản ký). */
  private _hasRemoval = false;
  /** List đã có cột SignedUploadBy chưa. */
  private _hasUploadBy = false;
  /** List đã có cột No chưa (có thì tự đánh số thứ tự khi tạo dòng mới). */
  private _hasNo = false;
  /** List đã có cột Locked chưa (chưa có thì không bật/tắt khoá được; báo cáo đã ký vẫn coi là khoá). */
  private _hasLocked = false;
  /** List có cột UnlockUntil chưa (chưa có thì mở khoá không có hạn). */
  private _hasUnlockUntil = false;

  /** Đọc cài đặt báo cáo của admin (chế độ chỉ nhập ngày hiện tại); lỗi hoặc không truyền = tắt. */
  private readonly _settings: () => Promise<ReportSettings>;
  /** Giờ máy chủ SharePoint (không tin giờ máy người dùng) cho ngày hôm nay và hạn mở khoá. */
  private readonly _clock: Clock;

  constructor(
    client: ISharePointClient,
    siteUrl: string,
    libraryPath: () => string,
    settings: () => Promise<ReportSettings> = () => Promise.resolve(DEFAULT_REPORT_SETTINGS),
    clock: Clock = LOCAL_CLOCK
  ) {
    this._settings = settings;
    this._clock = clock;
    this._client = client;
    this._siteUrl = siteUrl;
    this._libraryPath = libraryPath;
  }

  /** Khoá duy nhất của một báo cáo (cột Title): "DRP_20261112_ST01" = <DRP|WRP>_<yyyyMMdd>_<Unit>. */
  public static keyOf(unit: string, docId: string, dateStr: string): string {
    return `${getDocShortPrefix(docId)}_${dateStr.replace(/-/g, '')}_${unit}`;
  }

  /** Tên nội bộ của các cột (đọc một lần, nhớ lại); không tra được thì coi tên nội bộ = tên cột. */
  private names(): Promise<Record<Col, string>> {
    if (!this._names) {
      this._names = this._client
        .getListFieldMap(this._siteUrl, REPORT_REGISTRY_LIST)
        .then((map) => {
          const out = {} as Record<Col, string>;
          this._hasCounters = map.counters !== undefined;
          this._hasUploadBy = map.signeduploadby !== undefined;
          this._hasNo = map.no !== undefined;
          this._hasLocked = map.locked !== undefined;
          this._hasUnlockUntil = map.unlockuntil !== undefined;
          this._hasRemoval = map.signedremovedat !== undefined && map.signedremovedby !== undefined;
          COLS.forEach((c) => {
            out[c] = map[c.toLowerCase()] || c;
          });
          out.Title = 'Title'; // cột chính luôn tên nội bộ Title (tránh nhầm với cột hiển thị LinkTitle…)
          return out;
        })
        .catch((err: unknown) => {
          this._names = undefined; // lần sau thử lại
          throw err;
        });
    }
    return this._names;
  }

  /** Đổi dữ liệu theo tên cột chuẩn sang tên nội bộ để ghi vào list. */
  private toInternal(n: Record<Col, string>, data: Partial<Omit<IRegistryRow, 'id'>>): Record<string, unknown> {
    const out: Record<string, unknown> = {};
    (Object.keys(data) as Col[]).forEach((k) => {
      out[n[k]] = k === 'Locked' ? (data.Locked ? 1 : 0) : data[k]; // cột Locked là Number 1/0
    });
    return out;
  }

  private fromInternal(n: Record<Col, string>, raw: Record<string, unknown>): IRegistryRow {
    const str = (c: Col): string => (raw[n[c]] === undefined || raw[n[c]] === null ? '' : String(raw[n[c]]));
    return {
      id: Number(raw.Id),
      No: Number(raw[n.No]) || 0,
      Title: str('Title'),
      Unit: str('Unit'),
      DocID: normalizeDocId(str('DocID')),
      ReportDate: str('ReportDate'),
      ReportNo: Number(raw[n.ReportNo]) || 0,
      Reporter: str('Reporter'),
      Status: str('Status') === 'done' ? 'done' : 'unsigned',
      SavedAt: str('SavedAt'),
      SignedAt: str('SignedAt'),
      Counters: str('Counters'),
      SignedUploadBy: str('SignedUploadBy'),
      SignedRemovedAt: str('SignedRemovedAt'),
      SignedRemovedBy: str('SignedRemovedBy'),
      Locked: str('Locked') === '' ? undefined : /^(1|true|yes)$/i.test(str('Locked')),
      UnlockUntil: str('UnlockUntil')
    };
  }

  private async query(query: (n: Record<Col, string>) => string): Promise<IRegistryRow[]> {
    const n = await this.names();
    const raw = await this._client.queryItems<Record<string, unknown>>(this._siteUrl, REPORT_REGISTRY_LIST, query(n));
    return raw.map((r) => this.fromInternal(n, r));
  }

  private async findItem(unit: string, docId: string, dateStr: string): Promise<IRegistryRow | undefined> {
    const key = ReportRegistryService.keyOf(unit, docId, dateStr);
    const rows = await this.query((n) => `$filter=${n.Title} eq '${esc(key)}'&$top=1`);
    return rows[0];
  }

  /**
   * Số báo cáo cho một ngày: nếu ngày đó đã có dòng trong sổ thì dùng lại số đó (vd file báo cáo bị xoá),
   * nếu chưa thì số lớn nhất của đơn vị + 1 (đơn vị chưa có dòng nào = 00001). Lỗi đọc sổ thì ném ra.
   */
  public async getReportNoFor(unit: string, docId: string, dateStr: string): Promise<string> {
    const existing = await this.findItem(unit, docId, dateStr);
    if (existing && existing.ReportNo > 0) return formatReportNo(existing.ReportNo);
    const top = await this.query(
      (n) =>
        `$filter=${n.Unit} eq '${esc(unit)}' and ${n.DocID} eq '${esc(getDocShortPrefix(docId))}'&$orderby=${n.ReportNo} desc&$top=1`
    );
    return formatReportNo((top[0] ? top[0].ReportNo : 0) + 1);
  }

  /**
   * Bộ đếm mã tự tăng (Event ID, Issue ID…) lớn nhất của đơn vị: lấy từ dòng có số báo cáo cao nhất (mỗi lần lưu ghi
   * bộ đếm cộng dồn nên dòng mới nhất luôn lớn nhất). Chưa có cột Counters hoặc chưa có dòng nào thì trả rỗng.
   */
  public async getCounters(unit: string, docId: string): Promise<Record<string, number>> {
    await this.names();
    if (!this._hasCounters) return {};
    const top = await this.query(
      (n) => `$filter=${n.Unit} eq '${esc(unit)}' and ${n.DocID} eq '${esc(getDocShortPrefix(docId))}'&$orderby=${n.ReportNo} desc&$top=1`
    );
    return top[0] ? ReportRegistryService.parseCounters(top[0].Counters) : {};
  }

  /** JSON -> bộ đếm; chuỗi hỏng hoặc rỗng thì trả rỗng. */
  public static parseCounters(raw: string): Record<string, number> {
    try {
      const obj = raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
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

  /** Ghi nhận báo cáo vừa lưu: tạo dòng mới (báo cáo mới) hoặc cập nhật người lập + thời điểm lưu (+ bộ đếm mã tự tăng). */
  public async recordSave(
    unit: string,
    docId: string,
    dateStr: string,
    reportNo: string,
    reporter: string,
    counters?: Record<string, number>
  ): Promise<void> {
    const n = await this.names();
    const now = new Date().toISOString();
    const extra: Partial<Omit<IRegistryRow, 'id'>> =
      this._hasCounters && counters ? { Counters: JSON.stringify(counters) } : {};
    const existing = await this.findItem(unit, docId, dateStr);
    if (existing) {
      await this._client.updateItem<Record<string, unknown>>(
        this._siteUrl,
        REPORT_REGISTRY_LIST,
        existing.id,
        this.toInternal(n, {
          Reporter: reporter,
          SavedAt: now,
          ...(existing.ReportNo === 0 ? { ReportNo: parseInt(reportNo, 10) || 0 } : {}), // dòng giữ chỗ do admin khoá trước
          ...extra
        })
      );
      return;
    }
    let no: Partial<Omit<IRegistryRow, 'id'>> = {};
    if (this._hasNo) {
      const last = await this.query((c) => `$orderby=${c.No} desc&$top=1`);
      no = { No: (last[0] ? last[0].No : 0) + 1 };
    }
    await this._client.createItem<Record<string, unknown>>(
      this._siteUrl,
      REPORT_REGISTRY_LIST,
      this.toInternal(n, {
        ...no,
        Title: ReportRegistryService.keyOf(unit, docId, dateStr),
        Unit: unit,
        DocID: getDocShortPrefix(docId),
        ReportDate: dateStr,
        ReportNo: parseInt(reportNo, 10) || 0,
        Reporter: reporter,
        Status: 'unsigned',
        SavedAt: now,
        ...extra
      })
    );
  }

  /** Ghi nhận đã upload bản đã ký: trạng thái Hoàn tất + thời điểm ký + email người upload (`uploader`). */
  public async recordSigned(
    unit: string,
    docId: string,
    dateStr: string,
    reportNo: string,
    reporter: string,
    uploader = ''
  ): Promise<void> {
    const n = await this.names();
    let item = await this.findItem(unit, docId, dateStr);
    if (!item) {
      await this.recordSave(unit, docId, dateStr, reportNo, reporter);
      item = await this.findItem(unit, docId, dateStr);
    }
    if (!item) return;
    await this._client.updateItem<Record<string, unknown>>(
      this._siteUrl,
      REPORT_REGISTRY_LIST,
      item.id,
      this.toInternal(n, {
        Status: 'done',
        SignedAt: new Date().toISOString(),
        ...(this._hasUploadBy ? { SignedUploadBy: uploader || reporter } : {}),
        ...(this._hasLocked ? { Locked: true } : {}),
        ...(this._hasUnlockUntil ? { UnlockUntil: '' } : {})
      })
    );
  }

  /**
   * Ghi nhận đã xoá bản đã ký: trạng thái về "Chưa upload bản ký", xoá thời điểm ký, và (nếu list có cột
   * SignedRemovedAt/SignedRemovedBy) ghi thời điểm + người xoá GẦN NHẤT. Trả về true nếu đã ghi được người xoá.
   */
  public async recordUnsigned(unit: string, docId: string, dateStr: string, actor: string): Promise<boolean> {
    const n = await this.names();
    const item = await this.findItem(unit, docId, dateStr);
    if (!item) return false;
    const removal: Partial<Omit<IRegistryRow, 'id'>> = this._hasRemoval
      ? { SignedRemovedAt: new Date().toISOString(), SignedRemovedBy: actor }
      : {};
    await this._client.updateItem<Record<string, unknown>>(
      this._siteUrl,
      REPORT_REGISTRY_LIST,
      item.id,
      this.toInternal(n, { Status: 'unsigned', SignedAt: '', ...(this._hasUploadBy ? { SignedUploadBy: '' } : {}), ...(this._hasLocked ? { Locked: false } : {}), ...(this._hasUnlockUntil ? { UnlockUntil: '' } : {}), ...removal })
    );
    return this._hasRemoval;
  }

  /** Trạng thái khoá của báo cáo một ngày (chưa có dòng trong sổ = chưa khoá). Lỗi đọc sổ thì ném ra. */
  public async getLock(unit: string, docId: string, dateStr: string): Promise<IReportLock> {
    await this.names();
    await this._clock.ensureFresh();
    const now = this._clock.now();
    const past = isPastLocked(await this._settings(), dateStr, vnDateStr(now));
    const item = await this.findItem(unit, docId, dateStr);
    const until = item ? activeUnlockUntil(item, now) : undefined;
    return {
      exists: Boolean(item),
      ...(item ? { saved: !isPlaceholderRow(item.SavedAt, item.ReportNo, item.Status), signed: item.Status === 'done' } : {}),
      locked: item ? isRegistryRowLocked(item.Locked, item.Status, past, item.UnlockUntil, now) : past,
      ...(until ? { unlockUntil: until } : {}),
      canToggle: this._hasLocked
    };
  }

  /** Admin bật/tắt khoá một báo cáo; ngày chưa nhập thì tạo dòng giữ chỗ để khoá trước (xem `isPlaceholderRow`). Ném lỗi nếu list chưa có cột `Locked`. */
  public async setLocked(unit: string, docId: string, dateStr: string, locked: boolean, unlockHours = 0): Promise<void> {
    const n = await this.names();
    if (!this._hasLocked) throw new Error('List DRPRegistry chưa có cột Locked (Number 1/0) nên chưa bật/tắt khoá được.');
    await this._clock.ensureFresh();
    const now = this._clock.now();
    const item = await this.findItem(unit, docId, dateStr);
    // Mở khoá có hạn: ghi giờ tự khoá lại (cần cột UnlockUntil); khoá lại thì xoá hạn cũ
    const expiry: Partial<Omit<IRegistryRow, 'id'>> = this._hasUnlockUntil
      ? { UnlockUntil: !locked && unlockHours > 0 ? new Date(now + unlockHours * 3600000).toISOString() : '' }
      : {};
    if (item) {
      await this._client.updateItem<Record<string, unknown>>(this._siteUrl, REPORT_REGISTRY_LIST, item.id, this.toInternal(n, { Locked: locked, ...expiry }));
      return;
    }
    // Chưa có dòng nghĩa là chưa khoá: không cần mở, trừ ngày cũ đang bị khoá theo chế độ "chỉ nhập ngày hiện tại"
    // (cần dòng giữ chỗ `Locked = 0` để ghi nhận admin đã mở khoá riêng ngày đó)
    if (!locked && !isPastLocked(await this._settings(), dateStr, vnDateStr(now))) return;
    // Ngày chưa nhập báo cáo: tạo dòng giữ chỗ để khoá trước (người dùng chỉ xem, chưa lưu được)
    let no: Partial<Omit<IRegistryRow, 'id'>> = {};
    if (this._hasNo) {
      const last = await this.query((c) => `$orderby=${c.No} desc&$top=1`);
      no = { No: (last[0] ? last[0].No : 0) + 1 };
    }
    await this._client.createItem<Record<string, unknown>>(
      this._siteUrl,
      REPORT_REGISTRY_LIST,
      this.toInternal(n, {
        ...no,
        Title: ReportRegistryService.keyOf(unit, docId, dateStr),
        Unit: unit,
        DocID: getDocShortPrefix(docId),
        ReportDate: dateStr,
        ReportNo: 0,
        Status: 'unsigned',
        Locked: locked,
        ...expiry
      })
    );
  }

  /** Dấu vết bản đã ký của một ngày: người upload hiện tại và lần xoá gần nhất (cột chưa có thì trường tương ứng trống). */
  public async getSignedTrail(unit: string, docId: string, dateStr: string): Promise<ISignedTrail | undefined> {
    await this.names();
    if (!this._hasRemoval && !this._hasUploadBy) return undefined;
    const item = await this.findItem(unit, docId, dateStr);
    if (!item) return undefined;
    return {
      uploadBy: item.SignedUploadBy,
      removedAt: item.SignedRemovedAt,
      removedBy: item.SignedRemovedBy
    };
  }

  /**
   * Trạng thái các ngày từ `fromDate` (YYYY-MM-DD) của mọi đơn vị (hoặc một đơn vị), trong một lần đọc.
   * Kết quả: khoá "<unit>|<docId>" (chữ thường) -> ngày -> thông tin. Ngày không có dòng = chưa nhập.
   */
  public async getDayInfosSince(
    docId: string,
    fromDate: string,
    unit?: string
  ): Promise<Record<string, Record<string, ReportDayInfo>>> {
    const settings = await this._settings();
    await this._clock.ensureFresh();
    const now = this._clock.now();
    const today = vnDateStr(now);
    const rows = await this.query(
      (n) =>
        `$filter=${n.DocID} eq '${esc(getDocShortPrefix(docId))}' and ${n.ReportDate} ge '${esc(fromDate)}'` +
        `${unit ? ` and ${n.Unit} eq '${esc(unit)}'` : ''}&$top=5000`
    );
    const out: Record<string, Record<string, ReportDayInfo>> = {};
    rows.forEach((it) => {
      const key = `${it.Unit}|${it.DocID}`.toLowerCase();
      const folder = getReportFolderRelativePath(this._libraryPath(), it.Unit, it.DocID, it.ReportDate);
      const done = it.Status === 'done';
      const placeholder = isPlaceholderRow(it.SavedAt, it.ReportNo, it.Status);
      const info: ReportDayInfo = {
        dateStr: it.ReportDate,
        status: done ? 'done' : placeholder ? 'missing' : 'unsigned',
        savedAt: it.SavedAt || undefined,
        reporter: it.Reporter || undefined,
        reportNo: it.ReportNo > 0 ? formatReportNo(it.ReportNo) : undefined,
        locked: isRegistryRowLocked(it.Locked, it.Status, isPastLocked(settings, it.ReportDate, today), it.UnlockUntil, now),
        unlockUntil: activeUnlockUntil(it, now),
        signed: done
          ? {
              // Tên file = <DocCode>_signed.pdf: sổ không lưu Doc. code nên chỉ biết khi mở (openSignedDocument)
              name: '',
              serverRelativeUrl: folder,
              length: 0,
              timeLastModified: it.SignedAt || it.SavedAt
            }
          : undefined
      };
      if (!out[key]) out[key] = {};
      out[key][it.ReportDate] = info;
    });
    return out;
  }
}
