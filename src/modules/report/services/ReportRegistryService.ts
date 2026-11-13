import { SharePointClient } from '../../../core/services/SharePointClient';
import { ReportDayInfo } from '../types/report.types';
import { getReportFolderRelativePath } from '../utils/reportFileNaming';
import { formatReportNo } from '../utils/reportMeta';

/** Tên SharePoint List quản lý trạng thái và số báo cáo (cột cần tạo: xem AGENTS.md mục 5.3). */
export const REPORT_REGISTRY_LIST = 'ReportRegistry';

export type RegistryStatus = 'unsigned' | 'done';

/** Tên (tiêu đề) các cột của list; tên nội bộ thật được tra từ list vì có thể khác (vd Unit -> Unit0). */
const COLS = ['Title', 'Unit', 'DocID', 'ReportDate', 'ReportNo', 'Reporter', 'Status', 'SavedAt', 'SignedAt', 'Counters'] as const;
type Col = (typeof COLS)[number];

/** Một dòng của list (khoá = tên cột chuẩn ở trên, đã đổi từ tên nội bộ). */
interface IRegistryRow {
  id: number;
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
}

const esc = (v: string): string => v.replace(/'/g, "''");

/**
 * Sổ quản lý báo cáo (SharePoint List `ReportRegistry`): số báo cáo hiện tại của từng đơn vị, người lập,
 * trạng thái, thời điểm lưu/ký. Nguồn dữ liệu của số báo cáo, lịch 14 ngày và tab Tổng quan.
 */
export class ReportRegistryService {
  private _client: SharePointClient;
  private _siteUrl: string;
  private _libraryPath: () => string;
  private _names: Promise<Record<Col, string>> | undefined;
  /** List đã có cột Counters chưa (chưa thêm thì bỏ qua bộ đếm, mã tự tăng chỉ dựa vào báo cáo đang mở). */
  private _hasCounters = false;

  constructor(client: SharePointClient, siteUrl: string, libraryPath: () => string) {
    this._client = client;
    this._siteUrl = siteUrl;
    this._libraryPath = libraryPath;
  }

  /** Khoá duy nhất của một báo cáo: "ST01|DailyReport|2026-11-12". */
  public static keyOf(unit: string, docId: string, dateStr: string): string {
    return `${unit}|${docId}|${dateStr}`;
  }

  /** Tên nội bộ của các cột (đọc một lần, nhớ lại); không tra được thì coi tên nội bộ = tên cột. */
  private names(): Promise<Record<Col, string>> {
    if (!this._names) {
      this._names = this._client
        .getListFieldMap(this._siteUrl, REPORT_REGISTRY_LIST)
        .then((map) => {
          const out = {} as Record<Col, string>;
          this._hasCounters = map.counters !== undefined;
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
      out[n[k]] = data[k];
    });
    return out;
  }

  private fromInternal(n: Record<Col, string>, raw: Record<string, unknown>): IRegistryRow {
    const str = (c: Col): string => (raw[n[c]] === undefined || raw[n[c]] === null ? '' : String(raw[n[c]]));
    return {
      id: Number(raw.Id),
      Title: str('Title'),
      Unit: str('Unit'),
      DocID: str('DocID'),
      ReportDate: str('ReportDate'),
      ReportNo: Number(raw[n.ReportNo]) || 0,
      Reporter: str('Reporter'),
      Status: str('Status') === 'done' ? 'done' : 'unsigned',
      SavedAt: str('SavedAt'),
      SignedAt: str('SignedAt'),
      Counters: str('Counters')
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
        `$filter=${n.Unit} eq '${esc(unit)}' and ${n.DocID} eq '${esc(docId)}'&$orderby=${n.ReportNo} desc&$top=1`
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
      (n) => `$filter=${n.Unit} eq '${esc(unit)}' and ${n.DocID} eq '${esc(docId)}'&$orderby=${n.ReportNo} desc&$top=1`
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
        this.toInternal(n, { Reporter: reporter, SavedAt: now, ...extra })
      );
      return;
    }
    await this._client.createItem<Record<string, unknown>>(
      this._siteUrl,
      REPORT_REGISTRY_LIST,
      this.toInternal(n, {
        Title: ReportRegistryService.keyOf(unit, docId, dateStr),
        Unit: unit,
        DocID: docId,
        ReportDate: dateStr,
        ReportNo: parseInt(reportNo, 10) || 0,
        Reporter: reporter,
        Status: 'unsigned',
        SavedAt: now,
        ...extra
      })
    );
  }

  /** Ghi nhận đã upload bản đã ký: trạng thái Hoàn tất + thời điểm ký. */
  public async recordSigned(
    unit: string,
    docId: string,
    dateStr: string,
    reportNo: string,
    reporter: string
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
      this.toInternal(n, { Status: 'done', SignedAt: new Date().toISOString() })
    );
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
    const rows = await this.query(
      (n) =>
        `$filter=${n.DocID} eq '${esc(docId)}' and ${n.ReportDate} ge '${esc(fromDate)}'` +
        `${unit ? ` and ${n.Unit} eq '${esc(unit)}'` : ''}&$top=5000`
    );
    const out: Record<string, Record<string, ReportDayInfo>> = {};
    rows.forEach((it) => {
      const key = `${it.Unit}|${it.DocID}`.toLowerCase();
      const folder = getReportFolderRelativePath(this._libraryPath(), it.Unit, it.DocID, it.ReportDate);
      const done = it.Status === 'done';
      const info: ReportDayInfo = {
        dateStr: it.ReportDate,
        status: done ? 'done' : 'unsigned',
        savedAt: it.SavedAt || undefined,
        reporter: it.Reporter || undefined,
        reportNo: it.ReportNo > 0 ? formatReportNo(it.ReportNo) : undefined,
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
