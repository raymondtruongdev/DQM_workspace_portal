import * as React from 'react';
import { RoleEntry, ToastState } from '../../../core/types';
import { FilterChips, Icon, SearchInput } from '../../../core/components';
import { ToastAlert } from '../../../core/components/layout/ToastAlert';
import { buildExportRows, exportFileName, exportXlsx } from '../utils/reportExport';
import { ReportDataService } from '../services/ReportDataService';
import { ReportDayInfo, ReportDayStatus } from '../types/report.types';
import {
  buildRecentDates,
  DAY_STATUS_BADGE,
  DAY_STATUS_LABEL,
  formatDM,
  formatDMY,
  formatTimeRelative,
  monthShort,
  parseDateStr,
  weekdayLong
} from '../utils/reportCalendar';
import {
  buildTargets,
  buildReportList,
  countByStatus,
  countWindow,
  doneCount,
  matchesSearch,
  ownerDisplay,
  OverviewRow,
  runWithConcurrency,
  StatusFilter,
  statusOfDay
} from '../utils/reportOverview';

const WINDOW_DAYS = 14;
/** Số đơn vị đọc song song và số ngày đọc song song trong mỗi đơn vị (giữ số yêu cầu đồng thời ở mức vừa phải). */
const UNIT_CONCURRENCY = 3;
const DAY_CONCURRENCY = 5;
const CACHE_TTL_MS = 3 * 60 * 1000;
const PAGE_SIZE = 30;

type ListDateMode = 'ALL' | 'day';
/** Số người phụ trách hiện sẵn trong mỗi dòng; phần còn lại thu gọn thành "+N người khác". */
const OWNERS_COLLAPSED = 1;
/** Dưới bề rộng này (px, của chính khung danh sách) danh sách chuyển từ bảng 5 cột sang thẻ nhỏ. */
const COMPACT_WIDTH = 640;

interface IOverviewCache {
  key: string;
  at: number;
  rows: OverviewRow[];
}
// Nhớ kết quả vài phút trong bộ nhớ trang để chuyển qua lại không phải đọc lại (nút Làm mới bỏ qua)
const overviewCache: { current: IOverviewCache | undefined } = { current: undefined };

function saveOverviewCache(entry: IOverviewCache): void {
  overviewCache.current = entry;
}

export interface IReportOverviewViewProps {
  /** Toàn bộ phân công báo cáo (admin thấy mọi đơn vị). */
  assignments: RoleEntry[];
  dataService: ReportDataService;
  /** Mở lịch 14 ngày của một đơn vị. */
  onOpenUnit: (role: RoleEntry) => void;
  /** Mở form báo cáo của một đơn vị ở một ngày. */
  onOpenDate: (role: RoleEntry, dateStr: string) => void;
}

/** Tổng quan cho admin: hôm nay đơn vị nào thiếu gì, ma trận đơn vị × 14 ngày, danh sách cần xử lý. */
export const ReportOverviewView: React.FC<IReportOverviewViewProps> = ({
  assignments,
  dataService,
  onOpenUnit,
  onOpenDate
}) => {
  const dates = React.useMemo(() => buildRecentDates(WINDOW_DAYS), []);
  const todayStr = dates[dates.length - 1];
  const targets = React.useMemo(() => buildTargets(assignments), [assignments]);
  const cacheKey = React.useMemo(() => `${targets.map((t) => t.key).join(',')}@${dates[0]}`, [targets, dates]);

  const [rows, setRows] = React.useState<OverviewRow[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [filter, setFilter] = React.useState<StatusFilter>('ALL');
  const [search, setSearch] = React.useState<string>('');
  const [shown, setShown] = React.useState<number>(PAGE_SIZE);
  const [toast, setToast] = React.useState<ToastState | null>(null);
  const [listStatus, setListStatus] = React.useState<StatusFilter>('ALL');
  const [listDate, setListDate] = React.useState<ListDateMode>('ALL');
  // Ô chọn ngày luôn có giá trị, mặc định là hôm nay
  const [listDay, setListDay] = React.useState<string>(todayStr);
  const [listUnit, setListUnit] = React.useState<string>('ALL');
  const [expanded, setExpanded] = React.useState<Record<string, boolean>>({});
  const runId = React.useRef<number>(0);
  const listRef = React.useRef<HTMLDivElement | null>(null);
  const [compact, setCompact] = React.useState<boolean>(false);

  const load = React.useCallback(
    async (force: boolean): Promise<void> => {
      const cached = overviewCache.current;
      if (!force && cached && cached.key === cacheKey && Date.now() - cached.at < CACHE_TTL_MS) {
        setRows(cached.rows);
        return;
      }
      runId.current += 1;
      const myRun = runId.current;
      const initial: OverviewRow[] = targets.map((t) => ({ target: t, state: 'pending', infos: {} }));
      setRows(initial);
      setIsLoading(true);
      const current = initial.slice();
      await runWithConcurrency(targets, UNIT_CONCURRENCY, async (t) => {
        let next: OverviewRow;
        try {
          const list = await dataService.getDayInfos(t.unit, t.docId, dates, DAY_CONCURRENCY);
          const infos: Record<string, ReportDayInfo> = {};
          list.forEach((i) => {
            infos[i.dateStr] = i;
          });
          next = { target: t, state: 'ready', infos };
        } catch (err: unknown) {
          console.warn(`[ReportOverview] Could not read ${t.unit}:`, err);
          next = { target: t, state: 'error', infos: {} };
        }
        if (runId.current !== myRun) return;
        const idx = targets.indexOf(t);
        current[idx] = next;
        setRows(current.slice());
      });
      if (runId.current !== myRun) return;
      setIsLoading(false);
      if (current.every((r) => r.state === 'ready')) {
        saveOverviewCache({ key: cacheKey, at: Date.now(), rows: current.slice() });
      }
    },
    [cacheKey, dataService, dates, targets]
  );

  React.useEffect(() => {
    load(false).catch(() => {});
    return () => {
      runId.current += 1; // bỏ kết quả của lượt đọc đang dở khi rời màn hình
    };
  }, [load]);

  const readyCount = rows.filter((r) => r.state !== 'pending').length;
  const errorCount = rows.filter((r) => r.state === 'error').length;
  const today = countByStatus(rows, todayStr);
  const windowStats = countWindow(rows, dates);

  const visibleRows = rows.filter((r) => {
    if (!matchesSearch(r.target, search)) return false;
    if (filter === 'ALL' || r.state !== 'ready') return true;
    return dates.some((d) => statusOfDay(r, d) === filter);
  });
  // Danh sách "Cần xử lý" có bộ lọc riêng: trạng thái (Chưa nhập / Chưa upload bản ký) và đơn vị
  const allowedDates: string[] = listDate === 'day' ? dates.filter((d) => d === listDay) : dates;
  const listBase = buildReportList(rows, allowedDates, listStatus, '');
  const listUnitOptions = rows
    .map((r) => r.target.unit)
    .filter((u) => listBase.some((i) => i.target.unit === u));
  const activeListUnit = listUnit !== 'ALL' && listUnitOptions.indexOf(listUnit) === -1 ? 'ALL' : listUnit;
  const listItems = activeListUnit === 'ALL' ? listBase : listBase.filter((i) => i.target.unit === activeListUnit);

  // Chuyển giữa bảng 5 cột và thẻ nhỏ theo bề rộng thật của khung danh sách (không theo bề rộng cửa sổ)
  const hasList = listItems.length > 0;
  React.useEffect(() => {
    const el = listRef.current;
    if (!el) return undefined;
    const update = (): void => setCompact(el.clientWidth < COMPACT_WIDTH);
    update();
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', update);
      return () => window.removeEventListener('resize', update);
    }
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, [hasList]);

  const doExport = async (): Promise<void> => {
    const name = exportFileName('xlsx', activeListUnit);
    try {
      await exportXlsx(buildExportRows(listItems), name);
      setToast({ type: 'success', text: `Đã xuất ${listItems.length} dòng ra file ${name}.` });
    } catch (err: unknown) {
      console.error('Error exporting report list:', err);
      setToast({ type: 'error', text: 'Không xuất được file. Vui lòng thử lại.' });
    }
  };

  const filterOptions = [
    { value: 'ALL', label: 'Tất cả' },
    { value: 'missing', label: 'Chưa nhập' },
    { value: 'unsigned', label: 'Chưa upload bản ký' },
    { value: 'done', label: 'Hoàn tất' }
  ];

  const kpi = (
    status: ReportDayStatus,
    icon: 'error-circle' | 'file-up' | 'check-circle',
    value: number | undefined,
    percent?: number
  ): React.ReactNode => (
    <div className={`report-ov-kpi report-ov-kpi--${status}`}>
      <Icon name={icon} size={20} />
      <div>
        <b>{value === undefined ? '—' : value}</b>
        <span>{DAY_STATUS_LABEL[status]}</span>
        {percent !== undefined && <small>{percent}%</small>}
      </div>
    </div>
  );
  const pct = (n: number): number | undefined =>
    windowStats.total === 0 ? undefined : Math.round((n / windowStats.total) * 100);
  const hasRows = rows.some((r) => r.state === 'ready');

  if (targets.length === 0) {
    return (
      <div className="section-card">
        <div className="section-card__body">Chưa có phân công báo cáo ngày nào để theo dõi.</div>
      </div>
    );
  }

  return (
    <div className="report-ov-root">
      {toast && <ToastAlert toast={toast} onClose={() => setToast(null)} />}

      <section className="card report-ov-summary" aria-label="Tóm tắt hôm nay">
        <div className="report-ov-summary__title">
          <h3>Hôm nay · {weekdayLong(todayStr)}, {formatDMY(todayStr)}</h3>
          <span>
            {targets.length} đơn vị
            {isLoading && ` · đang đọc ${readyCount}/${targets.length}`}
          </span>
        </div>
        <div className="report-ov-kpis">
          {kpi('missing', 'error-circle', hasRows ? today.missing : undefined)}
          {kpi('unsigned', 'file-up', hasRows ? today.unsigned : undefined)}
          {kpi('done', 'check-circle', hasRows ? today.done : undefined)}
        </div>
        <button type="button" className="btn btn--secondary" onClick={() => load(true)} disabled={isLoading}>
          <Icon name="refresh" spin={isLoading} />
          <span>{isLoading ? 'Đang đọc...' : 'Làm mới'}</span>
        </button>
      </section>

      <section className="card report-ov-summary" aria-label={`Tóm tắt ${WINDOW_DAYS} ngày`}>
        <div className="report-ov-summary__title">
          <h3>
            {WINDOW_DAYS} ngày gần nhất · {formatDM(dates[0])} – {formatDMY(todayStr)}
          </h3>
          <span>
            {windowStats.total} báo cáo ({readyCount} đơn vị × {WINDOW_DAYS} ngày)
            {isLoading && ' · đang đọc'}
          </span>
        </div>
        <div className="report-ov-kpis">
          {kpi('missing', 'error-circle', hasRows ? windowStats.counts.missing : undefined, pct(windowStats.counts.missing))}
          {kpi('unsigned', 'file-up', hasRows ? windowStats.counts.unsigned : undefined, pct(windowStats.counts.unsigned))}
          {kpi('done', 'check-circle', hasRows ? windowStats.counts.done : undefined, pct(windowStats.counts.done))}
        </div>
      </section>

      {errorCount > 0 && (
        <div className="report-ov-warn" role="alert">
          Không đọc được dữ liệu của {errorCount} đơn vị (có thể do quyền hoặc mạng). Bấm Làm mới để thử lại.
        </div>
      )}

      <div className="section-card">
        <div className="section-card__banner">
          <span className="section-card__title">
            ĐƠN VỊ × {WINDOW_DAYS} NGÀY <span className="section-card__title-sub">| UNITS × DAYS</span>
          </span>
        </div>
        <div className="section-card__body">
          <div className="report-ov-toolbar">
            <div className="report-ov-toolbar__group">
              <span className="report-ov-toolbar__label">Làm nổi bật</span>
              <FilterChips
                ariaLabel="Làm nổi bật theo trạng thái trong ma trận"
                options={filterOptions}
                value={filter}
                onChange={(v) => setFilter(v as StatusFilter)}
              />
            </div>
            <SearchInput
              value={search}
              onChange={setSearch}
              placeholder="Tìm đơn vị, người phụ trách, email..."
            />
          </div>
          <div className="table-responsive">
            <table className="portal-table report-ov-matrix">
              <thead>
                <tr>
                  <th className="report-ov-matrix__unit">Đơn vị</th>
                  {dates.map((d) => (
                    <th key={d} className={d === todayStr ? 'is-today' : ''} title={`${weekdayLong(d)} ${formatDM(d)}`}>
                      <span>{parseDateStr(d).getDate()}</span>
                      <small>{monthShort(d)}</small>
                    </th>
                  ))}
                  <th>Hoàn tất</th>
                </tr>
              </thead>
              <tbody>
                {visibleRows.length === 0 && (
                  <tr>
                    <td colSpan={dates.length + 2}>Không có đơn vị phù hợp với bộ lọc.</td>
                  </tr>
                )}
                {visibleRows.map((r) => (
                  <tr key={r.target.key}>
                    <td className="report-ov-matrix__unit">
                      <button type="button" className="report-ov-unit-link" onClick={() => onOpenUnit(r.target.role)} title="Mở lịch của đơn vị này">
                        {r.target.unit}
                      </button>
                    </td>
                    {dates.map((d) => {
                      if (r.state !== 'ready') {
                        return (
                          <td key={d}>
                            <span className={`report-ov-cell report-ov-cell--${r.state}`} />
                          </td>
                        );
                      }
                      const st = statusOfDay(r, d);
                      const dim = filter !== 'ALL' && st !== filter;
                      return (
                        <td key={d}>
                          <button
                            type="button"
                            className={`report-ov-cell report-ov-cell--${st}${dim ? ' is-dim' : ''}${d === todayStr ? ' is-today' : ''}`}
                            title={`${r.target.unit} · ${weekdayLong(d)} ${formatDM(d)} · ${DAY_STATUS_LABEL[st]}`}
                            aria-label={`${r.target.unit} ${formatDM(d)} ${DAY_STATUS_LABEL[st]}`}
                            onClick={() => onOpenDate(r.target.role, d)}
                          />
                        </td>
                      );
                    })}
                    <td className="report-ov-matrix__rate">
                      {r.state === 'ready' ? `${doneCount(r, dates)}/${WINDOW_DAYS}` : r.state === 'error' ? 'Lỗi' : '…'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="report-cal-legend">
            <span><i className="report-cal-legend__dot report-cal-legend__dot--missing" />Chưa nhập</span>
            <span><i className="report-cal-legend__dot report-cal-legend__dot--unsigned" />Chưa upload bản ký</span>
            <span><i className="report-cal-legend__dot report-cal-legend__dot--done" />Hoàn tất</span>
            <span>Bấm một ô để mở báo cáo của đơn vị ở ngày đó; bấm mã đơn vị để mở lịch.</span>
          </div>
        </div>
      </div>

      <div className="section-card">
        <div className="section-card__banner">
          <span className="section-card__title">
            TẤT CẢ REPORT ({listItems.length}) <span className="section-card__title-sub">| ALL REPORT</span>
          </span>
          <button type="button" className="btn btn--light btn--sm" onClick={doExport} disabled={listItems.length === 0}>
            <Icon name="download" size={16} />
            <span>Xuất Excel</span>
          </button>
        </div>
        <div className="section-card__body">
          <div className="report-ov-list-filters">
            <div className="report-ov-list-filters__row">
              <span className="report-ov-list-filters__label">Trạng thái</span>
              <FilterChips
                ariaLabel="Lọc danh sách báo cáo theo trạng thái"
                value={listStatus}
                onChange={(v) => {
                  setListStatus(v as StatusFilter);
                  setShown(PAGE_SIZE);
                }}
                options={filterOptions}
              />
            </div>
            <div className="report-ov-list-filters__row">
              <span className="report-ov-list-filters__label">Ngày</span>
              <FilterChips
                ariaLabel="Lọc danh sách báo cáo theo ngày"
                value={listDate}
                onChange={(v) => {
                  setListDate(v as ListDateMode);
                  setShown(PAGE_SIZE);
                }}
                options={[
                  { value: 'ALL', label: `Tất cả (${WINDOW_DAYS} ngày)` },
                  { value: 'day', label: 'Theo ngày' }
                ]}
              />
              <input
                type="date"
                className="input report-ov-day-pick"
                aria-label="Chọn ngày"
                min={dates[0]}
                max={todayStr}
                value={listDay}
                onChange={(e) => {
                  if (e.target.value) {
                    setListDay(e.target.value);
                    setListDate('day');
                  } else {
                    setListDay(todayStr);
                    setListDate('ALL');
                  }
                  setShown(PAGE_SIZE);
                }}
              />
            </div>
            <div className="report-ov-list-filters__row">
              <span className="report-ov-list-filters__label">Đơn vị</span>
              <FilterChips
                ariaLabel="Lọc danh sách báo cáo theo đơn vị"
                value={activeListUnit}
                onChange={(v) => {
                  setListUnit(v);
                  setShown(PAGE_SIZE);
                }}
                options={[
                  { value: 'ALL', label: `Tất cả (${listBase.length})` },
                  ...listUnitOptions.map((u) => ({
                    value: u,
                    label: `${u} (${listBase.filter((i) => i.target.unit === u).length})`
                  }))
                ]}
              />
            </div>
          </div>
          {listItems.length === 0 ? (
            <p className="report-ov-note">
              {isLoading ? 'Đang đọc dữ liệu...' : 'Không có báo cáo nào theo bộ lọc hiện tại.'}
            </p>
          ) : (
            <div ref={listRef} className={`report-ov-list${compact ? ' is-compact' : ''}`} role="list" aria-label="Danh sách báo cáo">
              {listItems.slice(0, shown).map((i) => {
                const open = Boolean(expanded[i.target.key]);
                const owners = open ? i.target.owners : i.target.owners.slice(0, OWNERS_COLLAPSED);
                const hidden = i.target.owners.length - OWNERS_COLLAPSED;
                const isToday = i.dateStr === todayStr;
                const savedText = i.info && i.info.savedAt ? formatTimeRelative(i.info.savedAt, i.dateStr) : '';
                const signedText = i.info && i.info.signed ? formatTimeRelative(i.info.signed.timeLastModified, i.dateStr) : '';
                const openReport = (): void => onOpenDate(i.target.role, i.dateStr);
                return (
                  <div
                    key={`${i.target.key}|${i.dateStr}`}
                    className={`report-ov-row report-ov-row--${i.status}${isToday ? ' is-today' : ''}`}
                    role="listitem"
                  >
                    <div className="report-ov-daybox">
                      <b>{formatDM(i.dateStr)}</b>
                      <small>{isToday ? 'Hôm nay' : weekdayLong(i.dateStr)}</small>
                    </div>

                    <div className="report-ov-main">
                      <div className="report-ov-main__line">
                        <span className="report-ov-unit">{i.target.unit}</span>
                        <span className={`badge ${DAY_STATUS_BADGE[i.status]} report-ov-dotted`}>{DAY_STATUS_LABEL[i.status]}</span>
                      </div>
                      <div className="report-ov-main__line">
                        {owners.map((o, k) => {
                          const od = ownerDisplay(o);
                          return (
                            <a
                              key={`${o.email}-${k}`}
                              className="report-ov-person report-ov-tip"
                              href={o.email ? `mailto:${o.email}` : undefined}
                              data-tip={od.title}
                              aria-label={od.title.replace('\n', ', ')}
                            >
                              <span className="report-ov-person__avatar" aria-hidden="true">
                                {od.initials}
                              </span>
                              <span className="report-ov-person__name">{od.label}</span>
                            </a>
                          );
                        })}
                        {i.target.owners.length === 0 && <span className="report-ov-times">Chưa có người phụ trách</span>}
                        {hidden > 0 && (
                          <button
                            type="button"
                            className="report-ov-owner-more"
                            onClick={() => setExpanded({ ...expanded, [i.target.key]: !open })}
                          >
                            {open ? 'Thu gọn' : `+${hidden}`}
                          </button>
                        )}
                        <span className="report-ov-times">
                          {i.status === 'missing' ? (
                            <span>Chưa lưu</span>
                          ) : (
                            <>
                              {savedText && <span>Lưu {savedText}</span>}
                              {signedText && <span>Ký {signedText}</span>}
                            </>
                          )}
                        </span>
                      </div>
                    </div>

                    <div className="report-ov-action">
                      {i.status === 'missing' && (
                        <button type="button" className="btn btn--primary" onClick={openReport}>
                          Nhập báo cáo
                        </button>
                      )}
                      {i.status === 'unsigned' && (
                        <>
                          <button
                            type="button"
                            className="btn btn--secondary btn--icon report-ov-tip report-ov-tip--end"
                            onClick={openReport}
                            data-tip="Xem báo cáo"
                            aria-label="Xem báo cáo"
                          >
                            <Icon name="eye" size={20} />
                          </button>
                          <button type="button" className="btn btn--primary" onClick={openReport}>
                            <Icon name="upload" size={16} />
                            <span>Upload bản ký</span>
                          </button>
                        </>
                      )}
                      {i.status === 'done' && (
                        <>
                          <button type="button" className="btn btn--secondary" onClick={openReport}>
                            Xem báo cáo
                          </button>
                          {i.info && i.info.signed && (
                            <a
                              className="btn btn--secondary btn--icon report-ov-tip report-ov-tip--end"
                              href={i.info.signed.serverRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              data-tip="Tải bản ký"
                              aria-label="Tải bản ký"
                            >
                              <Icon name="download" size={20} />
                            </a>
                          )}
                        </>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
          {listItems.length > shown && (
            <div className="report-ov-more">
              <button type="button" className="btn btn--secondary" onClick={() => setShown(shown + PAGE_SIZE)}>
                Hiển thị thêm ({listItems.length - shown})
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
