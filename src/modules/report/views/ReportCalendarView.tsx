import * as React from 'react';
import { RoleEntry, ToastState } from '../../../core/types';
import {
  ConfirmDialog,
  FilterChips,
  Icon,
  LoadingState,
  PageHeader
} from '../../../core/components';
import { ToastAlert } from '../../../core/components/layout/ToastAlert';
import { ReportDataService } from '../services/ReportDataService';
import { ReportDayInfo, ReportDayStatus } from '../types/report.types';
import {
  addDaysStr,
  addMonthsKey,
  buildRecentDates,
  DAY_STATUS_BADGE,
  DAY_STATUS_LABEL,
  formatDM,
  formatDMY,
  formatFileSize,
  formatMonthKey,
  formatSavedAt,
  mondayIndex,
  monthDatesBefore,
  monthKeyOf,
  parseDateStr,
  WEEKDAY_SHORT_MON_FIRST,
  weekdayLong
} from '../utils/reportCalendar';
import { DEFAULT_WINDOW_DAYS, formatUnlockLeft, formatUnlockUntil } from '../utils/reportSettings';
import { getReportFolderName } from '../utils/reportFileNaming';

export interface IReportCalendarViewProps {
  role: RoleEntry;
  dataService: ReportDataService;
  /** true = portal đang bảo trì, không được ghi lên SharePoint. */
  onCheckMaintenance: () => Promise<boolean>;
  onBack: () => void;
  /** Mở form nhập báo cáo của một ngày (YYYY-MM-DD). */
  onOpenDate: (dateStr: string) => void;
  /** Email người đang đăng nhập (ghi vào sổ khi upload bản ký). */
  currentUserEmail: string;
  /** Được khoá/mở khoá (admin thật và đang xem bằng hồ sơ admin) báo cáo từng ngày. */
  canManageLock?: boolean;
  /** Số ngày gần nhất hiển thị (admin đặt ở Admin Config; mặc định 14). */
  windowDays?: number;
}

type CalendarMode = 'cal' | 'list' | 'history';

/** Lịch sử lùi tối đa bấy nhiêu tháng so với cửa sổ ngày gần nhất. */
const MAX_HISTORY_MONTHS = 36;

function StatusIcon({ status }: { status: ReportDayStatus }): React.ReactElement {
  const name = status === 'done' ? 'check-circle' : status === 'unsigned' ? 'file-up' : 'error-circle';
  return <Icon name={name} size={14} />;
}

/**
 * Lịch 14 ngày gần nhất của một báo cáo: mỗi ngày là Chưa nhập / Chưa upload bản ký / Hoàn tất.
 * Bấm một ngày để xem chi tiết, mở báo cáo hoặc upload file PDF đã ký.
 */
export const ReportCalendarView: React.FC<IReportCalendarViewProps> = ({
  role,
  dataService,
  onCheckMaintenance,
  onBack,
  onOpenDate,
  currentUserEmail,
  canManageLock = false,
  windowDays = DEFAULT_WINDOW_DAYS
}) => {
  const WINDOW_DAYS = windowDays;
  const dates = React.useMemo(() => buildRecentDates(WINDOW_DAYS, dataService.todayDate()), [WINDOW_DAYS, dataService]);
  const todayStr = dates[dates.length - 1];

  const [infos, setInfos] = React.useState<ReportDayInfo[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isUploading, setIsUploading] = React.useState<boolean>(false);
  const [mode, setMode] = React.useState<CalendarMode>('cal');
  const [selected, setSelected] = React.useState<string>(todayStr);
  const [toast, setToast] = React.useState<ToastState | null>(null);
  const [lockConfirm, setLockConfirm] = React.useState<'lock' | 'unlock' | null>(null);
  const [isTogglingLock, setIsTogglingLock] = React.useState<boolean>(false);
  // Lịch sử: MỌI ngày đến hôm nay, duyệt theo tháng; ngày cũ hơn cửa sổ `windowDays` chỉ xem, ngày trong cửa sổ thao tác như bình thường
  const newestHistMonth = monthKeyOf(todayStr);
  const oldestHistMonth = addMonthsKey(newestHistMonth, -MAX_HISTORY_MONTHS);
  const [histMonth, setHistMonth] = React.useState<string>(newestHistMonth);
  const [histInfos, setHistInfos] = React.useState<ReportDayInfo[]>([]);
  const [isHistLoading, setIsHistLoading] = React.useState<boolean>(false);
  const histDates = React.useMemo(() => monthDatesBefore(histMonth, addDaysStr(todayStr, 1)), [histMonth, todayStr]);
  const fileRef = React.useRef<HTMLInputElement | null>(null);
  const uploadTarget = React.useRef<string>(todayStr);

  const load = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    try {
      setInfos(await dataService.getDayInfos(role.unit, role.docId, dates));
    } catch (err: unknown) {
      console.error('Error loading report calendar:', err);
      const msg = err instanceof Error ? err.message : 'Không tải được trạng thái các ngày';
      setToast({ type: 'error', text: msg });
    } finally {
      setIsLoading(false);
    }
  }, [dataService, role.unit, role.docId, dates]);

  React.useEffect(() => {
    load().catch(() => {});
  }, [load]);

  React.useEffect(() => {
    if (mode !== 'history' || histDates.length === 0) return;
    let alive = true;
    setIsHistLoading(true);
    dataService
      .getDayInfos(role.unit, role.docId, histDates)
      .then((r) => {
        if (!alive) return;
        setHistInfos(r);
        setIsHistLoading(false);
      })
      .catch((err: unknown) => {
        console.error('Error loading report history:', err);
        if (!alive) return;
        setIsHistLoading(false);
        setToast({ type: 'error', text: err instanceof Error ? err.message : 'Không tải được lịch sử báo cáo' });
      });
    return () => {
      alive = false;
    };
  }, [mode, histDates, dataService, role.unit, role.docId]);

  const byDate = React.useMemo(() => {
    const m: Record<string, ReportDayInfo> = {};
    infos.forEach((i) => {
      m[i.dateStr] = i;
    });
    return m;
  }, [infos]);

  // Đếm ngược các ngày đang mở khoá có hạn: cập nhật mỗi 30 giây và đọc lại lịch khi có ngày vừa hết hạn (tự khoá lại)
  const [nowMs, setNowMs] = React.useState<number>(() => dataService.now());
  const nextExpiry = React.useMemo(() => {
    const ts = infos.filter((i) => i.unlockUntil).map((i) => Date.parse(i.unlockUntil as string)).filter((t) => Number.isFinite(t));
    return ts.length ? Math.min(...ts) : undefined;
  }, [infos]);
  React.useEffect(() => {
    if (nextExpiry === undefined) return;
    setNowMs(dataService.now());
    const tick = window.setInterval(() => setNowMs(dataService.now()), 30000);
    const wait = Math.max(0, Math.min(nextExpiry - dataService.now() + 1000, 2147483000));
    const expire = window.setTimeout(() => {
      load().catch(() => {});
    }, wait);
    return () => {
      window.clearInterval(tick);
      window.clearTimeout(expire);
    };
  }, [nextExpiry, load, dataService]);

  const statusOf = (dateStr: string): ReportDayStatus => (byDate[dateStr] ? byDate[dateStr].status : 'missing');

  const counts = { done: 0, unsigned: 0, missing: 0 };
  dates.forEach((d) => {
    counts[statusOf(d)] += 1;
  });
  const missingDates = dates.filter((d) => statusOf(d) === 'missing');
  const unsignedDates = dates.filter((d) => statusOf(d) === 'unsigned');

  const selInfo = byDate[selected];
  const selStatus = statusOf(selected);
  const selLocked = Boolean(selInfo && selInfo.locked);
  // Sổ DRPRegistry không lưu tên file đã ký (<DocCode>_signed.pdf): đọc thư mục khi chọn ngày đã hoàn tất
  const [resolvedName, setResolvedName] = React.useState<{ date: string; name: string } | null>(null);
  const needName = selStatus === 'done' && !!selInfo && !!selInfo.signed && !selInfo.signed.name;
  React.useEffect(() => {
    if (!needName) return;
    let alive = true;
    dataService
      .getSignedDocument(role.unit, role.docId, selected)
      .then((d) => {
        if (alive && d) setResolvedName({ date: selected, name: d.name });
      })
      .catch((err) => console.error(err));
    return () => {
      alive = false;
    };
  }, [needName, selected, role.unit, role.docId, dataService]);
  const signedLabel =
    selInfo && selInfo.signed
      ? selInfo.signed.name || (resolvedName && resolvedName.date === selected ? resolvedName.name : 'Đang đọc tên file…')
      : '';
  const openSigned = (dateStr: string): void => {
    dataService
      .openSignedDocument(role.unit, role.docId, dateStr)
      .then((ok) => {
        if (!ok) setToast({ type: 'error', text: 'Không tìm thấy file PDF đã ký trong thư mục báo cáo.' });
      })
      .catch((err) => {
        console.error(err);
        setToast({ type: 'error', text: 'Không mở được file đã ký.' });
      });
  };
  const daySteps = [
    { label: 'Lưu lần cuối', value: selInfo && selInfo.savedAt ? formatSavedAt(selInfo.savedAt) : '' },
    { label: 'Upload bản ký', value: selInfo && selInfo.signed && selInfo.signed.timeLastModified ? formatSavedAt(selInfo.signed.timeLastModified) : '' },
    {
      label: 'File PDF đã ký',
      value: selInfo && selInfo.signed
        ? `${signedLabel}${selInfo.signed.length > 0 ? ` · ${formatFileSize(selInfo.signed.length)}` : ''}`
        : ''
    }
  ];
  const copyPath = (text: string): void => {
    const done = (): void => setToast({ type: 'success', text: 'Đã sao chép đường dẫn thư mục.' });
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, () => setToast({ type: 'error', text: 'Không sao chép được.' }));
    }
  };
  const folderLabel = `data/${role.unit}/${getReportFolderName(role.docId, role.unit, selected)}/`;

  // Admin bật/tắt khoá báo cáo của ngày đang chọn (cần list DRPRegistry có cột Locked)
  const toggleLock = async (): Promise<void> => {
    const target = lockConfirm === 'lock';
    setLockConfirm(null);
    if (await onCheckMaintenance()) return;
    setIsTogglingLock(true);
    try {
      await dataService.setReportLocked(role.unit, role.docId, selected, target);
      setToast({
        type: 'success',
        text: target ? `Đã khoá báo cáo ngày ${formatDM(selected)}: người dùng chỉ xem được.` : `Đã mở khoá báo cáo ngày ${formatDM(selected)}: người dùng có thể chỉnh sửa lại.`
      });
      await load();
    } catch (err: unknown) {
      console.error('Error toggling report lock:', err);
      setToast({ type: 'error', text: err instanceof Error ? err.message : 'Chưa bật/tắt được khoá báo cáo.' });
    } finally {
      setIsTogglingLock(false);
    }
  };

  const pickFile = (dateStr: string): void => {
    // Báo cáo đã Lưu mà chưa ký (Chưa upload bản ký) vẫn upload được khi khoá; đã ký rồi thì cần admin mở khoá
    if (byDate[dateStr] && byDate[dateStr].locked && statusOf(dateStr) !== 'unsigned') {
      setToast({ type: 'error', text: 'Báo cáo ngày này đã khoá. Liên hệ admin để mở khoá nếu cần thay bản ký.' });
      return;
    }
    uploadTarget.current = dateStr;
    setSelected(dateStr);
    if (fileRef.current) {
      fileRef.current.value = '';
      fileRef.current.click();
    }
  };

  const handleFile = async (file: File | undefined): Promise<void> => {
    if (!file) return;
    const target = uploadTarget.current;
    if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') {
      setToast({ type: 'error', text: 'Chỉ nhận file PDF.' });
      return;
    }
    if (statusOf(target) === 'missing') {
      setToast({ type: 'error', text: 'Ngày này chưa có báo cáo. Hãy nhập và lưu báo cáo trước.' });
      return;
    }
    if (await onCheckMaintenance()) return;
    setIsUploading(true);
    try {
      const buf = await file.arrayBuffer();
      await dataService.uploadSignedDocument(role.unit, role.docId, target, buf, '', '', '', currentUserEmail);
      setToast({ type: 'success', text: `Đã upload bản ký cho ngày ${formatDM(target)}. Ngày này đã hoàn tất.` });
      await load();
    } catch (err: unknown) {
      console.error('Error uploading signed document:', err);
      const msg = err instanceof Error ? err.message : 'Lỗi khi upload file đã ký lên SharePoint';
      setToast({ type: 'error', text: msg });
    } finally {
      setIsUploading(false);
    }
  };

  const lead = mondayIndex(dates[0]);
  const tail = (7 - ((lead + WINDOW_DAYS) % 7)) % 7;
  const futureDays: string[] = [];
  for (let i = 1; i <= tail; i++) futureDays.push(addDaysStr(todayStr, i));

  const dayButton = (dateStr: string): React.ReactNode => {
    const st = statusOf(dateStr);
    const info = byDate[dateStr];
    // Ngày chưa nhập mà đã khoá: màu xám riêng để không lẫn với "Chưa nhập" (đỏ, người dùng nhập được)
    const lockedMissing = st === 'missing' && Boolean(info && info.locked);
    return (
      <button
        key={dateStr}
        type="button"
        className={`report-cal-day report-cal-day--${lockedMissing ? 'locked' : st}${dateStr === todayStr ? ' is-today' : ''}`}
        aria-pressed={selected === dateStr}
        onClick={() => setSelected(dateStr)}
      >
        {dateStr === todayStr && <span className="report-cal-day__today">Hôm nay</span>}
        <span className="report-cal-day__num">{parseDateStr(dateStr).getDate()}</span>
        <span className="report-cal-day__wd">
          {weekdayLong(dateStr)} · {formatDM(dateStr)}
        </span>
        <span className="report-cal-day__status">
          <StatusIcon status={st} />
          {lockedMissing ? 'Đã khoá' : DAY_STATUS_LABEL[st]}
          {info && info.locked && <Icon name="lock" size={14} className="report-cal-day__lock" />}
        </span>
        {lockedMissing && <span className="report-cal-day__note">Chưa nhập · chỉ xem</span>}
        {st === 'unsigned' && info && info.locked && <span className="report-cal-day__sign">Khoá nội dung · vẫn upload bản ký</span>}
        {info && info.unlockUntil && (
          <span className="report-cal-day__open" title={`Admin đã mở khoá, tự khoá lại lúc ${formatUnlockUntil(info.unlockUntil)}`}>
            <Icon name="lock-open" size={14} />
            Mở khoá · còn {formatUnlockLeft(info.unlockUntil, nowMs)}
          </span>
        )}
        {info && info.savedAt && <span className="report-cal-day__note">{formatSavedAt(info.savedAt)}</span>}
      </button>
    );
  };

  return (
    <div className="report-cal-root">
      <PageHeader
        divider
        leading={
          <button type="button" className="btn btn--secondary btn--icon" onClick={onBack} aria-label="Quay lại" title="Quay lại danh sách báo cáo">
            <Icon name="arrow-left" />
          </button>
        }
        title={`${role.unit} · ${role.docTitle || 'Báo cáo ngày'}`}
        badge={<span className="badge badge--info">{role.docId}</span>}
        meta={
          <span>
            {formatDM(dates[0])} – {formatDMY(todayStr)} · {weekdayLong(todayStr)}
          </span>
        }
        actions={
          <>
            <button type="button" className="btn btn--secondary" onClick={() => load()} disabled={isLoading || isUploading} aria-label="Làm mới">
              <Icon name="refresh" spin={isLoading} />
              <span>Làm mới</span>
            </button>
            <button type="button" className="btn btn--primary" onClick={() => onOpenDate(todayStr)}>
              <Icon name="add" />
              <span>Nhập báo cáo hôm nay</span>
            </button>
          </>
        }
      />

      {toast && <ToastAlert toast={toast} onClose={() => setToast(null)} />}

      {isLoading && infos.length === 0 ? (
        <LoadingState message={`Đang kiểm tra các báo cáo của ${WINDOW_DAYS} ngày gần nhất...`} />
      ) : (
        <>
          <section className="card report-cal-summary" aria-label={`Tóm tắt ${WINDOW_DAYS} ngày`}>
            <div className="report-cal-summary__big">
              <b>{counts.done}</b>
              <span>/ {WINDOW_DAYS} ngày hoàn tất</span>
            </div>
            <div className="report-cal-strip" role="img" aria-label={`Trạng thái ${WINDOW_DAYS} ngày`}>
              {dates.map((d) => (
                <i key={d} className={`report-cal-strip__seg report-cal-strip__seg--${statusOf(d)}`} title={`${formatDM(d)} · ${DAY_STATUS_LABEL[statusOf(d)]}`} />
              ))}
            </div>
            <div className="report-cal-counts">
              <div className="report-cal-count report-cal-count--done">
                <b>{counts.done}</b>
                <span>Hoàn tất</span>
              </div>
              <div className="report-cal-count report-cal-count--unsigned">
                <b>{counts.unsigned}</b>
                <span>Chưa upload bản ký</span>
              </div>
              <div className="report-cal-count report-cal-count--missing">
                <b>{counts.missing}</b>
                <span>Chưa nhập</span>
              </div>
            </div>
          </section>

          <div className="report-cal-alerts" role="status">
            {missingDates.length > 0 && (
              <div className="report-cal-alert report-cal-alert--missing">
                <strong>{missingDates.length} ngày chưa nhập báo cáo:</strong>
                <div className="report-cal-alert__days">
                  {missingDates.map((d) => (
                    <button key={d} type="button" className="btn btn--secondary btn--sm" onClick={() => { setSelected(d); setMode('cal'); }}>
                      {formatDM(d)}
                    </button>
                  ))}
                </div>
              </div>
            )}
            {unsignedDates.length > 0 && (
              <div className="report-cal-alert report-cal-alert--unsigned">
                <strong>{unsignedDates.length} ngày chưa upload bản đã ký:</strong>
                <div className="report-cal-alert__days">
                  {unsignedDates.map((d) => (
                    <button key={d} type="button" className="btn btn--secondary btn--sm" onClick={() => { setSelected(d); setMode('cal'); }}>
                      {formatDM(d)}
                    </button>
                  ))}
                </div>
              </div>
            )}
            {missingDates.length === 0 && unsignedDates.length === 0 && (
              <div className="report-cal-alert report-cal-alert--done">
                <strong>Cả {WINDOW_DAYS} ngày đều hoàn tất.</strong>
              </div>
            )}
          </div>

          <div className="section-card">
            <div className="section-card__banner">
              <span className="section-card__title">
                {mode === 'history' ? (
                  <>LỊCH SỬ BÁO CÁO <span className="section-card__title-sub">| REPORT HISTORY</span></>
                ) : (
                  <>{WINDOW_DAYS} NGÀY GẦN NHẤT <span className="section-card__title-sub">| LAST {WINDOW_DAYS} DAYS</span></>
                )}
              </span>
              <FilterChips
                ariaLabel="Chế độ xem"
                value={mode}
                onChange={(v) => setMode(v as CalendarMode)}
                options={[
                  { value: 'cal', label: 'Lịch' },
                  { value: 'list', label: 'Danh sách' },
                  { value: 'history', label: 'Lịch sử' }
                ]}
              />
            </div>
            <div className="section-card__body">
              {mode === 'history' ? (
                <>
                  <div className="report-cal-month">
                    <button
                      type="button"
                      className="btn btn--secondary btn--icon btn--sm"
                      onClick={() => setHistMonth(addMonthsKey(histMonth, -1))}
                      disabled={histMonth <= oldestHistMonth}
                      aria-label="Tháng trước"
                    >
                      <Icon name="chevron-left" size={16} />
                    </button>
                    <strong className="report-cal-month__label">{formatMonthKey(histMonth)}</strong>
                    <button
                      type="button"
                      className="btn btn--secondary btn--icon btn--sm"
                      onClick={() => setHistMonth(addMonthsKey(histMonth, 1))}
                      disabled={histMonth >= newestHistMonth}
                      aria-label="Tháng sau"
                    >
                      <Icon name="chevron-right" size={16} />
                    </button>
                    {isHistLoading && <Icon name="refresh" size={16} spin />}
                  </div>
                  <p className="report-cal-history-note">
                    <Icon name="lock" size={16} />
                    <span>
                      Báo cáo cũ hơn {WINDOW_DAYS} ngày (trước {formatDMY(dates[0])}) chỉ xem, không chỉnh sửa được; cần sửa thì liên hệ admin để mở khoá. {WINDOW_DAYS} ngày gần nhất thao tác như bình thường.
                    </span>
                  </p>
                  {isHistLoading ? (
                    <LoadingState message={`Đang tải lịch sử ${formatMonthKey(histMonth)}...`} />
                  ) : histDates.length === 0 ? (
                    <p className="report-cal-history-note">Tháng này chưa có ngày nào.</p>
                  ) : (
                    <div className="table-responsive">
                      <table className="portal-table">
                        <thead>
                          <tr>
                            <th>Ngày</th>
                            <th>Trạng thái</th>
                            <th>Lưu lúc</th>
                            <th>File đã ký</th>
                            <th />
                          </tr>
                        </thead>
                        <tbody>
                          {histDates.map((d) => {
                            const info = histInfos.find((i) => i.dateStr === d);
                            const st: ReportDayStatus = info ? info.status : 'missing';
                            const old = d < dates[0];
                            return (
                              <tr key={d}>
                                <td>
                                  <strong>{formatDM(d)}</strong> {weekdayLong(d)}
                                </td>
                                <td>
                                  <span className={`badge ${DAY_STATUS_BADGE[st]}`}>
                                    <StatusIcon status={st} />
                                    {DAY_STATUS_LABEL[st]}
                                  </span>
                                  {info && info.locked && <Icon name="lock" size={14} className="report-cal-day__lock" />}
                                  {st === 'missing' && old && <span className="report-cal-day__sign">chỉ xem</span>}
                                </td>
                                <td>{info && info.savedAt ? formatSavedAt(info.savedAt) : '—'}</td>
                                <td>{info && info.signed ? 'Có' : '—'}</td>
                                <td>
                                  {!(st === 'missing' && old) && (
                                    <span className="report-cal-history-actions">
                                      <button
                                        type="button"
                                        className={`btn btn--sm ${st === 'missing' ? 'btn--primary' : 'btn--secondary'}`}
                                        onClick={() => onOpenDate(d)}
                                      >
                                        {st === 'missing' ? 'Nhập' : old ? 'Xem báo cáo' : 'Mở'}
                                      </button>
                                      {st === 'done' && (
                                        <button
                                          type="button"
                                          className="btn btn--secondary btn--icon btn--sm"
                                          onClick={() => openSigned(d)}
                                          title="Tải file đã ký"
                                          aria-label={`Tải file đã ký ngày ${formatDM(d)}`}
                                        >
                                          <Icon name="download" size={16} />
                                        </button>
                                      )}
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </>
              ) : mode === 'cal' ? (
                <>
                  <div className="report-cal-weekdays" aria-hidden="true">
                    {WEEKDAY_SHORT_MON_FIRST.map((w) => (
                      <span key={w}>{w}</span>
                    ))}
                  </div>
                  <div className="report-cal-grid">
                    {Array.from({ length: lead }).map((_, i) => (
                      <div key={`lead-${i}`} className="report-cal-day report-cal-day--blank" aria-hidden="true" />
                    ))}
                    {dates.map((d) => dayButton(d))}
                    {futureDays.map((d) => (
                      <div key={d} className="report-cal-day report-cal-day--future" aria-hidden="true">
                        <span className="report-cal-day__num">{parseDateStr(d).getDate()}</span>
                        <span className="report-cal-day__note">Chưa đến</span>
                      </div>
                    ))}
                  </div>
                </>
              ) : (
                <div className="table-responsive">
                  <table className="portal-table">
                    <thead>
                      <tr>
                        <th>Ngày</th>
                        <th>Trạng thái</th>
                        <th>Lưu lúc</th>
                        <th>File đã ký</th>
                        <th />
                      </tr>
                    </thead>
                    <tbody>
                      {dates
                        .slice()
                        .reverse()
                        .map((d) => {
                          const st = statusOf(d);
                          const info = byDate[d];
                          return (
                            <tr key={d} className={selected === d ? 'report-cal-row-selected' : ''} onClick={() => setSelected(d)}>
                              <td>
                                <strong>{formatDM(d)}</strong> {weekdayLong(d)}
                              </td>
                              <td>
                                <span className={`badge ${DAY_STATUS_BADGE[st]}`}>
                                  <StatusIcon status={st} />
                                  {DAY_STATUS_LABEL[st]}
                                </span>
                                {info && info.locked && <Icon name="lock" size={14} className="report-cal-day__lock" />}
                                {st === 'unsigned' && info && info.locked && <span className="report-cal-day__sign">Khoá nội dung · vẫn upload bản ký</span>}
                                {info && info.unlockUntil && (
                                  <span className="report-cal-day__open">
                                    <Icon name="lock-open" size={14} />
                                    Mở khoá đến {formatUnlockUntil(info.unlockUntil)} · còn {formatUnlockLeft(info.unlockUntil, nowMs)}
                                  </span>
                                )}
                              </td>
                              <td>{info ? formatSavedAt(info.savedAt) : '—'}</td>
                              <td>{info && info.signed ? 'Có' : '—'}</td>
                              <td>
                                {st === 'unsigned' ? (
                                  <button type="button" className="btn btn--primary btn--sm" onClick={(e) => { e.stopPropagation(); pickFile(d); }} disabled={isUploading}>
                                    Upload bản ký
                                  </button>
                                ) : (
                                  <button type="button" className={`btn btn--sm ${st === 'missing' ? 'btn--primary' : 'btn--secondary'}`} onClick={(e) => { e.stopPropagation(); onOpenDate(d); }}>
                                    {st === 'missing' ? 'Nhập' : 'Mở'}
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              )}

              {mode !== 'history' && (
              <div className="report-cal-legend">
                <span><i className="report-cal-legend__dot report-cal-legend__dot--missing" />Chưa nhập: chưa bấm Lưu</span>
                <span><i className="report-cal-legend__dot report-cal-legend__dot--unsigned" />Chưa upload bản ký: đã lưu, thiếu file PDF đã ký</span>
                <span><i className="report-cal-legend__dot report-cal-legend__dot--done" />Hoàn tất: có file PDF trong thư mục báo cáo</span>
                <span><i className="report-cal-legend__dot report-cal-legend__dot--locked" /><Icon name="lock" size={14} /> Đã khoá: chỉ xem, admin mới mở khoá được. Đã lưu mà chưa ký thì vẫn upload được bản ký</span>
              </div>
              )}
            </div>
          </div>

          {mode !== 'history' && (
          <div className="card report-day-panel" aria-live="polite">
            <div className="report-day-panel__top">
              <div className="report-day-panel__tile" aria-hidden="true">
                <span>T{parseDateStr(selected).getMonth() + 1}</span>
                <b>{parseDateStr(selected).getDate()}</b>
              </div>
              <div className="report-day-panel__title">
                <span className="report-day-panel__eyebrow">Ngày đang chọn</span>
                <div className="report-day-panel__name">
                  <h3>
                    {weekdayLong(selected)}, {formatDMY(selected)}
                  </h3>
                  <span className={`badge ${DAY_STATUS_BADGE[selStatus]}`}>
                    <StatusIcon status={selStatus} />
                    {DAY_STATUS_LABEL[selStatus]}
                  </span>
                  {selInfo && selInfo.unlockUntil && (
                    <span className="badge badge--success" title="Admin đã mở khoá có hạn, hết giờ báo cáo tự khoá lại">
                      <Icon name="lock-open" size={14} />
                      Mở khoá đến {formatUnlockUntil(selInfo.unlockUntil)} · còn {formatUnlockLeft(selInfo.unlockUntil, nowMs)}
                    </span>
                  )}
                  {selLocked && (
                    <span
                      className={`badge ${selStatus === 'unsigned' ? 'badge--info' : 'badge--danger'}`}
                      title={selStatus === 'unsigned' ? 'Nội dung đã khoá (chỉ xem) nhưng vẫn upload được bản đã ký' : 'Báo cáo đã khoá: chỉ xem, không chỉnh sửa được'}
                    >
                      <Icon name="lock" size={14} />
                      {selStatus === 'unsigned' ? 'Khoá nội dung · vẫn upload bản ký' : 'Đã khoá'}
                    </span>
                  )}
                </div>
              </div>
              <div className="report-day-panel__actions">
                {selStatus === 'missing' && (
                  <button type="button" className="btn btn--primary" onClick={() => onOpenDate(selected)}>
                    <Icon name="add" />
                    <span>Nhập báo cáo</span>
                  </button>
                )}
                {selStatus === 'unsigned' && (
                  <>
                    <button type="button" className="btn btn--primary" onClick={() => pickFile(selected)} disabled={isUploading}>
                      <Icon name="upload" />
                      <span>{isUploading ? 'Đang tải lên...' : 'Upload bản ký'}</span>
                    </button>
                    <button type="button" className="btn btn--secondary" onClick={() => onOpenDate(selected)}>
                      Mở báo cáo
                    </button>
                  </>
                )}
                {selStatus === 'done' && (
                  <>
                    <button type="button" className="btn btn--primary" onClick={() => onOpenDate(selected)}>
                      Mở báo cáo
                    </button>
                    {selInfo && selInfo.signed && (
                      <button
                        type="button"
                        className="btn btn--secondary btn--icon"
                        onClick={() => openSigned(selected)}
                        title="Tải file đã ký"
                        aria-label="Tải file đã ký"
                      >
                        <Icon name="download" />
                      </button>
                    )}
                  </>
                )}
                {canManageLock && (
                  <button
                    type="button"
                    className={`btn btn--secondary report-lock-toggle report-lock-toggle--${selLocked ? 'locked' : 'open'}`}
                    onClick={() => setLockConfirm(selLocked ? 'unlock' : 'lock')}
                    disabled={isTogglingLock || isUploading}
                    title={selLocked ? 'Đang khoá (người dùng chỉ xem). Bấm để mở khoá cho phép chỉnh sửa lại' : 'Đang mở (người dùng sửa được). Bấm để khoá: người dùng chỉ xem được'}
                    aria-label={selLocked ? 'Trạng thái LOCKED, bấm để mở khoá' : 'Trạng thái OPEN, bấm để khoá'}
                  >
                    <Icon name={selLocked ? 'lock' : 'lock-open'} />
                    <span>{selLocked ? 'LOCKED' : selInfo && selInfo.unlockUntil ? `OPEN · đến ${formatUnlockUntil(selInfo.unlockUntil)}` : 'OPEN'}</span>
                  </button>
                )}
              </div>
            </div>

            <ol className="report-day-panel__steps">
              {daySteps.map((st, i) => (
                <li key={st.label} className={`report-day-step${st.value ? ' is-done' : ''}`}>
                  <span className="report-day-step__mark">{st.value ? <Icon name="check" size={16} /> : i + 1}</span>
                  <span className="report-day-step__text">
                    <span className="report-day-step__label">{st.label}</span>
                    <span className="report-day-step__value">{st.value || 'Chưa có'}</span>
                  </span>
                </li>
              ))}
            </ol>

            <div className="report-day-panel__foot">
              <span className="report-day-panel__path">
                <code>{folderLabel}</code>
                <button
                  type="button"
                  className="btn btn--ghost btn--icon btn--sm"
                  title="Sao chép đường dẫn"
                  aria-label="Sao chép đường dẫn thư mục"
                  onClick={() => copyPath(folderLabel)}
                >
                  <Icon name="copy" size={16} />
                </button>
              </span>
              <span className="report-day-panel__hint">
                <Icon name="info" size={16} />
                <span>
                  {selStatus === 'missing' && 'Chưa có báo cáo ngày này. Biểu mẫu mới sẽ điền sẵn từ báo cáo gần nhất (không sao chép ảnh).'}
                  {selStatus === 'unsigned' && 'Báo cáo đã lưu. In PDF, ký tên rồi upload bản đã ký để hoàn tất. '}
                  {selLocked
                    ? selStatus === 'unsigned'
                      ? `Nội dung báo cáo đã khoá (chỉ xem) nhưng bạn vẫn upload được bản đã ký bằng nút "Upload bản ký".${canManageLock ? ' Admin: bấm nút LOCKED để mở khoá nếu cần sửa nội dung.' : ' Cần sửa nội dung thì liên hệ admin để mở khoá.'}`
                      : canManageLock
                      ? 'Báo cáo đã khoá: người dùng chỉ xem được. Bấm nút LOCKED để mở khoá, cho phép chỉnh sửa lại.'
                      : 'Báo cáo đã khoá: chỉ xem, không chỉnh sửa được. Cần sửa thì liên hệ admin để mở khoá.'
                    : selStatus === 'done' && 'Đã có PDF đã ký nên ngày này hoàn tất.'}
                </span>
              </span>
            </div>
          </div>
          )}
        </>
      )}

      {lockConfirm && (
        <ConfirmDialog
          danger={lockConfirm === 'lock'}
          title={lockConfirm === 'lock' ? 'Khoá báo cáo?' : 'Mở khoá báo cáo?'}
          message={
            lockConfirm === 'lock'
              ? `Báo cáo ${role.unit} ngày ${formatDMY(selected)} sẽ chuyển sang chỉ xem: người dùng không sửa/lưu/upload được cho tới khi admin mở khoá.`
              : `Người dùng sẽ chỉnh sửa lại được báo cáo ${role.unit} ngày ${formatDMY(selected)}. Báo cáo tự khoá lại khi có bản đã ký mới được upload.`
          }
          confirmLabel={lockConfirm === 'lock' ? 'Khoá' : 'Mở khoá'}
          onConfirm={() => {
            toggleLock().catch((err) => console.error(err));
          }}
          onCancel={() => setLockConfirm(null)}
        />
      )}

      <input
        ref={fileRef}
        type="file"
        accept="application/pdf,.pdf"
        className="report-signed-file-pick"
        tabIndex={-1}
        aria-label="Chọn file PDF đã ký"
        onChange={(e) => {
          const f = e.target.files && e.target.files[0];
          handleFile(f || undefined).catch((err) => console.error(err));
        }}
      />
    </div>
  );
};
