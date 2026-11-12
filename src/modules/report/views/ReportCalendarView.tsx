import * as React from 'react';
import { RoleEntry, ToastState } from '../../../core/types';
import {
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
  buildRecentDates,
  DAY_STATUS_BADGE,
  DAY_STATUS_LABEL,
  formatDM,
  formatDMY,
  formatFileSize,
  formatSavedAt,
  mondayIndex,
  parseDateStr,
  WEEKDAY_SHORT_MON_FIRST,
  weekdayLong
} from '../utils/reportCalendar';
import { getReportFolderName } from '../utils/reportFileNaming';

const WINDOW_DAYS = 14;

export interface IReportCalendarViewProps {
  role: RoleEntry;
  dataService: ReportDataService;
  /** true = portal đang bảo trì, không được ghi lên SharePoint. */
  onCheckMaintenance: () => Promise<boolean>;
  onBack: () => void;
  /** Mở form nhập báo cáo của một ngày (YYYY-MM-DD). */
  onOpenDate: (dateStr: string) => void;
}

type CalendarMode = 'cal' | 'list';

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
  onOpenDate
}) => {
  const dates = React.useMemo(() => buildRecentDates(WINDOW_DAYS), []);
  const todayStr = dates[dates.length - 1];

  const [infos, setInfos] = React.useState<ReportDayInfo[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isUploading, setIsUploading] = React.useState<boolean>(false);
  const [mode, setMode] = React.useState<CalendarMode>('cal');
  const [selected, setSelected] = React.useState<string>(todayStr);
  const [toast, setToast] = React.useState<ToastState | null>(null);
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

  const byDate = React.useMemo(() => {
    const m: Record<string, ReportDayInfo> = {};
    infos.forEach((i) => {
      m[i.dateStr] = i;
    });
    return m;
  }, [infos]);

  const statusOf = (dateStr: string): ReportDayStatus => (byDate[dateStr] ? byDate[dateStr].status : 'missing');

  const counts = { done: 0, unsigned: 0, missing: 0 };
  dates.forEach((d) => {
    counts[statusOf(d)] += 1;
  });
  const missingDates = dates.filter((d) => statusOf(d) === 'missing');
  const unsignedDates = dates.filter((d) => statusOf(d) === 'unsigned');

  const selInfo = byDate[selected];
  const selStatus = statusOf(selected);
  // Sổ ReportRegistry không lưu tên file đã ký (<DocCode>_signed.pdf): đọc thư mục khi chọn ngày đã hoàn tất
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
  const folderLabel = `data/${role.unit}/${getReportFolderName(role.docId, selected)}/`;

  const pickFile = (dateStr: string): void => {
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
      await dataService.uploadSignedDocument(role.unit, role.docId, target, buf);
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
    return (
      <button
        key={dateStr}
        type="button"
        className={`report-cal-day report-cal-day--${st}${dateStr === todayStr ? ' is-today' : ''}`}
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
          {DAY_STATUS_LABEL[st]}
        </span>
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
        <LoadingState message="Đang kiểm tra các báo cáo của 14 ngày gần nhất..." />
      ) : (
        <>
          <section className="card report-cal-summary" aria-label="Tóm tắt 14 ngày">
            <div className="report-cal-summary__big">
              <b>{counts.done}</b>
              <span>/ {WINDOW_DAYS} ngày hoàn tất</span>
            </div>
            <div className="report-cal-strip" role="img" aria-label="Trạng thái 14 ngày">
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
                {WINDOW_DAYS} NGÀY GẦN NHẤT <span className="section-card__title-sub">| LAST {WINDOW_DAYS} DAYS</span>
              </span>
              <FilterChips
                ariaLabel="Chế độ xem"
                value={mode}
                onChange={(v) => setMode(v as CalendarMode)}
                options={[
                  { value: 'cal', label: 'Lịch' },
                  { value: 'list', label: 'Danh sách' }
                ]}
              />
            </div>
            <div className="section-card__body">
              {mode === 'cal' ? (
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

              <div className="report-cal-legend">
                <span><i className="report-cal-legend__dot report-cal-legend__dot--missing" />Chưa nhập: chưa bấm Lưu</span>
                <span><i className="report-cal-legend__dot report-cal-legend__dot--unsigned" />Chưa upload bản ký: đã lưu, thiếu file PDF đã ký</span>
                <span><i className="report-cal-legend__dot report-cal-legend__dot--done" />Hoàn tất: có file PDF trong thư mục báo cáo</span>
              </div>
            </div>
          </div>

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
                  {selStatus === 'unsigned' && 'Báo cáo đã lưu. In PDF, ký tên rồi upload bản đã ký để hoàn tất.'}
                  {selStatus === 'done' && 'Đã có PDF đã ký nên ngày này hoàn tất.'}
                </span>
              </span>
            </div>
          </div>
        </>
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
