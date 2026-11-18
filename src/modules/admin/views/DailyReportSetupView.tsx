import * as React from 'react';
import type { PortalConfig, ToastState } from '../../../core/types';
import { Icon, LoadingState, NumberStepper, PageHeader, Switch, ToastAlert } from '../../../core/components';
import { PortalConfigService } from '../../../core/services';

/**
 * Cài đặt module báo cáo, lưu ở `moduleSettings.report` trong portal_config.json.
 * Khoá và ý nghĩa PHẢI khớp `modules/report/utils/reportSettings.ts` (module không import lẫn nhau).
 */
const SETTINGS_ID = 'report';
const DEFAULT_DAYS = 14;
const MIN_DAYS = 1;
const MAX_DAYS = 60;
const DEFAULT_UNLOCK = 3;
const MIN_UNLOCK = 1;
const MAX_UNLOCK = 72;

export interface DailyReportSetupViewProps {
  configService: PortalConfigService;
  adminName: string;
  onBack: () => void;
}

function readSettings(config: PortalConfig): { todayOnly: boolean; days: number; unlock: number } {
  const raw = config.moduleSettings[SETTINGS_ID] || {};
  const n = Math.round(Number(raw.windowDays));
  const u = Math.round(Number(raw.unlockHours));
  return {
    unlock: Number.isFinite(u) ? Math.min(MAX_UNLOCK, Math.max(MIN_UNLOCK, u)) : DEFAULT_UNLOCK,
    todayOnly: raw.todayOnly === true,
    days: Number.isFinite(n) ? Math.min(MAX_DAYS, Math.max(MIN_DAYS, n)) : DEFAULT_DAYS
  };
}

/** Đồng hồ ở banner (giờ máy người dùng, chỉ để xem; khoá báo cáo vẫn tính theo giờ máy chủ). */
const BannerClock: React.FC = () => {
  const [now, setNow] = React.useState<Date>(() => new Date());
  React.useEffect(() => {
    const t = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(t);
  }, []);
  const p = (n: number): string => (n < 10 ? '0' : '') + String(n);
  const text = `${p(now.getDate())}/${p(now.getMonth() + 1)}/${now.getFullYear()} - ${p(now.getHours())}:${p(now.getMinutes())}:${p(now.getSeconds())}`;
  return (
    <span className="admin-setup-clock" title="Giờ máy của bạn">
      <Icon name="clock" size={16} />
      <span>{text}</span>
    </span>
  );
};

/** Setup DailyReport: chỉ cho nhập báo cáo ngày hiện tại + số ngày thống kê người dùng thấy. */
export const DailyReportSetupView: React.FC<DailyReportSetupViewProps> = ({ configService, adminName, onBack }) => {
  const [config, setConfig] = React.useState<PortalConfig | undefined>(undefined);
  const [todayOnly, setTodayOnly] = React.useState<boolean>(false);
  const [daysText, setDaysText] = React.useState<string>(String(DEFAULT_DAYS));
  const [unlockText, setUnlockText] = React.useState<string>(String(DEFAULT_UNLOCK));
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [toast, setToast] = React.useState<ToastState | undefined>(undefined);

  React.useEffect(() => {
    let alive = true;
    configService
      .getConfig()
      .then((c) => {
        if (!alive) return;
        const s = readSettings(c);
        setConfig(c);
        setTodayOnly(s.todayOnly);
        setDaysText(String(s.days));
        setUnlockText(String(s.unlock));
      })
      .catch((e) => console.error(e));
    return () => {
      alive = false;
    };
  }, [configService]);

  const days = Number(daysText);
  const daysValid = Number.isInteger(days) && days >= MIN_DAYS && days <= MAX_DAYS;
  const unlock = Number(unlockText);
  const unlockValid = Number.isInteger(unlock) && unlock >= MIN_UNLOCK && unlock <= MAX_UNLOCK;
  const saved = config ? readSettings(config) : undefined;
  const dirty = Boolean(saved) && (saved?.todayOnly !== todayOnly || saved?.days !== days || saved?.unlock !== unlock);

  const save = async (): Promise<void> => {
    if (!daysValid || !unlockValid) return;
    setIsSaving(true);
    try {
      const next = await configService.setModuleSettings(SETTINGS_ID, { todayOnly, windowDays: days, unlockHours: unlock }, adminName);
      setConfig(next);
      setToast({ type: 'success', text: 'Đã lưu cài đặt DailyReport.' });
    } catch (err: unknown) {
      console.error('Error saving report settings:', err);
      setToast({ type: 'error', text: err instanceof Error ? err.message : 'Không lưu được cấu hình lên SharePoint' });
    } finally {
      setIsSaving(false);
    }
  };

  const resetChanges = (): void => {
    if (!saved) return;
    setTodayOnly(saved.todayOnly);
    setDaysText(String(saved.days));
    setUnlockText(String(saved.unlock));
  };

  return (
    <div className="admin-module-root">
      <PageHeader
        leading={
          <button type="button" className="btn btn--secondary" onClick={onBack}>
            <Icon name="arrow-left" />
            <span>Quay lại</span>
          </button>
        }
        title="Setup DailyReport"
        meta={<span>Cài đặt báo cáo ngày cho người dùng</span>}
      />

      <ToastAlert toast={toast} onClose={() => setToast(undefined)} />

      {!config ? (
        <LoadingState message="Đang tải cấu hình..." />
      ) : (
        <div className="section-card admin-setup-panel">
          <div className="section-card__banner">
            <span className="section-card__title">
              CÀI ĐẶT BÁO CÁO <span className="section-card__title-sub">| REPORT SETTINGS</span>
            </span>
            <BannerClock />
          </div>

          <div className="admin-setup-body">
            <section className="admin-setup-group">
              <h3 className="admin-setup-group__label">Nhập liệu</h3>
              <div className="admin-setup-row">
                <div className="admin-setup-row__text">
                  <h4 className="admin-setup-title">Chỉ nhập báo cáo ngày hiện tại</h4>
                  <p className="admin-setup-hint">
                    Báo cáo của các ngày trước hôm nay tự khoá: người dùng chỉ xem, không nhập mới, sửa, lưu hay upload bản ký.
                    Ngày và giờ tính theo giờ máy chủ SharePoint (UTC+7), báo cáo hôm qua khoá lúc 0:00.
                  </p>
                  <p className="admin-setup-legend">
                    <span className="admin-setup-tag admin-setup-tag--open"><Icon name="lock-open" size={16} /></span>
                    <span><strong>OPEN</strong> — admin mở riêng một ngày cũ để cho sửa; tự khoá lại khi upload bản ký.</span>
                  </p>
                  <p className="admin-setup-legend">
                    <span className="admin-setup-tag admin-setup-tag--locked"><Icon name="lock" size={16} /></span>
                    <span><strong>LOCKED</strong> — khoá riêng một ngày bất kỳ.</span>
                  </p>
                </div>
                <Switch checked={todayOnly} onChange={setTodayOnly} disabled={isSaving} label={todayOnly ? 'Đang bật' : 'Đang tắt'} />
              </div>
            </section>

            <section className="admin-setup-group">
              <h3 className="admin-setup-group__label">Hiển thị</h3>
              <div className="admin-setup-row">
                <div className="admin-setup-row__text">
                  <h4 className="admin-setup-title">Số ngày thống kê</h4>
                  <p className="admin-setup-hint">
                    Số ngày gần nhất hiện ở lịch và tab Tổng quan. Chỉ đổi phạm vi hiển thị, không xoá hay đổi dữ liệu báo cáo.
                  </p>
                </div>
                <div className="admin-setup-row__control">
                  <NumberStepper id="admin-setup-days" text={daysText} unit="ngày" min={MIN_DAYS} max={MAX_DAYS} valid={daysValid} disabled={isSaving} onChange={setDaysText} />
                  <span className="admin-setup-range">{MIN_DAYS}–{MAX_DAYS} · mặc định {DEFAULT_DAYS}</span>
                </div>
              </div>
            </section>

            <section className="admin-setup-group">
              <h3 className="admin-setup-group__label">Khoá báo cáo</h3>
              <div className="admin-setup-row">
                <div className="admin-setup-row__text">
                  <h4 className="admin-setup-title">Tự khoá lại sau khi mở khoá</h4>
                  <p className="admin-setup-hint">
                    Khi admin bấm OPEN một báo cáo, hết thời gian này báo cáo quay về quy tắc mặc định (đã ký, hoặc ngày cũ khi bật chế độ trên, thì khoá).
                    Người đang nhập dở cần Lưu trước giờ đó.
                  </p>
                  <p className="admin-setup-note">
                    <Icon name="warning" size={16} />
                    <span>
                      Cần cột <code>UnlockUntil</code> (Text) trong list <strong>DRPRegistry</strong>. Chưa có cột thì mở khoá không có hạn.
                    </span>
                  </p>
                </div>
                <div className="admin-setup-row__control">
                  <NumberStepper id="admin-setup-unlock" text={unlockText} unit="giờ" min={MIN_UNLOCK} max={MAX_UNLOCK} valid={unlockValid} disabled={isSaving} onChange={setUnlockText} />
                  <span className="admin-setup-range">{MIN_UNLOCK}–{MAX_UNLOCK} · mặc định {DEFAULT_UNLOCK}</span>
                </div>
              </div>
            </section>
          </div>

          <div className="admin-setup-foot">
            <span className="admin-setup-foot__meta">
              {config.updatedAt ? (
                <>
                  <span className="admin-setup-foot__dot" />
                  <span>
                    Cập nhật bởi <strong>{config.updatedBy || 'admin'}</strong>
                    <span className="admin-setup-foot__time"> lúc {new Date(config.updatedAt).toLocaleString('vi-VN')}</span>
                  </span>
                </>
              ) : (
                <span>Chưa có cập nhật nào</span>
              )}
            </span>
            <div className="admin-setup-foot__actions">
              <button type="button" className="btn btn--secondary" onClick={resetChanges} disabled={isSaving || !dirty}>
                <span>Huỷ thay đổi</span>
              </button>
              <button type="button" className="btn btn--primary" onClick={save} disabled={isSaving || !daysValid || !unlockValid || !dirty}>
                <Icon name="save" />
                <span>{isSaving ? 'Đang lưu...' : 'Lưu cài đặt'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
