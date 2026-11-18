import * as React from 'react';
import type { PortalConfig, ToastState } from '../../../core/types';
import { Icon, LoadingState, PageHeader, ToastAlert } from '../../../core/components';
import { PortalConfigService } from '../../../core/services';

/**
 * Cài đặt module báo cáo, lưu ở `moduleSettings.report` trong portal_config.json.
 * Khoá và ý nghĩa PHẢI khớp `modules/report/utils/reportSettings.ts` (module không import lẫn nhau).
 */
const SETTINGS_ID = 'report';
const DEFAULT_DAYS = 14;
const MIN_DAYS = 1;
const MAX_DAYS = 60;

export interface DailyReportSetupViewProps {
  configService: PortalConfigService;
  adminName: string;
  onBack: () => void;
}

function readSettings(config: PortalConfig): { todayOnly: boolean; days: number } {
  const raw = config.moduleSettings[SETTINGS_ID] || {};
  const n = Math.round(Number(raw.windowDays));
  return {
    todayOnly: raw.todayOnly === true,
    days: Number.isFinite(n) ? Math.min(MAX_DAYS, Math.max(MIN_DAYS, n)) : DEFAULT_DAYS
  };
}

/** Setup DailyReport: chỉ cho nhập báo cáo ngày hiện tại + số ngày thống kê người dùng thấy. */
export const DailyReportSetupView: React.FC<DailyReportSetupViewProps> = ({ configService, adminName, onBack }) => {
  const [config, setConfig] = React.useState<PortalConfig | undefined>(undefined);
  const [todayOnly, setTodayOnly] = React.useState<boolean>(false);
  const [daysText, setDaysText] = React.useState<string>(String(DEFAULT_DAYS));
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
      })
      .catch((e) => console.error(e));
    return () => {
      alive = false;
    };
  }, [configService]);

  const days = Number(daysText);
  const daysValid = Number.isInteger(days) && days >= MIN_DAYS && days <= MAX_DAYS;
  const saved = config ? readSettings(config) : undefined;
  const dirty = Boolean(saved) && (saved?.todayOnly !== todayOnly || saved?.days !== days);

  const save = async (): Promise<void> => {
    if (!daysValid) return;
    setIsSaving(true);
    try {
      const next = await configService.setModuleSettings(SETTINGS_ID, { todayOnly, windowDays: days }, adminName);
      setConfig(next);
      setToast({ type: 'success', text: 'Đã lưu cài đặt DailyReport.' });
    } catch (err: unknown) {
      console.error('Error saving report settings:', err);
      setToast({ type: 'error', text: err instanceof Error ? err.message : 'Không lưu được cấu hình lên SharePoint' });
    } finally {
      setIsSaving(false);
    }
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
        <div className="card admin-setup-panel">
          <section className="admin-setup-block">
            <h3 className="admin-setup-title">Chỉ nhập báo cáo ngày hiện tại</h3>
            <label className="admin-setup-check">
              <input type="checkbox" checked={todayOnly} onChange={(e) => setTodayOnly(e.target.checked)} disabled={isSaving} />
              <span>Chỉ cho người dùng nhập báo cáo của ngày hôm nay</span>
            </label>
            <p className="admin-setup-hint">
              Khi bật, mọi báo cáo của ngày cũ hơn hôm nay tự khoá: người dùng chỉ xem, không nhập mới, sửa, lưu hay upload bản ký được.
              Nút OPEN / LOCKED của từng báo cáo vẫn dùng bình thường: admin bấm OPEN ở một ngày cũ để cho phép sửa riêng ngày đó
              (báo cáo tự khoá lại khi upload bản ký), bấm LOCKED để khoá riêng một ngày. Ngày tính theo giờ máy người dùng.
            </p>
          </section>

          <section className="admin-setup-block">
            <h3 className="admin-setup-title">Số ngày thống kê</h3>
            <div className="field admin-setup-days">
              <label htmlFor="admin-setup-days" className="field__label">
                Số ngày gần nhất hiện ở lịch và tab Tổng quan ({MIN_DAYS}–{MAX_DAYS})
              </label>
              <input
                id="admin-setup-days"
                type="number"
                className={`input${daysValid ? '' : ' input--danger'}`}
                min={MIN_DAYS}
                max={MAX_DAYS}
                step={1}
                inputMode="numeric"
                value={daysText}
                onChange={(e) => setDaysText(e.target.value)}
                disabled={isSaving}
              />
            </div>
            <p className="admin-setup-hint">Mặc định {DEFAULT_DAYS} ngày. Chỉ đổi phạm vi hiển thị, không xoá hay đổi dữ liệu báo cáo.</p>
          </section>

          <div className="admin-setup-actions">
            <button type="button" className="btn btn--primary" onClick={save} disabled={isSaving || !daysValid || !dirty}>
              <Icon name="save" />
              <span>{isSaving ? 'Đang lưu...' : 'Lưu cài đặt'}</span>
            </button>
            {config.updatedAt && (
              <span className="admin-setup-hint">
                Cập nhật bởi {config.updatedBy || 'admin'} lúc {new Date(config.updatedAt).toLocaleString('vi-VN')}
              </span>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
