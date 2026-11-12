import * as React from 'react';
import type { PortalConfig, ToastState } from '../../../core/types';
import { Icon, LoadingState, PageHeader, ToastAlert } from '../../../core/components';
import { PortalConfigService } from '../../../core/services';

export interface ShutdownViewProps {
  configService: PortalConfigService;
  adminName: string;
  onBack: () => void;
}

/** Shutdown Server: bật/tắt chế độ bảo trì của portal (người dùng thường bị khoá, admin vẫn vào được). */
export const ShutdownView: React.FC<ShutdownViewProps> = ({ configService, adminName, onBack }) => {
  const [config, setConfig] = React.useState<PortalConfig | undefined>(undefined);
  const [message, setMessage] = React.useState<string>('');
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [toast, setToast] = React.useState<ToastState | undefined>(undefined);

  React.useEffect(() => {
    let alive = true;
    configService
      .getConfig()
      .then((c) => {
        if (!alive) return;
        setConfig(c);
        setMessage(c.message);
      })
      .catch((e) => console.error(e));
    return () => {
      alive = false;
    };
  }, [configService]);

  const apply = async (on: boolean): Promise<void> => {
    if (
      on &&
      !window.confirm('Bật chế độ bảo trì? Mọi người dùng không phải admin sẽ bị chặn truy cập portal cho đến khi bạn tắt.')
    ) {
      return;
    }
    setIsSaving(true);
    try {
      const next = await configService.setMaintenance(on, message, adminName);
      setConfig(next);
      setToast({ type: 'success', text: on ? 'Đã bật chế độ bảo trì.' : 'Đã tắt chế độ bảo trì.' });
    } catch (err: unknown) {
      console.error('Error saving portal config:', err);
      const msg = err instanceof Error ? err.message : 'Không lưu được cấu hình lên SharePoint';
      setToast({ type: 'error', text: msg });
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
        title="Shutdown Server"
        meta={<span>Chế độ bảo trì của portal</span>}
      />

      <ToastAlert toast={toast} onClose={() => setToast(undefined)} />

      {!config ? (
        <LoadingState message="Đang tải cấu hình..." />
      ) : (
        <div className="card admin-shutdown-panel">
          <div className="admin-shutdown-status">
            <span className={config.maintenance ? 'badge badge--danger' : 'badge badge--success'}>
              {config.maintenance ? 'Đang bảo trì' : 'Đang hoạt động'}
            </span>
            {config.updatedAt && (
              <span className="admin-shutdown-meta">
                Cập nhật bởi {config.updatedBy || 'admin'} lúc {new Date(config.updatedAt).toLocaleString('vi-VN')}
              </span>
            )}
          </div>

          <p className="admin-shutdown-hint">
            Khi bật, người dùng thường chỉ thấy màn hình bảo trì; tài khoản admin vẫn truy cập bình thường.
          </p>

          <div className="field">
            <label htmlFor="admin-shutdown-message" className="field__label">
              Thông báo hiển thị cho người dùng
            </label>
            <textarea
              id="admin-shutdown-message"
              className="textarea"
              rows={3}
              placeholder="VD: Hệ thống đang bảo trì, vui lòng quay lại sau 15:00."
              value={message}
              onChange={(e) => setMessage(e.target.value)}
              disabled={isSaving}
            />
          </div>

          <div className="admin-shutdown-actions">
            {config.maintenance ? (
              <button type="button" className="btn btn--primary" onClick={() => apply(false)} disabled={isSaving}>
                <Icon name="power" />
                <span>Tắt chế độ bảo trì</span>
              </button>
            ) : (
              <button type="button" className="btn btn--danger" onClick={() => apply(true)} disabled={isSaving}>
                <Icon name="power" />
                <span>Bật chế độ bảo trì</span>
              </button>
            )}
            {config.maintenance && (
              <button type="button" className="btn btn--secondary" onClick={() => apply(true)} disabled={isSaving}>
                <Icon name="save" />
                <span>Cập nhật thông báo</span>
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
