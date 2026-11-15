import * as React from 'react';

export interface ConfirmDialogProps {
  title: string;
  /** Nội dung giải thích hậu quả của thao tác. */
  message: React.ReactNode;
  confirmLabel?: string;
  /** Thao tác phá huỷ: nút xác nhận đỏ. */
  danger?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
  /** true = chỉ một nút (thông báo bắt buộc xác nhận, vd "Trở về Home"): ẩn nút Huỷ, bấm nền / Esc không đóng. */
  hideCancel?: boolean;
}

/**
 * Hộp thoại xác nhận Có/Không (Esc = huỷ). Dùng cho thao tác không hoàn tác được như xoá file.
 * @example {open && <ConfirmDialog danger title="Xoá file?" message="Không thể hoàn tác." confirmLabel="Xoá" onConfirm={remove} onCancel={() => setOpen(false)} />}
 */
export const ConfirmDialog: React.FC<ConfirmDialogProps> = ({
  title,
  message,
  confirmLabel = 'Xác nhận',
  danger = false,
  onConfirm,
  onCancel,
  hideCancel = false
}) => {
  const confirmRef = React.useRef<HTMLButtonElement | null>(null);
  React.useEffect(() => {
    if (confirmRef.current) confirmRef.current.focus();
  }, []);

  return (
    <div className="dialog-backdrop" onClick={hideCancel ? undefined : onCancel}>
      <div
        className="dialog"
        role="alertdialog"
        aria-modal="true"
        aria-label={title}
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Escape' && !hideCancel) onCancel();
        }}
      >
        <div className="dialog__title">{title}</div>
        <div className="dialog__body">{message}</div>
        <div className="dialog__actions">
          {!hideCancel && (
            <button type="button" className="btn btn--secondary" onClick={onCancel}>
              Huỷ
            </button>
          )}
          <button ref={confirmRef} type="button" className={`btn ${danger ? 'btn--danger' : 'btn--primary'}`} onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};
