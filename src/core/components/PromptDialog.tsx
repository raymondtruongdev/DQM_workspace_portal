import * as React from 'react';

export interface PromptDialogProps {
  title: string;
  label?: string;
  placeholder?: string;
  initialValue?: string;
  confirmLabel?: string;
  onSubmit: (value: string) => void;
  onCancel: () => void;
}

/**
 * Hộp thoại nhập một giá trị văn bản (Enter = xác nhận, Esc = huỷ). Không cho xác nhận khi để trống.
 * @example {open && <PromptDialog title="Thêm loại mới" onSubmit={add} onCancel={() => setOpen(false)} />}
 */
export const PromptDialog: React.FC<PromptDialogProps> = ({
  title,
  label,
  placeholder,
  initialValue = '',
  confirmLabel = 'Thêm',
  onSubmit,
  onCancel
}) => {
  const [value, setValue] = React.useState(initialValue);
  const inputRef = React.useRef<HTMLInputElement | null>(null);
  const trimmed = value.trim();

  React.useEffect(() => {
    if (inputRef.current) inputRef.current.focus();
  }, []);

  const submit = (): void => {
    if (trimmed) onSubmit(trimmed);
  };

  return (
    <div className="dialog-backdrop" onClick={onCancel}>
      <div
        className="dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onCancel();
        }}
      >
        <div className="dialog__title">{title}</div>
        <div className="dialog__body">
          <div className="field">
            {label && <label className="field__label">{label}</label>}
            <input
              ref={inputRef}
              type="text"
              className="input"
              placeholder={placeholder}
              value={value}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  submit();
                }
              }}
            />
          </div>
        </div>
        <div className="dialog__actions">
          <button type="button" className="btn btn--secondary" onClick={onCancel}>
            Huỷ
          </button>
          <button type="button" className="btn btn--primary" onClick={submit} disabled={!trimmed}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};
