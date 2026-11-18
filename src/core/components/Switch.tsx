import * as React from 'react';

export interface SwitchProps {
  checked: boolean;
  onChange: (checked: boolean) => void;
  /** Nhãn cạnh công tắc, vd "Đang bật" / "Đang tắt". */
  label?: React.ReactNode;
  disabled?: boolean;
  /** Tên truy cập cho trình đọc màn hình khi không có `label`. */
  ariaLabel?: string;
}

/**
 * Công tắc bật/tắt (role="switch").
 * @example <Switch checked={on} onChange={setOn} label={on ? 'Đang bật' : 'Đang tắt'} />
 */
export const Switch: React.FC<SwitchProps> = ({ checked, onChange, label, disabled, ariaLabel }) => (
  <button
    type="button"
    role="switch"
    aria-checked={checked}
    aria-label={ariaLabel}
    className={`switch${checked ? ' switch--on' : ''}`}
    onClick={() => onChange(!checked)}
    disabled={disabled}
  >
    <span className="switch__track"><span className="switch__knob" /></span>
    {label !== undefined && <span className="switch__text">{label}</span>}
  </button>
);
