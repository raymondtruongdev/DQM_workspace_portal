import * as React from 'react';
import { Icon } from './Icon';

export interface NumberStepperProps {
  id?: string;
  /** Giá trị đang nhập dạng chữ (cho phép tạm thời trống/không hợp lệ khi gõ). */
  text: string;
  onChange: (text: string) => void;
  min: number;
  max: number;
  /** Đơn vị hiện cạnh số, vd "ngày", "giờ". */
  unit?: string;
  /** false → viền đỏ (giá trị ngoài khoảng/không phải số nguyên). */
  valid?: boolean;
  disabled?: boolean;
}

/**
 * Ô số nguyên có nút − / + và đơn vị; nút tự giới hạn trong [min, max].
 * @example <NumberStepper text={t} onChange={setT} min={1} max={60} unit="ngày" valid={ok} />
 */
export const NumberStepper: React.FC<NumberStepperProps> = ({ id, text, onChange, min, max, unit, valid = true, disabled }) => {
  const n = Number(text);
  const step = (d: number): void => {
    const base = Number.isFinite(n) && text.trim() !== '' ? Math.round(n) + d : min;
    onChange(String(Math.min(max, Math.max(min, base))));
  };
  return (
    <div className={`stepper${valid ? '' : ' stepper--invalid'}`}>
      <button type="button" className="stepper__step" aria-label="Giảm" onClick={() => step(-1)} disabled={disabled || (valid && n <= min)}>
        <Icon name="subtract" size={16} />
      </button>
      <input
        id={id}
        type="number"
        className="stepper__value"
        min={min}
        max={max}
        step={1}
        inputMode="numeric"
        value={text}
        onChange={(e) => onChange(e.target.value)}
        disabled={disabled}
      />
      {unit && <span className="stepper__unit">{unit}</span>}
      <button type="button" className="stepper__step" aria-label="Tăng" onClick={() => step(1)} disabled={disabled || (valid && n >= max)}>
        <Icon name="add" size={16} />
      </button>
    </div>
  );
};
