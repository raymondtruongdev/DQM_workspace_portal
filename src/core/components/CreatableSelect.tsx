import * as React from 'react';
import { Icon } from './Icon';
import { PromptDialog } from './PromptDialog';

export interface CreatableSelectProps {
  value: string;
  options: string[];
  /** `isNew` = true khi người dùng vừa nhập một giá trị chưa có trong danh sách (nơi gọi cần lưu vào danh sách). */
  onChange: (value: string, isNew: boolean) => void;
  id?: string;
  className?: string;
  /** Class cho khung ngoài (bố cục: flex, bề rộng…). */
  wrapperClassName?: string;
  disabled?: boolean;
  /** Nhãn mục cuối danh sách (mặc định "Khác…"). */
  otherLabel?: string;
  /** Tiêu đề hộp thoại khi chọn "Khác…". */
  dialogTitle?: string;
  ariaInvalid?: boolean;
  title?: string;
  /** Dòng gợi ý khi chưa chọn gì (value rỗng). */
  placeholder?: string;
  /** false = ẩn mục "Khác…" (danh sách cố định, chỉ chọn trong danh sách). Mặc định true. */
  allowOther?: boolean;
  /** Nhãn hiển thị thay cho giá trị của từng mục (giá trị lưu vẫn là `options[i]`), ví dụ { 'A1': 'A1 – Mô tả' }. */
  optionLabels?: Record<string, string>;
}

const OTHER = '__other__';

/**
 * Ô chọn (dropdown) có mục cuối "Khác…": chọn mục này mở hộp thoại nhập giá trị mới.
 * Giá trị mới được chọn ngay và báo lại qua `onChange(value, true)` để nơi gọi thêm vào danh sách.
 * Danh sách cố định: `allowOther={false}`; hiển thị nhãn khác giá trị: `optionLabels`.
 * @example <CreatableSelect value={v} options={opts} onChange={(val, isNew) => …} className="input" />
 */
export const CreatableSelect: React.FC<CreatableSelectProps> = ({
  value,
  options,
  onChange,
  id,
  className = 'input',
  wrapperClassName,
  disabled = false,
  otherLabel = 'Khác…',
  dialogTitle = 'Thêm giá trị mới',
  ariaInvalid,
  title,
  placeholder,
  allowOther = true,
  optionLabels
}) => {
  const [open, setOpen] = React.useState(false);
  const list = value && options.indexOf(value) === -1 ? [value].concat(options) : options;

  const handleAdd = (raw: string): void => {
    setOpen(false);
    const existing = options.filter((o) => o.toLowerCase() === raw.toLowerCase())[0];
    if (existing) onChange(existing, false);
    else onChange(raw, true);
  };

  return (
    <>
      <span className={['select', wrapperClassName || ''].filter(Boolean).join(' ')}>
      <select
        id={id}
        className={className}
        value={value}
        disabled={disabled}
        aria-invalid={ariaInvalid || undefined}
        title={title}
        onChange={(e) => {
          if (e.target.value === OTHER) setOpen(true);
          else onChange(e.target.value, false);
        }}
      >
        {placeholder && (
          <option value="" disabled hidden>
            {placeholder}
          </option>
        )}
        {list.map((o) => (
          <option key={o} value={o}>
            {(optionLabels && optionLabels[o]) || o}
          </option>
        ))}
        {allowOther && <option value={OTHER}>{otherLabel}</option>}
      </select>
      <Icon name="chevron-down" size={16} className="select__icon" />
      </span>
      {open && (
        <PromptDialog
          title={dialogTitle}
          label="Giá trị mới (sẽ có trong danh sách chọn của các ngày sau)"
          onSubmit={handleAdd}
          onCancel={() => setOpen(false)}
        />
      )}
    </>
  );
};
