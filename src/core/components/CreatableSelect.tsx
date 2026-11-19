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
  /** Nhãn mục cuối danh sách (mặc định "Thêm mới…"). */
  otherLabel?: string;
  /** Tiêu đề hộp thoại khi chọn "Thêm mới…". */
  dialogTitle?: string;
  ariaInvalid?: boolean;
  title?: string;
  /** Dòng gợi ý khi chưa chọn gì (value rỗng). */
  placeholder?: string;
  /** false = ẩn mục "Thêm mới…" (danh sách cố định, chỉ chọn trong danh sách). Mặc định true. */
  allowOther?: boolean;
  /** Nhãn hiển thị thay cho giá trị của từng mục (giá trị lưu vẫn là `options[i]`), ví dụ { 'A1': 'A1 – Mô tả' }. */
  optionLabels?: Record<string, string>;
  /**
   * true = nhãn dài (`optionLabels`) chỉ hiện trong danh sách xổ ra; ô đã chọn chỉ hiện giá trị lưu.
   * Ô `<select>` gốc luôn hiện nhãn của mục đang chọn nên chữ của nó được làm trong suốt và phủ lên một dòng chữ chỉ có giá trị.
   */
  compactSelected?: boolean;
  /** true (mặc định) = sắp xếp các mục tăng dần (so tự nhiên: A2 trước A10), "---" lên đầu, "NA" (nếu có) xuống cuối các mục, "Thêm mới…" luôn ở cuối cùng. false = giữ đúng thứ tự `options`. */
  sorted?: boolean;
}

/** Mục "chưa chọn" luôn đứng đầu danh sách. */
const NONE_VALUE = '---';

/** Mục "không áp dụng" (nếu danh mục có) luôn đứng cuối danh sách mục, ngay trước "Thêm mới…". */
const NOT_APPLICABLE_VALUE = 'NA';

/** Sắp xếp tăng dần theo giá trị (so tự nhiên, không phân biệt hoa thường/dấu), "---" lên đầu, "NA" xuống cuối. */
export function sortSelectOptions(options: string[]): string[] {
  return options.slice().sort((a, b) => {
    if (a === NONE_VALUE) return b === NONE_VALUE ? 0 : -1;
    if (b === NONE_VALUE) return 1;
    if (a === NOT_APPLICABLE_VALUE) return b === NOT_APPLICABLE_VALUE ? 0 : 1;
    if (b === NOT_APPLICABLE_VALUE) return -1;
    return a.localeCompare(b, 'vi', { numeric: true, sensitivity: 'base' });
  });
}

const OTHER = '__other__';

/**
 * Báo cho các CreatableSelect bên trong biết nơi cần làm mới danh sách chọn: được gọi khi người dùng bấm/focus vào ô
 * (nơi cung cấp tự giới hạn tần suất). Không có Provider thì không làm gì.
 * @example <SelectOpenContext.Provider value={refreshOptions}>…</SelectOpenContext.Provider>
 */
export const SelectOpenContext = React.createContext<(() => void) | undefined>(undefined);

/**
 * Ô chọn (dropdown) có mục cuối "Thêm mới…": chọn mục này mở hộp thoại nhập giá trị mới.
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
  otherLabel = 'Thêm mới…',
  dialogTitle = 'Thêm giá trị mới',
  ariaInvalid,
  title,
  placeholder,
  allowOther = true,
  optionLabels,
  compactSelected = false,
  sorted = true
}) => {
  const [open, setOpen] = React.useState(false);
  const compact = compactSelected && Boolean(value) && Boolean(optionLabels && optionLabels[value] && optionLabels[value] !== value);
  const alignMatch = /\btext-(center|right)\b/.exec(className);
  const onSelectOpen = React.useContext(SelectOpenContext);
  const withValue = value && options.indexOf(value) === -1 ? [value].concat(options) : options;
  const list = React.useMemo(() => (sorted ? sortSelectOptions(withValue) : withValue), [sorted, withValue]);

  const handleAdd = (raw: string): void => {
    setOpen(false);
    const existing = options.filter((o) => o.toLowerCase() === raw.toLowerCase())[0];
    if (existing) onChange(existing, false);
    else onChange(raw, true);
  };

  return (
    <>
      <span className={['select', compact ? 'select--compact' : '', wrapperClassName || ''].filter(Boolean).join(' ')}>
      <select
        id={id}
        className={className}
        value={value}
        disabled={disabled}
        aria-invalid={ariaInvalid || undefined}
        title={title}
        onMouseDown={() => onSelectOpen?.()}
        onFocus={() => onSelectOpen?.()}
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
      {compact && (
        <span
          className={`select__shown${disabled ? ' select__shown--disabled' : ''}`}
          style={alignMatch ? { justifyContent: alignMatch[1] === 'center' ? 'center' : 'flex-end' } : undefined}
          aria-hidden="true"
        >
          <span className="select__shown-text">{value}</span>
        </span>
      )}
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
