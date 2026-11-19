import * as React from 'react';
import { Icon } from './Icon';

export interface MoveButtonsProps {
  /** Vị trí hiện tại (bắt đầu từ 0) và tổng số phần tử: nút lên tắt ở đầu danh sách, nút xuống tắt ở cuối. */
  index: number;
  count: number;
  /** `-1` = đưa lên, `1` = đưa xuống. */
  onMove: (delta: -1 | 1) => void;
  disabled?: boolean;
  /** Tên mục dùng cho tooltip/aria, vd "dòng" → "Đưa dòng lên". */
  itemLabel?: string;
  /** Mặc định `secondary` (viền + mũi tên xanh, dễ thấy trong bảng); `light` khi đặt trên banner xanh đậm (cùng quy ước `btn--light`). */
  variant?: 'secondary' | 'light';
}

/**
 * Cặp nút mũi tên đưa một dòng/mục lên hoặc xuống một vị trí.
 * @example <MoveButtons index={i} count={rows.length} onMove={(d) => setRows(moveItem(rows, i, d))} />
 */
export const MoveButtons: React.FC<MoveButtonsProps> = ({ index, count, onMove, disabled, itemLabel = 'dòng', variant = 'secondary' }) => (
  <span className="move-buttons">
    <button
      type="button"
      className={`btn btn--${variant} btn--icon btn--sm`}
      onClick={() => onMove(-1)}
      disabled={disabled || index <= 0}
      title={`Đưa ${itemLabel} lên`}
      aria-label={`Đưa ${itemLabel} lên`}
    >
      <Icon name="arrow-up" size={16} />
    </button>
    <button
      type="button"
      className={`btn btn--${variant} btn--icon btn--sm`}
      onClick={() => onMove(1)}
      disabled={disabled || index >= count - 1}
      title={`Đưa ${itemLabel} xuống`}
      aria-label={`Đưa ${itemLabel} xuống`}
    >
      <Icon name="arrow-down" size={16} />
    </button>
  </span>
);
