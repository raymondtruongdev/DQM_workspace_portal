import * as React from 'react';
import { Icon } from './Icon';

export interface TileItem {
  id: string;
  title: string;
  /** Nội dung trong ô icon (vd <Icon name="…" size={26} /> hoặc <svg> thương hiệu). */
  icon: React.ReactNode;
  onClick: () => void;
}

export interface TileGridProps {
  items: TileItem[];
}

/**
 * Lưới ô chức năng (nút xanh đặc: icon + tên + mũi tên), dùng ở trang chủ và các trang menu.
 * @example <TileGrid items={[{ id: 'ui', title: 'UI Kit', icon: <Icon name="paint" size={26} />, onClick: open }]} />
 */
export const TileGrid: React.FC<TileGridProps> = ({ items }) => (
  <div className="sp-tiles-grid">
    {items.map((item) => (
      <button
        key={item.id}
        type="button"
        className="landing-grid-btn"
        onClick={item.onClick}
        title={`Nhấn để truy cập ${item.title}`}
      >
        <div className="landing-grid-btn-icon">{item.icon}</div>
        <span className="landing-grid-btn-text">{item.title}</span>
        <Icon name="chevron-right" size={20} className="landing-grid-btn-arrow" />
      </button>
    ))}
  </div>
);
