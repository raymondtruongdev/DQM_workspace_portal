import * as React from 'react';
import { Icon } from './Icon';

export interface ImageLightboxProps {
  url: string;
  /** Tên ảnh hiện ở thanh trên và làm `alt`. */
  name: string;
  onClose: () => void;
}

const MIN_SCALE = 0.25;
const MAX_SCALE = 8;
const BUTTON_STEP = 1.25;

const clamp = (v: number): number => Math.min(MAX_SCALE, Math.max(MIN_SCALE, v));

/**
 * Khung xem ảnh phóng to NGAY TRONG trang (không điều hướng, nên không rời form đang nhập). Esc hoặc bấm nền để đóng.
 * Thu/phóng: nút +/−, con lăn chuột (phóng theo vị trí con trỏ), nhấp đúp (Vừa khung ↔ 100%); "Vừa" đưa ảnh về vừa khung; kéo chuột để di chuyển khi đã phóng.
 * Dùng thay cho thẻ `<a href=ảnh>`: trên trang SharePoint, liên kết tới file trong thư viện có thể mở ngay tab hiện tại và làm mất dữ liệu chưa lưu.
 * @example {open && <ImageLightbox url={src} name="layout1.png" onClose={() => setOpen(false)} />}
 */
export const ImageLightbox: React.FC<ImageLightboxProps> = ({ url, name, onClose }) => {
  // scale = 1 là vừa khung; tx/ty là độ dịch (px) tính từ tâm khung xem
  const [view, setView] = React.useState({ scale: 1, tx: 0, ty: 0 });
  const [loaded, setLoaded] = React.useState(false);
  const stageRef = React.useRef<HTMLDivElement | null>(null);
  const imgRef = React.useRef<HTMLImageElement | null>(null);
  const drag = React.useRef<{ x: number; y: number; tx: number; ty: number; moved: boolean } | null>(null);
  const justDragged = React.useRef(false);

  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  /** Đổi tỉ lệ sao cho điểm (px, py) (so với tâm khung xem) đứng yên. */
  const zoomAt = React.useCallback((factor: number, px: number, py: number): void => {
    setView((v) => {
      const scale = clamp(v.scale * factor);
      if (scale === v.scale) return v;
      const ratio = scale / v.scale;
      return { scale, tx: px - (px - v.tx) * ratio, ty: py - (py - v.ty) * ratio };
    });
  }, []);

  // Con lăn chuột: cần listener không-thụ-động để chặn cuộn trang phía sau
  React.useEffect(() => {
    const stage = stageRef.current;
    if (!stage) return undefined;
    const onWheel = (e: WheelEvent): void => {
      e.preventDefault();
      const rect = stage.getBoundingClientRect();
      const px = e.clientX - (rect.left + rect.width / 2);
      const py = e.clientY - (rect.top + rect.height / 2);
      zoomAt(Math.pow(1.0015, -e.deltaY), px, py);
    };
    stage.addEventListener('wheel', onWheel, { passive: false });
    return () => stage.removeEventListener('wheel', onWheel);
  }, [zoomAt]);

  const fit = (): void => setView({ scale: 1, tx: 0, ty: 0 });
  /** Tỉ lệ để ảnh hiện đúng kích thước gốc (100%). */
  const actualScale = (): number => {
    const img = imgRef.current;
    return img && img.clientWidth > 0 && img.naturalWidth > 0 ? img.naturalWidth / img.clientWidth : 1;
  };
  const actual = (): void => setView({ scale: clamp(actualScale()), tx: 0, ty: 0 });

  const img = imgRef.current;
  const percent = loaded && img && img.naturalWidth > 0 ? Math.round(((img.clientWidth * view.scale) / img.naturalWidth) * 100) : null;

  const onPointerDown = (e: React.PointerEvent<HTMLImageElement>): void => {
    if (e.button !== 0) return;
    drag.current = { x: e.clientX, y: e.clientY, tx: view.tx, ty: view.ty, moved: false };
    e.currentTarget.setPointerCapture(e.pointerId);
  };
  const onPointerMove = (e: React.PointerEvent<HTMLImageElement>): void => {
    const d = drag.current;
    if (!d) return;
    const dx = e.clientX - d.x;
    const dy = e.clientY - d.y;
    if (!d.moved && Math.abs(dx) + Math.abs(dy) < 4) return;
    d.moved = true;
    setView((v) => ({ ...v, tx: d.tx + dx, ty: d.ty + dy }));
  };
  const onPointerUp = (): void => {
    if (drag.current?.moved) {
      justDragged.current = true;
      window.setTimeout(() => { justDragged.current = false; }, 0);
    }
    drag.current = null;
  };

  return (
    <div className="lightbox" role="dialog" aria-modal="true" aria-label={name} onClick={onClose}>
      <div className="lightbox__bar" onClick={(e) => e.stopPropagation()}>
        <span className="lightbox__name" title={name}>{name}</span>
        <button type="button" className="btn btn--light btn--icon btn--sm" onClick={() => zoomAt(1 / BUTTON_STEP, 0, 0)} aria-label="Thu nhỏ" title="Thu nhỏ">
          <Icon name="zoom-out" size={16} />
        </button>
        <span className="lightbox__zoom" aria-live="polite">{percent === null ? '' : `${percent}%`}</span>
        <button type="button" className="btn btn--light btn--icon btn--sm" onClick={() => zoomAt(BUTTON_STEP, 0, 0)} aria-label="Phóng to" title="Phóng to">
          <Icon name="zoom-in" size={16} />
        </button>
        <button type="button" className="btn btn--light btn--sm" onClick={fit} title="Vừa khung">
          <Icon name="scale-fit" size={16} />
          <span>Vừa</span>
        </button>
        <button type="button" className="btn btn--light btn--sm" onClick={actual} title="Kích thước gốc (100%)">
          <span>100%</span>
        </button>
        <button
          type="button"
          className="btn btn--light btn--sm"
          onClick={() => window.open(url, '_blank', 'noopener,noreferrer')}
          title="Mở ảnh gốc ở tab mới"
        >
          <Icon name="image" size={16} />
          <span>Mở tab mới</span>
        </button>
        <button type="button" className="btn btn--light btn--icon btn--sm" onClick={onClose} aria-label="Đóng" title="Đóng (Esc)">
          <Icon name="dismiss" size={16} />
        </button>
      </div>
      <div
        ref={stageRef}
        className="lightbox__stage"
        onClick={(e) => {
          e.stopPropagation();
          if (!justDragged.current) onClose();
        }}
      >
        <img
          ref={imgRef}
          className={`lightbox__img${view.scale > 1 ? ' lightbox__img--pan' : ''}`}
          src={url}
          alt={name}
          draggable={false}
          style={{ transform: `translate(${view.tx}px, ${view.ty}px) scale(${view.scale})` }}
          onLoad={() => setLoaded(true)}
          onClick={(e) => e.stopPropagation()}
          onDoubleClick={() => (Math.abs(view.scale - 1) < 0.01 ? actual() : fit())}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onPointerCancel={onPointerUp}
        />
      </div>
    </div>
  );
};
