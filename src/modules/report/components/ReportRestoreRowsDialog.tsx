import * as React from 'react';
import { RestorableRow } from '../utils/reportRestore';

export interface IReportRestoreRowsDialogProps {
  /** Tiêu đề mục (vd "VII. NHÂN LỰC") và ngày báo cáo nguồn (dd/MM/yyyy). */
  sectionTitle: string;
  sourceDate: string;
  /** Dòng ngày trước đang thiếu; rỗng = không có gì để khôi phục. */
  items: RestorableRow[];
  onConfirm: (ids: string[]) => void;
  onCancel: () => void;
}

/** Hộp thoại chọn các dòng của ngày trước đã bị xoá để khôi phục lại vào bảng. */
export const ReportRestoreRowsDialog: React.FC<IReportRestoreRowsDialogProps> = ({ sectionTitle, sourceDate, items, onConfirm, onCancel }) => {
  const [picked, setPicked] = React.useState<string[]>(() => items.map((i) => i.id));
  const allPicked = picked.length === items.length;
  const toggle = (id: string): void => setPicked((p) => (p.indexOf(id) !== -1 ? p.filter((x) => x !== id) : [...p, id]));

  return (
    <div className="dialog-backdrop" onClick={onCancel}>
      <div
        className="dialog report-restore-dialog"
        role="dialog"
        aria-modal="true"
        aria-label="Khôi phục dòng từ ngày trước"
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onCancel();
        }}
      >
        <div className="dialog__title">Khôi phục dòng từ ngày trước</div>
        <div className="dialog__body">
          {items.length === 0 ? (
            <p>Không có dòng nào bị thiếu so với báo cáo ngày {sourceDate} ở mục {sectionTitle}.</p>
          ) : (
            <>
              <p>
                Các dòng của báo cáo ngày {sourceDate} (mục {sectionTitle}) hiện không còn trong báo cáo này. Chọn dòng cần khôi phục; giá trị được tính lại như khi mở ngày mới (luỹ kế cộng dồn, số thực hiện về 0).
              </p>
              <label className="report-restore-dialog__all">
                <input type="checkbox" checked={allPicked} onChange={() => setPicked(allPicked ? [] : items.map((i) => i.id))} />
                <span>Chọn tất cả ({items.length})</span>
              </label>
              <ul className="report-restore-dialog__list">
                {items.map((it) => (
                  <li key={it.id}>
                    <label>
                      <input type="checkbox" checked={picked.indexOf(it.id) !== -1} onChange={() => toggle(it.id)} />
                      <span title={it.label}>{it.label}</span>
                    </label>
                  </li>
                ))}
              </ul>
            </>
          )}
        </div>
        <div className="dialog__actions">
          <button type="button" className="btn btn--secondary" onClick={onCancel}>
            {items.length === 0 ? 'Đóng' : 'Huỷ'}
          </button>
          {items.length > 0 && (
            <button type="button" className="btn btn--primary" onClick={() => onConfirm(picked)} disabled={picked.length === 0}>
              Khôi phục {picked.length} dòng
            </button>
          )}
        </div>
      </div>
    </div>
  );
};
