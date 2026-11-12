import * as React from 'react';
import type { ToastState } from '../../../core/types';
import {
  CreatableSelect,
  EmptyState,
  FilterChips,
  ICON_NAMES,
  Icon,
  LoadingState,
  PageHeader,
  PromptDialog,
  SearchInput,
  TileGrid,
  ToastAlert
} from '../../../core/components';

const COLOR_TOKENS = [
  '--color-accent', '--color-accent-hover', '--color-accent-subtle', '--color-banner',
  '--color-text-primary', '--color-text-secondary', '--color-text-muted',
  '--color-surface', '--color-canvas', '--color-border',
  '--color-success', '--color-warning', '--color-danger'
];

const SectionBlock: React.FC<{ title: string; hint?: React.ReactNode; children: React.ReactNode }> = ({
  title,
  hint,
  children
}) => (
  <section className="card admin-guide-section">
    <h3 className="admin-guide-section-title">{title}</h3>
    {hint && <p className="admin-guide-hint">{hint}</p>}
    {children}
  </section>
);

export interface StyleGuideViewProps {
  onBack: () => void;
}

/** Trang UI Kit: liệt kê mọi thành phần dùng chung để module mới đối chiếu (chỉ admin). */
export const StyleGuideView: React.FC<StyleGuideViewProps> = ({ onBack }) => {
  const [chip, setChip] = React.useState<string>('ALL');
  const [query, setQuery] = React.useState<string>('');
  const [kind, setKind] = React.useState<string>('NGHIEM_THU');
  const [kinds, setKinds] = React.useState<string[]>(['KIEM_TRA_ATLD', 'NGHIEM_THU', 'CHI_THI_TVGS']);
  const [promptOpen, setPromptOpen] = React.useState<boolean>(false);
  const [toast, setToast] = React.useState<ToastState | undefined>({ type: 'success', text: 'Đã lưu thành công' });

  return (
    <div className="admin-guide-root">
      <PageHeader
        leading={
          <button type="button" className="btn btn--secondary" onClick={onBack}>
            <Icon name="arrow-left" />
            <span>Quay lại</span>
          </button>
        }
        title="UI Kit"
        badge={<span className="badge badge--info">core</span>}
        meta={<span>Danh mục thành phần dùng chung – xem AGENTS.md mục 8</span>}
      />

      <SectionBlock title="Màu (token)" hint={<>Chỉ dùng <code>var(--color-*)</code>; không viết hex.</>}>
        <div className="admin-guide-grid">
          {COLOR_TOKENS.map((t) => (
            <div key={t} className="admin-guide-cell">
              <div className="admin-guide-swatch" style={{ background: `var(${t})` }} />
              <span className="admin-guide-code">{t}</span>
            </div>
          ))}
        </div>
      </SectionBlock>

      <SectionBlock title="Button" hint={<>Một tông xanh. <code>primary</code> cho hành động chính, <code>secondary</code> cho mọi nút còn lại.</>}>
        <div className="admin-guide-row">
          <button type="button" className="btn btn--primary"><Icon name="save" /><span>primary</span></button>
          <button type="button" className="btn btn--secondary"><Icon name="refresh" /><span>secondary</span></button>
          <button type="button" className="btn btn--ghost"><span>ghost</span></button>
          <button type="button" className="btn btn--danger"><span>danger</span></button>
          <button type="button" className="btn btn--primary" disabled><span>disabled</span></button>
        </div>
        <div className="admin-guide-row">
          <button type="button" className="btn btn--secondary btn--sm"><span>btn--sm</span></button>
          <button type="button" className="btn btn--secondary btn--icon" aria-label="Trang chủ"><Icon name="home" /></button>
          <button type="button" className="btn btn--ghost-danger btn--icon btn--sm" aria-label="Xoá"><Icon name="dismiss" size={16} /></button>
        </div>
        <div className="admin-guide-row" style={{ background: 'var(--color-banner)', padding: 'var(--space-3)', borderRadius: 'var(--radius-md)' }}>
          <button type="button" className="btn btn--light btn--sm"><span>btn--light (trên banner)</span></button>
        </div>
      </SectionBlock>

      <SectionBlock title="Form" hint={<><code>.input</code> / <code>.textarea</code> trong <code>.field</code>. Cao theo <code>--control-h</code>.</>}>
        <div className="admin-guide-form-grid">
          <div className="field">
            <label className="field__label field__label--bilingual">
              <span className="field__label-main">Vị trí / Trạm</span>
              <span className="field__label-sub">Location</span>
            </label>
            <input className="input" defaultValue="ST01 – Ga Bến Thành" />
          </div>
          <div className="field">
            <label className="field__label">Chỉ đọc (input--readonly)</label>
            <input className="input input--readonly" readOnly defaultValue="truongquoctoan1@minhtek.com" />
          </div>
          <div className="field">
            <label className="field__label">Cố định theo mẫu (input--locked)</label>
            <input className="input input--locked" readOnly defaultValue="TCJV" />
          </div>
          <div className="field">
            <label className="field__label">Bị vô hiệu hoá</label>
            <input className="input" disabled defaultValue="disabled" />
          </div>
          <div className="field">
            <label className="field__label">Mức độ cao (input--danger)</label>
            <input className="input input--cell input--danger" defaultValue="Cao" />
          </div>
          <div className="field">
            <label className="field__label">Trung bình (input--warning)</label>
            <input className="input input--cell input--warning" defaultValue="Trung bình" />
          </div>
        </div>
        <div className="field">
          <label className="field__label">Textarea</label>
          <textarea className="textarea" rows={3} placeholder="Nhập nội dung…" />
        </div>
      </SectionBlock>

      <SectionBlock title="Badge" hint={<><code>badge</code> + <code>badge--info | success | warning | danger</code></>}>
        <div className="admin-guide-row">
          <span className="badge">mặc định</span>
          <span className="badge badge--info">badge--info</span>
          <span className="badge badge--success"><Icon name="check-circle" size={16} />badge--success</span>
          <span className="badge badge--warning">badge--warning</span>
          <span className="badge badge--danger">badge--danger</span>
        </div>
      </SectionBlock>

      <SectionBlock title="Card & Section card" hint={<><code>.card</code> và <code>.section-card</code> (banner xanh đậm + body)</>}>
        <div className="section-card">
          <div className="section-card__banner">
            <span className="section-card__title">
              II. NHÂN LỰC <span className="section-card__title-sub">| MANPOWER</span>
            </span>
            <button type="button" className="btn btn--light btn--sm">+ Thêm dòng</button>
          </div>
          <div className="section-card__body">Nội dung của mục.</div>
        </div>
      </SectionBlock>

      <SectionBlock title="PageHeader · FilterChips · SearchInput">
        <PageHeader
          title="Quản lý Báo cáo thi công"
          badge={<span className="badge badge--info">ST01</span>}
          actions={<button type="button" className="btn btn--secondary"><Icon name="refresh" /><span>Làm mới</span></button>}
        />
        <div className="admin-guide-row">
          <FilterChips
            options={[{ value: 'ALL', label: 'Tất cả (3)' }, { value: 'ST01', label: 'ST01 (2)' }, { value: 'ST02', label: 'ST02 (1)' }]}
            value={chip}
            onChange={setChip}
          />
          <SearchInput value={query} onChange={setQuery} placeholder="Tìm kiếm…" />
        </div>
      </SectionBlock>

      <SectionBlock
        title="CreatableSelect · PromptDialog"
        hint={<>Dropdown có mục cuối “Khác…” mở hộp thoại nhập giá trị mới; nơi gọi lưu giá trị mới vào danh sách (<code>onChange(value, isNew)</code>)</>}
      >
        <div className="admin-guide-row">
          <CreatableSelect
            value={kind}
            options={kinds}
            onChange={(v, isNew) => {
              if (isNew) setKinds(kinds.concat(v));
              setKind(v);
            }}
            dialogTitle="Thêm loại sự kiện mới"
          />
          <button type="button" className="btn btn--secondary" onClick={() => setPromptOpen(true)}>
            Mở PromptDialog
          </button>
        </div>
        {promptOpen && (
          <PromptDialog
            title="Nhập tên mới"
            label="Tên"
            onSubmit={() => setPromptOpen(false)}
            onCancel={() => setPromptOpen(false)}
          />
        )}
      </SectionBlock>

      <SectionBlock title="Trạng thái & thông báo" hint={<><code>LoadingState</code>, <code>EmptyState</code>, <code>ToastAlert</code></>}>
        <LoadingState message="Đang tải dữ liệu…" />
        <EmptyState message="Không tìm thấy kết quả phù hợp" />
        <EmptyState narrow tone="danger" icon="error-circle" title="Chưa được phân quyền" message="Liên hệ quản trị viên để được cấp quyền." />
        <div className="admin-guide-row">
          <button type="button" className="btn btn--secondary btn--sm" onClick={() => setToast({ type: 'success', text: 'Đã lưu thành công' })}>Toast thành công</button>
          <button type="button" className="btn btn--secondary btn--sm" onClick={() => setToast({ type: 'error', text: 'Có lỗi xảy ra khi lưu' })}>Toast lỗi</button>
        </div>
        <ToastAlert toast={toast} onClose={() => setToast(undefined)} />
      </SectionBlock>

      <SectionBlock title="Bảng" hint={<>Luôn bọc <code>.table-responsive</code> để cuộn ngang trên màn hẹp.</>}>
        <div className="table-responsive">
          <table className="portal-table">
            <thead>
              <tr><th>Mã</th><th>Nội dung</th><th className="text-right">Số lượng</th></tr>
            </thead>
            <tbody>
              <tr><td>VM-001</td><td>Vướng tuyến ống cấp nước</td><td className="text-right">12</td></tr>
              <tr><td>VM-002</td><td>Phân luồng giao thông</td><td className="text-right">8</td></tr>
            </tbody>
          </table>
        </div>
      </SectionBlock>

      <SectionBlock title="TileGrid" hint={<>Lưới ô chức năng dùng ở trang chủ và menu module: <code>{'<TileGrid items={[{ id, title, icon, onClick }]} />'}</code></>}>
        <TileGrid
          items={[
            { id: 'a', title: 'Ô chức năng A', icon: <Icon name="paint" size={26} />, onClick: () => undefined },
            { id: 'b', title: 'Ô chức năng B', icon: <Icon name="power" size={26} />, onClick: () => undefined }
          ]}
        />
      </SectionBlock>

      <SectionBlock title={`Icon (${ICON_NAMES.length})`} hint={<><code>{'<Icon name="…" size={16|20|24} />'}</code> – Fluent UI Regular</>}>
        <div className="admin-guide-grid">
          {ICON_NAMES.map((n) => (
            <div key={n} className="admin-guide-icon-cell">
              <Icon name={n} size={24} />
              <span className="admin-guide-icon-name">{n}</span>
            </div>
          ))}
        </div>
      </SectionBlock>
    </div>
  );
};
