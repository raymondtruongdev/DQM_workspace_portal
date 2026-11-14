import * as React from 'react';
import { SignedDocInfo } from '../types/report.types';
import type { ISignedTrail } from '../services/ReportRegistryService';
import { Icon } from '../../../core/components';
import { formatFileSize, formatSavedAt } from '../utils/reportCalendar';

export interface IReportSignedDocumentProps {
  /** Thư mục báo cáo hiển thị cho người dùng, vd data/ST01/DRP_20261114_ST01/ */
  folderLabel: string;
  /** Báo cáo đã lưu lên SharePoint (và không còn thay đổi chưa lưu). */
  canUpload: boolean;
  /** Lý do chưa upload được (hiện khi canUpload = false). */
  blockedReason: string;
  signed: SignedDocInfo | undefined;
  isLoading: boolean;
  isUploading: boolean;
  disabled?: boolean;
  onUpload: (file: File) => Promise<void>;
  /** Xoá file đã ký (có hộp thoại xác nhận + lý do). Không truyền = không hiện nút. */
  onRemove?: () => void;
  isRemoving?: boolean;
  /** Dấu vết bản đã ký từ sổ ReportRegistry: người upload (khi đã có bản ký), lần xoá gần nhất (khi chưa có). */
  trail?: ISignedTrail;
}

/**
 * Mục "BẢN ĐÃ KÝ": sau khi lưu báo cáo, người dùng in PDF, ký tên rồi upload file đã ký vào
 * thư mục báo cáo. Có file PDF trong thư mục = ngày đó Hoàn tất.
 */
export const ReportSignedDocument: React.FC<IReportSignedDocumentProps> = ({
  folderLabel,
  canUpload,
  blockedReason,
  signed,
  isLoading,
  isUploading,
  disabled = false,
  onUpload,
  onRemove,
  isRemoving = false,
  trail
}) => {
  const fileRef = React.useRef<HTMLInputElement | null>(null);
  const [dragOver, setDragOver] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');

  const busy = disabled || isUploading || isLoading || isRemoving;

  const accept = async (file: File | undefined): Promise<void> => {
    if (!file) return;
    if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') {
      setError('Chỉ nhận file PDF.');
      return;
    }
    setError('');
    await onUpload(file);
  };

  const step = (n: number, title: string, hint: string, state: 'done' | 'now' | 'off'): React.ReactNode => (
    <li className={`report-signed-step report-signed-step--${state}`}>
      <span className="report-signed-step__num">{state === 'done' ? <Icon name="check-circle" size={16} /> : n}</span>
      <div>
        <b>{title}</b>
        <span>{hint}</span>
      </div>
    </li>
  );

  const hasSigned = Boolean(signed);

  return (
    <div id="report-sec-signed" className="section-card">
      <div className="section-card__banner">
        <span className="section-card__title">
          BẢN ĐÃ KÝ <span className="section-card__title-sub">| SIGNED DOCUMENT</span>
        </span>
      </div>

      <div className="section-card__body">
        <ol className="report-signed-steps">
          {step(1, 'Lưu báo cáo', canUpload || hasSigned ? 'Đã lưu' : 'Bấm Lưu để tạo báo cáo', canUpload || hasSigned ? 'done' : 'now')}
          {step(2, 'In PDF và ký tên', 'In bản A4, ký trực tiếp', hasSigned ? 'done' : canUpload ? 'now' : 'off')}
          {step(3, 'Upload bản đã ký', hasSigned ? 'Hoàn tất' : 'Chọn file PDF đã ký', hasSigned ? 'done' : 'off')}
        </ol>

        <input
          ref={fileRef}
          type="file"
          accept="application/pdf,.pdf"
          className="report-signed-file-pick"
          tabIndex={-1}
          aria-label="Chọn file PDF đã ký"
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            accept(f || undefined).catch((err) => console.error(err));
            e.target.value = '';
          }}
        />

        {hasSigned && signed ? (
          <div className="report-signed-file">
            <div className="report-signed-file__info">
              <Icon name="file-up" size={24} />
              <div>
                <a href={signed.serverRelativeUrl} target="_blank" rel="noopener noreferrer" className="report-signed-file__name">
                  {signed.name}
                </a>
                <div className="report-signed-file__meta">
                  {formatFileSize(signed.length)} · {formatSavedAt(signed.timeLastModified)}
                  {trail && trail.uploadBy ? <> · upload bởi {trail.uploadBy}</> : null}
                </div>
              </div>
            </div>
            <div className="report-signed-file__actions">
              <a className="btn btn--secondary btn--sm" href={signed.serverRelativeUrl} target="_blank" rel="noopener noreferrer">
                <Icon name="download" size={16} />
                <span>Tải xuống</span>
              </a>
              <button
                type="button"
                className="btn btn--secondary btn--sm"
                onClick={() => fileRef.current?.click()}
                disabled={busy || !canUpload}
                title={canUpload ? 'Thay bằng file PDF khác' : blockedReason}
              >
                <Icon name="upload" size={16} />
                <span>{isUploading ? 'Đang tải...' : 'Thay file'}</span>
              </button>
              {onRemove && (
                <button
                  type="button"
                  className="btn btn--danger btn--sm"
                  onClick={onRemove}
                  disabled={busy}
                  title="Xoá file đã ký (có ghi nhật ký)"
                >
                  <Icon name="delete" size={16} />
                  <span>{isRemoving ? 'Đang xoá...' : 'Xoá file đã ký'}</span>
                </button>
              )}
            </div>
          </div>
        ) : (
          <div
            className={`report-signed-drop${dragOver ? ' is-over' : ''}${canUpload ? '' : ' is-off'}`}
            onDragOver={(e) => {
              if (!canUpload || busy) return;
              e.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              if (!canUpload || busy) return;
              accept(e.dataTransfer.files && e.dataTransfer.files[0]).catch((err) => console.error(err));
            }}
          >
            <Icon name="upload" size={28} />
            <p>{canUpload ? 'Kéo thả file PDF đã ký vào đây hoặc chọn từ máy' : blockedReason}</p>
            <button
              type="button"
              className="btn btn--primary"
              onClick={() => fileRef.current?.click()}
              disabled={busy || !canUpload}
            >
              <Icon name="upload" size={16} />
              <span>{isUploading ? 'Đang tải lên...' : 'Upload file đã ký'}</span>
            </button>
          </div>
        )}

        {!hasSigned && trail && trail.removedAt && (
          <p className="report-signed-folder">
            Bản đã ký từng bị xoá gần nhất lúc {formatSavedAt(trail.removedAt)}
            {trail.removedBy ? <> bởi <b>{trail.removedBy}</b></> : null}.
          </p>
        )}

        {error && (
          <p className="report-field-error" role="alert">
            {error}
          </p>
        )}
        <p className="report-signed-folder">
          File được lưu cùng thư mục với báo cáo: <code>{folderLabel}</code>
        </p>
      </div>
    </div>
  );
};
