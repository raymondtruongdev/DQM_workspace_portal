import * as React from 'react';
import { ReportPhotoFile } from '../types/report.types';
import { Icon } from '../../../core/components/Icon';

interface IPhotoCardProps {
  url: string;
  name: string;
  disabled: boolean;
  onDelete: () => void;
}

/** Ảnh hiện trường: hình thu nhỏ (bấm để mở ảnh gốc) + tên file + nút xoá. */
const PhotoCard: React.FC<IPhotoCardProps> = ({ url, name, disabled, onDelete }) => {
  const [failed, setFailed] = React.useState(false);
  return (
    <figure className="report-photo-item">
      <a
        href={url}
        target="_blank"
        rel="noopener noreferrer"
        className="report-photo-thumb"
        title={`Mở ảnh: ${name}`}
      >
        {failed ? (
          <span className="report-photo-fallback">
            <Icon name="image" size={24} />
          </span>
        ) : (
          <img src={url} alt={name} loading="lazy" onError={() => setFailed(true)} />
        )}
      </a>
      <figcaption className="report-photo-caption">
        <span className="report-photo-name" title={name}>
          {name}
        </span>
        <button
          type="button"
          className="btn btn--ghost-danger btn--icon btn--sm"
          onClick={onDelete}
          disabled={disabled}
          title="Xoá ảnh này khỏi SharePoint"
          aria-label={`Xoá ảnh ${name}`}
        >
          <Icon name="dismiss" size={16} />
        </button>
      </figcaption>
    </figure>
  );
};

export interface IReportPhotoManagerProps {
  unit: string;
  docId: string;
  dateStr: string;
  photos: ReportPhotoFile[];
  disabled?: boolean;
  isUploading?: boolean;
  onUpload: (files: FileList) => Promise<void>;
  onDelete: (serverRelativeUrl: string, name: string) => Promise<void>;
}

/** Khu tải / xem / xoá ảnh hiện trường trong data/<Unit>/<DocID>_<yyyyMMdd>/photos/. */
export const ReportPhotoManager: React.FC<IReportPhotoManagerProps> = ({
  unit,
  docId,
  dateStr,
  photos,
  disabled = false,
  isUploading = false,
  onUpload,
  onDelete
}) => {
  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  const handleFileInputChange = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const files = e.target.files;
    if (files && files.length > 0) {
      await onUpload(files);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  return (
    <div className="report-photos-upload-box">
      <div className="report-photos-header">
        <div>
          <div className="report-photos-title">
            Thư mục ảnh hiện trường:{' '}
            <code>
              data/{unit}/{docId}_{dateStr.replace(/-/g, '')}/photos/
            </code>
          </div>
          <div className="report-photos-subtitle">
            Đã tải lên <strong>{photos.length}</strong> ảnh trên SharePoint
          </div>
        </div>

        <div>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            multiple
            style={{ display: 'none' }}
            onChange={handleFileInputChange}
            disabled={disabled || isUploading}
          />
          <button
            type="button"
            className="btn btn--primary btn--sm"
            onClick={() => fileInputRef.current?.click()}
            disabled={disabled || isUploading}
          >
            <Icon name="upload" size={16} />
            <span>{isUploading ? 'Đang tải ảnh...' : 'Tải ảnh hiện trường lên'}</span>
          </button>
        </div>
      </div>

      {photos.length > 0 && (
        <div className="report-photos-list">
          {photos.map((p) => (
            <PhotoCard
              key={p.serverRelativeUrl}
              url={p.serverRelativeUrl}
              name={p.name}
              disabled={disabled || isUploading}
              onDelete={() => onDelete(p.serverRelativeUrl, p.name)}
            />
          ))}
        </div>
      )}
    </div>
  );
};
