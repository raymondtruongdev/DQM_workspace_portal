import * as React from 'react';
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
  /** Thư mục (server-relative) chứa ảnh của bộ phận. */
  folder: string;
  files: string[];
  disabled?: boolean;
  isUploading?: boolean;
  onUpload: (files: FileList) => Promise<void>;
  onDelete: (name: string) => Promise<void>;
}

/** Tải / xem / xoá ảnh của một bộ phận trong photos/<groupId>/. */
export const ReportPhotoManager: React.FC<IReportPhotoManagerProps> = ({
  folder,
  files,
  disabled = false,
  isUploading = false,
  onUpload,
  onDelete
}) => {
  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  const handleFileInputChange = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const picked = e.target.files;
    if (picked && picked.length > 0) {
      await onUpload(picked);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  return (
    <div className="report-photos-upload-box">
      <div className="report-photos-header">
        <div className="report-photos-subtitle">
          Đã tải lên <strong>{files.length}</strong> ảnh
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
            <span>{isUploading ? 'Đang tải ảnh...' : 'Tải ảnh lên'}</span>
          </button>
        </div>
      </div>

      {files.length > 0 && (
        <div className="report-photos-list">
          {files.map((name) => (
            <PhotoCard
              key={name}
              url={`${folder}/${name}`}
              name={name}
              disabled={disabled || isUploading}
              onDelete={() => onDelete(name)}
            />
          ))}
        </div>
      )}
    </div>
  );
};
