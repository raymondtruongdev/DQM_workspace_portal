import * as React from 'react';
import {
  EquipmentRow,
  IssueRow,
  ManpowerRow,
  MaterialRow,
  ReportDoc,
  ReportMetadata,
  WorkQtyRow
} from '../types/report.types';
import { Icon } from '../../../core/components/Icon';
import { buildSectionNumbering, hasSection } from '../utils/reportSections';

export interface IReportSectionsEditorProps {
  doc: ReportDoc;
  disabled?: boolean;
  isUploadingPhoto?: boolean;
  onMetadataChange: <K extends keyof ReportMetadata>(field: K, value: ReportMetadata[K]) => void;
  onManpowerChange: (rows: ManpowerRow[]) => void;
  onEquipmentChange: (rows: EquipmentRow[]) => void;
  onWorkItemsChange: (rows: WorkQtyRow[]) => void;
  onMaterialsChange: (rows: MaterialRow[]) => void;
  onIssuesChange: (rows: IssueRow[]) => void;
  onUploadPhotos: (files: FileList) => Promise<void>;
  onDeletePhoto: (serverRelativeUrl: string, name: string) => Promise<void>;
}

function parseInputNum(val: string): number | undefined {
  if (val.trim() === '') return undefined;
  const n = parseFloat(val);
  return isNaN(n) ? undefined : n;
}

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

export const ReportSectionsEditor: React.FC<IReportSectionsEditorProps> = ({
  doc,
  disabled = false,
  isUploadingPhoto = false,
  onMetadataChange,
  onManpowerChange,
  onEquipmentChange,
  onWorkItemsChange,
  onMaterialsChange,
  onIssuesChange,
  onUploadPhotos,
  onDeletePhoto
}) => {
  const fileInputRef = React.useRef<HTMLInputElement | null>(null);
  const { metadata, manpower, equipment, workItems, materials, issues, photos = [] } = doc;
  const num = buildSectionNumbering(doc);

  // Totals for Manpower
  const totalNos = manpower.reduce((acc, r) => acc + (r.nos || 0), 0);
  const totalManHrs = manpower.reduce((acc, r) => acc + (r.manHrs || 0), 0);

  // --- Section II: Manpower Handlers ---
  const updateManpowerRow = (index: number, patch: Partial<ManpowerRow>): void => {
    const next = manpower.map((row, idx) => {
      if (idx !== index) return row;
      const updated = { ...row, ...patch };
      const nos = updated.nos ?? 0;
      const hrs = updated.hrs ?? 0;
      updated.manHrs = Math.round(nos * hrs * 100) / 100;
      return updated;
    });
    onManpowerChange(next);
  };

  const addManpowerRow = (): void => {
    onManpowerChange([
      ...manpower,
      {
        id: `mp-${Date.now()}`,
        org: '',
        category: '',
        nos: 0,
        hrs: 8,
        manHrs: 0,
        isLocked: false
      }
    ]);
  };

  const removeManpowerRow = (index: number): void => {
    if (manpower[index]?.isLocked) return;
    onManpowerChange(manpower.filter((_, idx) => idx !== index));
  };

  // --- Section III: Equipment Handlers ---
  const updateEquipmentRow = (index: number, patch: Partial<EquipmentRow>): void => {
    const next = equipment.map((row, idx) => (idx === index ? { ...row, ...patch } : row));
    onEquipmentChange(next);
  };

  const addEquipmentRow = (): void => {
    onEquipmentChange([
      ...equipment,
      {
        id: `eq-${Date.now()}`,
        equipment: '',
        qty: 1,
        hrs: 8,
        status: 'Hoạt động',
        remark: '',
        isLocked: false
      }
    ]);
  };

  const removeEquipmentRow = (index: number): void => {
    if (equipment[index]?.isLocked) return;
    onEquipmentChange(equipment.filter((_, idx) => idx !== index));
  };

  // --- Section IV: WorkQty Handlers ---
  const updateWorkQtyRow = (index: number, patch: Partial<WorkQtyRow>): void => {
    const next = workItems.map((row, idx) => {
      if (idx !== index) return row;
      const updated = { ...row, ...patch };
      if (updated.plan !== undefined && updated.plan > 0 && updated.actual !== undefined) {
        updated.percent = Math.round((updated.actual / updated.plan) * 100);
      }
      return updated;
    });
    onWorkItemsChange(next);
  };

  const addWorkQtyRow = (): void => {
    const nextNum = (workItems.length + 1) * 10 + 1100;
    onWorkItemsChange([
      ...workItems,
      {
        id: `wq-${Date.now()}`,
        activityId: `${doc.unit}-DW-${nextNum}`,
        item: '',
        unit: 'm3',
        plan: 0,
        actual: 0,
        percent: 0,
        remark: '',
        isLocked: false
      }
    ]);
  };

  const removeWorkQtyRow = (index: number): void => {
    if (workItems[index]?.isLocked) return;
    onWorkItemsChange(workItems.filter((_, idx) => idx !== index));
  };

  // --- Section V: Materials Handlers ---
  const updateMaterialRow = (index: number, patch: Partial<MaterialRow>): void => {
    const next = materials.map((row, idx) => (idx === index ? { ...row, ...patch } : row));
    onMaterialsChange(next);
  };

  const addMaterialRow = (): void => {
    onMaterialsChange([
      ...materials,
      {
        id: `mat-${Date.now()}`,
        material: '',
        unit: 'm3',
        recv: 0,
        used: 0,
        mirMar: `MIR-${doc.unit}-`,
        isLocked: false
      }
    ]);
  };

  const removeMaterialRow = (index: number): void => {
    if (materials[index]?.isLocked) return;
    onMaterialsChange(materials.filter((_, idx) => idx !== index));
  };

  // --- Section VII: Issues Handlers ---
  const updateIssueRow = (index: number, patch: Partial<IssueRow>): void => {
    const next = issues.map((row, idx) => (idx === index ? { ...row, ...patch } : row));
    onIssuesChange(next);
  };

  const addIssueRow = (): void => {
    const codeNum = ('000' + (issues.length + 12)).slice(-3);
    onIssuesChange([
      ...issues,
      {
        id: `iss-${Date.now()}`,
        issueId: `VM-${doc.unit}-${codeNum}`,
        description: '',
        severity: 'Trung bình',
        criticalPath: '–',
        owner: '',
        dueDate: '',
        status: 'Mở',
        isLocked: false
      }
    ]);
  };

  const removeIssueRow = (index: number): void => {
    if (issues[index]?.isLocked) return;
    onIssuesChange(issues.filter((_, idx) => idx !== index));
  };

  const handleFileInputChange = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const files = e.target.files;
    if (files && files.length > 0) {
      await onUploadPhotos(files);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  return (
    <div className="report-sections-stack">
      {/* II. NHÂN LỰC | MANPOWER */}
      {hasSection(doc, 'manpower') && (
        <div id="report-sec-manpower" className="section-card">
          <div className="section-card__banner">
            <span className="section-card__title">
              {num.roman('manpower')}. NHÂN LỰC <span className="section-card__title-sub">| MANPOWER</span>
            </span>
            <button
              type="button"
              className="btn btn--light btn--sm"
              onClick={addManpowerRow}
              disabled={disabled}
            >
              + Thêm nhân lực
            </button>
          </div>

          <div className="table-responsive">
            <table className="portal-table report-grid-table">
              <thead>
                <tr>
                  <th style={{ width: '16%' }}>Đơn vị | Org.</th>
                  <th style={{ width: '32%' }}>Loại | Category</th>
                  <th style={{ width: '16%', textAlign: 'right' }}>Số người | Nos.</th>
                  <th style={{ width: '16%', textAlign: 'right' }}>Giờ/người | Hrs</th>
                  <th style={{ width: '14%', textAlign: 'right' }}>Giờ công | Man-hrs</th>
                  <th style={{ width: '6%', textAlign: 'center' }} />
                </tr>
              </thead>
              <tbody>
                {manpower.map((row, idx) => (
                  <tr key={row.id || idx} className={row.isLocked ? 'report-row-locked' : ''}>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.org}
                        onChange={(e) => updateManpowerRow(idx, { org: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                        placeholder="Đơn vị..."
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.category}
                        onChange={(e) => updateManpowerRow(idx, { category: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                        placeholder="Loại nhân lực..."
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        className="input input--cell text-right"
                        value={row.nos !== undefined ? row.nos : ''}
                        onChange={(e) => updateManpowerRow(idx, { nos: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="0.5"
                        className="input input--cell text-right"
                        value={row.hrs !== undefined ? row.hrs : ''}
                        onChange={(e) => updateManpowerRow(idx, { hrs: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td className="text-right font-semibold">{(row.manHrs || 0).toLocaleString('vi-VN')}</td>
                    <td className="text-center">
                      {row.isLocked ? (
                        <span
                          className="report-row-lock-badge"
                          title="Mục cố định theo template (không được xoá)"
                        >
                          <Icon name="lock" size={16} />
                        </span>
                      ) : (
                        <button
                          type="button"
                          className="btn btn--ghost-danger btn--icon btn--sm"
                          onClick={() => removeManpowerRow(idx)}
                          disabled={disabled}
                          title="Xoá dòng"
                        >
                          <Icon name="dismiss" size={16} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
                <tr className="report-table-total-row">
                  <td colSpan={2}>
                    <strong>Tổng | Total</strong>
                  </td>
                  <td className="text-right">
                    <strong>{totalNos.toLocaleString('vi-VN')}</strong>
                  </td>
                  <td />
                  <td className="text-right">
                    <strong>{totalManHrs.toLocaleString('vi-VN')}</strong>
                  </td>
                  <td />
                </tr>
              </tbody>
            </table>
          </div>

          <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
            <div className="field">
              <label className="field__label">Tóm tắt | Summary</label>
              <textarea
                className="textarea"
                rows={2}
                placeholder="Nhập tóm tắt tình hình nhân lực (nếu có)..."
                value={metadata.manpowerSummary || ''}
                onChange={(e) => onMetadataChange('manpowerSummary', e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>
        </div>
      )}

      {/* III. THIẾT BỊ | EQUIPMENT */}
      {hasSection(doc, 'equipment') && (
        <div id="report-sec-equipment" className="section-card">
          <div className="section-card__banner">
            <span className="section-card__title">
              {num.roman('equipment')}. THIẾT BỊ <span className="section-card__title-sub">| EQUIPMENT</span>
            </span>
            <button
              type="button"
              className="btn btn--light btn--sm"
              onClick={addEquipmentRow}
              disabled={disabled}
            >
              + Thêm thiết bị
            </button>
          </div>

          <div className="table-responsive">
            <table className="portal-table report-grid-table">
              <thead>
                <tr>
                  <th style={{ width: '32%' }}>Thiết bị | Equipment</th>
                  <th style={{ width: '11%', textAlign: 'right' }}>SL | Qty</th>
                  <th style={{ width: '13%', textAlign: 'right' }}>Giờ HĐ | Hrs</th>
                  <th style={{ width: '18%', textAlign: 'center' }}>Tình trạng | Status</th>
                  <th style={{ width: '20%' }}>Ghi chú | Remark</th>
                  <th style={{ width: '6%', textAlign: 'center' }} />
                </tr>
              </thead>
              <tbody>
                {equipment.map((row, idx) => (
                  <tr key={row.id || idx} className={row.isLocked ? 'report-row-locked' : ''}>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.equipment}
                        onChange={(e) => updateEquipmentRow(idx, { equipment: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                        placeholder="Tên thiết bị / máy móc..."
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        className="input input--cell text-right"
                        value={row.qty !== undefined ? row.qty : ''}
                        onChange={(e) => updateEquipmentRow(idx, { qty: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="0.5"
                        className="input input--cell text-right"
                        value={row.hrs !== undefined ? row.hrs : ''}
                        onChange={(e) => updateEquipmentRow(idx, { hrs: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <select
                        className="input input--cell text-center"
                        value={row.status || 'Hoạt động'}
                        onChange={(e) => updateEquipmentRow(idx, { status: e.target.value })}
                        disabled={disabled}
                      >
                        <option value="Hoạt động">Hoạt động</option>
                        <option value="Chờ việc">Chờ việc</option>
                        <option value="Bảo dưỡng">Bảo dưỡng</option>
                        <option value="Hỏng">Hỏng</option>
                      </select>
                    </td>
                    <td>
                      <input
                        type="text"
                        className="input input--cell"
                        value={row.remark}
                        onChange={(e) => updateEquipmentRow(idx, { remark: e.target.value })}
                        disabled={disabled}
                        placeholder="Ghi chú..."
                      />
                    </td>
                    <td className="text-center">
                      {row.isLocked ? (
                        <span
                          className="report-row-lock-badge"
                          title="Mục cố định theo template (không được xoá)"
                        >
                          <Icon name="lock" size={16} />
                        </span>
                      ) : (
                        <button
                          type="button"
                          className="btn btn--ghost-danger btn--icon btn--sm"
                          onClick={() => removeEquipmentRow(idx)}
                          disabled={disabled}
                          title="Xoá dòng"
                        >
                          <Icon name="dismiss" size={16} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
            <div className="field">
              <label className="field__label">Tóm tắt | Summary</label>
              <textarea
                className="textarea"
                rows={2}
                placeholder="Nhập tóm tắt tình trạng thiết bị (nếu có)..."
                value={metadata.equipmentSummary || ''}
                onChange={(e) => onMetadataChange('equipmentSummary', e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>
        </div>
      )}

      {/* IV. KHỐI LƯỢNG THỰC HIỆN | WORK QUANTITIES */}
      {hasSection(doc, 'workqty') && (
        <div id="report-sec-workqty" className="section-card">
          <div className="section-card__banner">
            <span className="section-card__title">
              {num.roman('workqty')}. KHỐI LƯỢNG THỰC HIỆN <span className="section-card__title-sub">| WORK QUANTITIES</span>
            </span>
            <button
              type="button"
              className="btn btn--light btn--sm"
              onClick={addWorkQtyRow}
              disabled={disabled}
            >
              + Thêm hạng mục
            </button>
          </div>

          <div className="table-responsive">
            <table className="portal-table report-grid-table">
              <thead>
                <tr>
                  <th style={{ width: '16%' }}>Activity ID</th>
                  <th style={{ width: '28%' }}>Hạng mục | Item</th>
                  <th style={{ width: '9%', textAlign: 'center' }}>ĐV | Unit</th>
                  <th style={{ width: '11%', textAlign: 'right' }}>KH ngày | Plan</th>
                  <th style={{ width: '11%', textAlign: 'right' }}>Thực hiện | Actual</th>
                  <th style={{ width: '8%', textAlign: 'right' }}>%</th>
                  <th style={{ width: '12%' }}>Ghi chú | Remark</th>
                  <th style={{ width: '5%', textAlign: 'center' }} />
                </tr>
              </thead>
              <tbody>
                {workItems.map((row, idx) => (
                  <tr key={row.id || idx} className={row.isLocked ? 'report-row-locked' : ''}>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.activityId}
                        onChange={(e) => updateWorkQtyRow(idx, { activityId: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.item}
                        onChange={(e) => updateWorkQtyRow(idx, { item: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell text-center ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.unit}
                        onChange={(e) => updateWorkQtyRow(idx, { unit: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        className="input input--cell text-right"
                        value={row.plan !== undefined ? row.plan : ''}
                        onChange={(e) => updateWorkQtyRow(idx, { plan: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        className="input input--cell text-right"
                        value={row.actual !== undefined ? row.actual : ''}
                        onChange={(e) => updateWorkQtyRow(idx, { actual: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td className="text-right font-semibold">
                      {row.percent !== undefined ? `${row.percent}%` : '—'}
                    </td>
                    <td>
                      <input
                        type="text"
                        className="input input--cell"
                        value={row.remark}
                        onChange={(e) => updateWorkQtyRow(idx, { remark: e.target.value })}
                        disabled={disabled}
                      />
                    </td>
                    <td className="text-center">
                      {row.isLocked ? (
                        <span
                          className="report-row-lock-badge"
                          title="Mục cố định theo template (không được xoá)"
                        >
                          <Icon name="lock" size={16} />
                        </span>
                      ) : (
                        <button
                          type="button"
                          className="btn btn--ghost-danger btn--icon btn--sm"
                          onClick={() => removeWorkQtyRow(idx)}
                          disabled={disabled}
                          title="Xoá dòng"
                        >
                          <Icon name="dismiss" size={16} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
            <div className="field">
              <label className="field__label">Tóm tắt | Summary</label>
              <textarea
                className="textarea"
                rows={2}
                placeholder="Nhập tóm tắt khối lượng thực hiện (nếu có)..."
                value={metadata.workQtySummary || ''}
                onChange={(e) => onMetadataChange('workQtySummary', e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>
        </div>
      )}

      {/* V. VẬT TƯ | MATERIALS */}
      {hasSection(doc, 'materials') && (
        <div id="report-sec-materials" className="section-card">
          <div className="section-card__banner">
            <span className="section-card__title">
              {num.roman('materials')}. VẬT TƯ <span className="section-card__title-sub">| MATERIALS</span>
            </span>
            <button
              type="button"
              className="btn btn--light btn--sm"
              onClick={addMaterialRow}
              disabled={disabled}
            >
              + Thêm vật tư
            </button>
          </div>

          <div className="table-responsive">
            <table className="portal-table report-grid-table">
              <thead>
                <tr>
                  <th style={{ width: '34%' }}>Vật tư | Material</th>
                  <th style={{ width: '10%', textAlign: 'center' }}>ĐV | Unit</th>
                  <th style={{ width: '14%', textAlign: 'right' }}>Nhập về | Recv.</th>
                  <th style={{ width: '14%', textAlign: 'right' }}>Sử dụng | Used</th>
                  <th style={{ width: '22%' }}>Phiếu MIR/MAR</th>
                  <th style={{ width: '6%', textAlign: 'center' }} />
                </tr>
              </thead>
              <tbody>
                {materials.map((row, idx) => (
                  <tr key={row.id || idx} className={row.isLocked ? 'report-row-locked' : ''}>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.material}
                        onChange={(e) => updateMaterialRow(idx, { material: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell text-center ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.unit}
                        onChange={(e) => updateMaterialRow(idx, { unit: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        className="input input--cell text-right"
                        value={row.recv !== undefined ? row.recv : ''}
                        onChange={(e) => updateMaterialRow(idx, { recv: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        className="input input--cell text-right"
                        value={row.used !== undefined ? row.used : ''}
                        onChange={(e) => updateMaterialRow(idx, { used: parseInputNum(e.target.value) })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className="input input--cell"
                        value={row.mirMar}
                        onChange={(e) => updateMaterialRow(idx, { mirMar: e.target.value })}
                        disabled={disabled}
                      />
                    </td>
                    <td className="text-center">
                      {row.isLocked ? (
                        <span
                          className="report-row-lock-badge"
                          title="Mục cố định theo template (không được xoá)"
                        >
                          <Icon name="lock" size={16} />
                        </span>
                      ) : (
                        <button
                          type="button"
                          className="btn btn--ghost-danger btn--icon btn--sm"
                          onClick={() => removeMaterialRow(idx)}
                          disabled={disabled}
                          title="Xoá dòng"
                        >
                          <Icon name="dismiss" size={16} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
            <div className="field">
              <label className="field__label">Tóm tắt | Summary</label>
              <textarea
                className="textarea"
                rows={2}
                placeholder="Nhập tóm tắt tình hình vật tư (nếu có)..."
                value={metadata.materialsSummary || ''}
                onChange={(e) => onMetadataChange('materialsSummary', e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>
        </div>
      )}

      {/* VI. CÔNG VIỆC THỰC HIỆN & KẾ HOẠCH NGÀY MAI | WORKS TODAY & PLAN FOR TOMORROW */}
      <div id="report-sec-worksplan" className="section-card">
        <div className="section-card__banner">
          <span className="section-card__title">
            {num.roman('worksplan')}. CÔNG VIỆC THỰC HIỆN &amp; KẾ HOẠCH NGÀY MAI{' '}
            <span className="section-card__title-sub">| WORKS TODAY &amp; PLAN FOR TOMORROW</span>
          </span>
        </div>

        <div className="section-card__body">
          <div className="report-two-col-grid">
            <div className="field">
              <label className="field__label">Hôm nay | Today</label>
              <textarea
                className="textarea"
                rows={4}
                placeholder="• Đổ bê tông tấm P-12...&#10;• Đào gầu ngoạm tấm P-13..."
                value={metadata.worksToday || ''}
                onChange={(e) => onMetadataChange('worksToday', e.target.value)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label">Ngày mai | Tomorrow</label>
              <textarea
                className="textarea"
                rows={4}
                placeholder="• Đào hoàn thành P-13...&#10;• Nghiệm thu cốt thép lồng P-14..."
                value={metadata.planTomorrow || ''}
                onChange={(e) => onMetadataChange('planTomorrow', e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>

          <div className="field" style={{ marginTop: '14px' }}>
            <label className="field__label">Tóm tắt | Summary</label>
            <textarea
              className="textarea"
              rows={2}
              placeholder="Nhập tóm tắt công việc thực hiện & kế hoạch (nếu có)..."
              value={metadata.worksPlanSummary || ''}
              onChange={(e) => onMetadataChange('worksPlanSummary', e.target.value)}
              disabled={disabled}
            />
          </div>
        </div>
      </div>

      {/* VII. VƯỚNG MẮC / RÀNG BUỘC ĐANG MỞ | OPEN ISSUES & CONSTRAINTS */}
      {hasSection(doc, 'issues') && (
        <div id="report-sec-issues" className="section-card">
          <div className="section-card__banner">
            <span className="section-card__title">
              {num.roman('issues')}. VƯỚNG MẮC / RÀNG BUỘC ĐANG MỞ{' '}
              <span className="section-card__title-sub">| OPEN ISSUES &amp; CONSTRAINTS</span>
            </span>
            <button
              type="button"
              className="btn btn--light btn--sm"
              onClick={addIssueRow}
              disabled={disabled}
            >
              + Thêm vướng mắc
            </button>
          </div>

          <div className="table-responsive">
            <table className="portal-table report-grid-table">
              <thead>
                <tr>
                  <th style={{ width: '13%' }}>Mã | ID</th>
                  <th style={{ width: '34%' }}>Nội dung | Description</th>
                  <th style={{ width: '11%', textAlign: 'center' }}>Mức | Sev.</th>
                  <th style={{ width: '8%', textAlign: 'center' }}>Găng | CP</th>
                  <th style={{ width: '15%' }}>Bên xử lý | By</th>
                  <th style={{ width: '8%', textAlign: 'center' }}>Hạn | Due</th>
                  <th style={{ width: '11%', textAlign: 'center' }}>Trạng thái</th>
                  <th style={{ width: '4%', textAlign: 'center' }} />
                </tr>
              </thead>
              <tbody>
                {issues.map((row, idx) => (
                  <tr key={row.id || idx} className={row.isLocked ? 'report-row-locked' : ''}>
                    <td>
                      <input
                        type="text"
                        className={`input input--cell ${row.isLocked ? 'input--locked' : ''}`}
                        value={row.issueId}
                        onChange={(e) => updateIssueRow(idx, { issueId: e.target.value })}
                        disabled={disabled}
                        readOnly={Boolean(row.isLocked)}
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className="input input--cell"
                        value={row.description}
                        onChange={(e) => updateIssueRow(idx, { description: e.target.value })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <select
                        className={`input input--cell text-center ${
                          row.severity === 'Cao' ? 'input--danger' : row.severity === 'Trung bình' ? 'input--warning' : ''
                        }`}
                        value={row.severity || 'Trung bình'}
                        onChange={(e) => updateIssueRow(idx, { severity: e.target.value })}
                        disabled={disabled}
                      >
                        <option value="Cao">Cao</option>
                        <option value="Trung bình">Trung bình</option>
                        <option value="Thấp">Thấp</option>
                      </select>
                    </td>
                    <td>
                      <select
                        className="input input--cell text-center"
                        value={row.criticalPath || '–'}
                        onChange={(e) => updateIssueRow(idx, { criticalPath: e.target.value })}
                        disabled={disabled}
                      >
                        <option value="Có">Có</option>
                        <option value="–">–</option>
                      </select>
                    </td>
                    <td>
                      <input
                        type="text"
                        className="input input--cell"
                        value={row.owner}
                        onChange={(e) => updateIssueRow(idx, { owner: e.target.value })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        className="input input--cell text-center"
                        placeholder="dd/MM"
                        value={row.dueDate}
                        onChange={(e) => updateIssueRow(idx, { dueDate: e.target.value })}
                        disabled={disabled}
                      />
                    </td>
                    <td>
                      <select
                        className="input input--cell text-center"
                        value={row.status || 'Mở'}
                        onChange={(e) => updateIssueRow(idx, { status: e.target.value })}
                        disabled={disabled}
                      >
                        <option value="Mở">Mở</option>
                        <option value="Đang xử lý">Đang xử lý</option>
                        <option value="Đóng">Đóng</option>
                      </select>
                    </td>
                    <td className="text-center">
                      {row.isLocked ? (
                        <span
                          className="report-row-lock-badge"
                          title="Mục cố định theo template (không được xoá)"
                        >
                          <Icon name="lock" size={16} />
                        </span>
                      ) : (
                        <button
                          type="button"
                          className="btn btn--ghost-danger btn--icon btn--sm"
                          onClick={() => removeIssueRow(idx)}
                          disabled={disabled}
                          title="Xoá dòng"
                        >
                          <Icon name="dismiss" size={16} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="section-card__body" style={{ borderTop: '1px solid var(--color-border)' }}>
            <div className="field">
              <label className="field__label">Tóm tắt | Summary</label>
              <textarea
                className="textarea"
                rows={2}
                placeholder="Nhập tóm tắt vướng mắc / ràng buộc (nếu có)..."
                value={metadata.issuesSummary || ''}
                onChange={(e) => onMetadataChange('issuesSummary', e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>
        </div>
      )}

      {/* VIII. AN TOÀN – CHẤT LƯỢNG – CHỈ THỊ TVGS | HSE – QUALITY – ENGINEER'S INSTRUCTIONS */}
      <div id="report-sec-hse" className="section-card">
        <div className="section-card__banner">
          <span className="section-card__title">
            {num.roman('hse')}. AN TOÀN – CHẤT LƯỢNG – CHỈ THỊ TVGS{' '}
            <span className="section-card__title-sub">
              | HSE – QUALITY – ENGINEER&apos;S INSTRUCTIONS
            </span>
          </span>
        </div>

        <div className="section-card__body">
          {/* 5 KPI Inputs */}
          <div className="report-hse-kpi-grid">
            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">Sự cố HSE</span>
                <span className="field__label-sub">Incidents</span>
              </label>
              <input
                type="number"
                min="0"
                className="input"
                value={metadata.hseIncidents ?? 0}
                onChange={(e) => onMetadataChange('hseIncidents', parseInputNum(e.target.value) ?? 0)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">Kiểm tra ATLĐ</span>
                <span className="field__label-sub">HSE inspections</span>
              </label>
              <input
                type="number"
                min="0"
                className="input"
                value={metadata.hseInspections ?? 0}
                onChange={(e) => onMetadataChange('hseInspections', parseInputNum(e.target.value) ?? 0)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">Nghiệm thu</span>
                <span className="field__label-sub">IR</span>
              </label>
              <input
                type="number"
                min="0"
                className="input"
                value={metadata.inspectionRequests ?? 0}
                onChange={(e) => onMetadataChange('inspectionRequests', parseInputNum(e.target.value) ?? 0)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">NCR mới</span>
                <span className="field__label-sub">New NCR</span>
              </label>
              <input
                type="number"
                min="0"
                className="input"
                value={metadata.newNCR ?? 0}
                onChange={(e) => onMetadataChange('newNCR', parseInputNum(e.target.value) ?? 0)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">Ảnh hiện trường</span>
                <span className="field__label-sub">Photos</span>
              </label>
              <input
                type="number"
                min="0"
                className="input"
                value={metadata.photosCount ?? photos.length}
                onChange={(e) => onMetadataChange('photosCount', parseInputNum(e.target.value) ?? 0)}
                disabled={disabled}
              />
            </div>
          </div>

          {/* Direct Photo Upload to data/<Unit>/<DocID>_<yyyyMMdd>/photos/ */}
          <div className="report-photos-upload-box">
            <div className="report-photos-header">
              <div>
                <div className="report-photos-title">
                  Thư mục ảnh hiện trường:{' '}
                  <code>
                    data/{doc.unit}/{doc.docId}_{doc.dateStr.replace(/-/g, '')}/photos/
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
                  disabled={disabled || isUploadingPhoto}
                />
                <button
                  type="button"
                  className="btn btn--primary btn--sm"
                  onClick={() => fileInputRef.current?.click()}
                  disabled={disabled || isUploadingPhoto}
                >
                  <Icon name="upload" size={16} />
                  <span>{isUploadingPhoto ? 'Đang tải ảnh...' : 'Tải ảnh hiện trường lên'}</span>
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
                    disabled={disabled || isUploadingPhoto}
                    onDelete={() => onDeletePhoto(p.serverRelativeUrl, p.name)}
                  />
                ))}
              </div>
            )}
          </div>

          <div className="field" style={{ marginTop: '14px' }}>
            <label className="field__label">Mô tả sự cố | Incident details</label>
            <input
              type="text"
              className="input"
              value={metadata.incidentDetails || ''}
              onChange={(e) => onMetadataChange('incidentDetails', e.target.value)}
              disabled={disabled}
            />
          </div>

          <div className="field" style={{ marginTop: '12px' }}>
            <label className="field__label">
              Chỉ thị/ý kiến TVGS | Engineer&apos;s instructions
            </label>
            <textarea
              className="textarea"
              rows={2}
              value={metadata.engineerInstructions || ''}
              onChange={(e) => onMetadataChange('engineerInstructions', e.target.value)}
              disabled={disabled}
            />
          </div>

          <div className="field" style={{ marginTop: '12px' }}>
            <label className="field__label">
              Kế hoạch ngày mai (tóm tắt) | Plan summary
            </label>
            <textarea
              className="textarea"
              rows={2}
              value={metadata.planSummary || ''}
              onChange={(e) => onMetadataChange('planSummary', e.target.value)}
              disabled={disabled}
            />
          </div>
        </div>
      </div>

      {/* IX. THÔNG TIN KÝ DUYỆT */}
      <div id="report-sec-signoff" className="section-card">
        <div className="section-card__banner">
          <span className="section-card__title">
            {num.roman('signoff')}. THÔNG TIN KÝ DUYỆT <span className="section-card__title-sub">| SIGN-OFF</span>
          </span>
        </div>

        <div className="section-card__body">
          <div className="report-signoff-grid">
            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">Người lập</span>
                <span className="field__label-sub">Prepared by</span>
              </label>
              <input
                type="text"
                className="input"
                value={metadata.preparedBy || metadata.email || ''}
                onChange={(e) => onMetadataChange('preparedBy', e.target.value)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">Trưởng ga / CHT</span>
                <span className="field__label-sub">Site manager</span>
              </label>
              <input
                type="text"
                className="input"
                value={metadata.siteManager || ''}
                onChange={(e) => onMetadataChange('siteManager', e.target.value)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">Đại diện TCJV</span>
                <span className="field__label-sub">TCJV representative</span>
              </label>
              <input
                type="text"
                className="input"
                value={metadata.tcjvRep || ''}
                onChange={(e) => onMetadataChange('tcjvRep', e.target.value)}
                disabled={disabled}
              />
            </div>

            <div className="field">
              <label className="field__label field__label--bilingual">
                <span className="field__label-main">TVGS xác nhận</span>
                <span className="field__label-sub">Engineer acknowledged</span>
              </label>
              <input
                type="text"
                className="input"
                value={metadata.engineerAck || ''}
                onChange={(e) => onMetadataChange('engineerAck', e.target.value)}
                disabled={disabled}
              />
            </div>
          </div>

          <div className="field" style={{ marginTop: '12px' }}>
            <label className="field__label">Ý kiến duyệt | Approval comment</label>
            <input
              type="text"
              className="input"
              placeholder="VD: Đồng ý..."
              value={metadata.approvalComment || ''}
              onChange={(e) => onMetadataChange('approvalComment', e.target.value)}
              disabled={disabled}
            />
          </div>
        </div>
      </div>

    </div>
  );
};
