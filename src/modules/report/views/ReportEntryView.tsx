import * as React from 'react';
import { RoleEntry, ToastState } from '../../../core/types';
import { ToastAlert } from '../../../core/components/layout/ToastAlert';
import { formatDateToYYYYMMDD } from '../../../core/utils';
import { ReportDoc, ReportMetadata, ReportSection } from '../types/report.types';
import { ReportDataService } from '../services/ReportDataService';
import { formatTemplateVersionDisplay } from '../utils/reportFileNaming';
import { buildSectionNumbering } from '../utils/reportSections';
import { recalculateAllRates } from '../utils/reportTaskTree';
import { getMissingRequiredFields, getFieldLabel, getMissingFormFields } from '../utils/reportValidation';
import { ReportMetaForm } from '../components/ReportMetaForm';
import { ReportTaskTable } from '../components/ReportTaskTable';
import { ReportSectionsEditor } from '../components/ReportSectionsEditor';
import { ReportPrintPreviewModal } from '../components/ReportPrintPreviewModal';
import { Icon, PageHeader, LoadingState, EmptyState } from '../../../core/components';

export interface IReportEntryViewProps {
  role: RoleEntry;
  currentUserEmail: string;
  dataService: ReportDataService;
  onBack: () => void;
}

export const ReportEntryView: React.FC<IReportEntryViewProps> = ({
  role,
  currentUserEmail,
  dataService,
  onBack
}) => {
  const [dateStr, setDateStr] = React.useState<string>(() => formatDateToYYYYMMDD(new Date()));
  const [doc, setDoc] = React.useState<ReportDoc | null>(null);
  const [templateBuffer, setTemplateBuffer] = React.useState<ArrayBuffer | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [isUploadingPhoto, setIsUploadingPhoto] = React.useState<boolean>(false);
  const [hasUnsavedChanges, setHasUnsavedChanges] = React.useState<boolean>(false);
  const [showRequiredErrors, setShowRequiredErrors] = React.useState<boolean>(false);
  const [showPrintPreview, setShowPrintPreview] = React.useState<boolean>(false);
  const [toast, setToast] = React.useState<ToastState | null>(null);

  // Load report data for selected date
  const loadData = React.useCallback(
    async (targetDate: string): Promise<void> => {
      setIsLoading(true);
      try {
        const result = await dataService.loadReport(
          role.docId,
          role.unit,
          targetDate,
          currentUserEmail,
          role.siteManager
        );
        setDoc(result.doc);
        setShowRequiredErrors(false);
        setTemplateBuffer(result.templateBuffer);
        setHasUnsavedChanges(false);
      } catch (err: unknown) {
        console.error('Error loading report:', err);
        const msg = err instanceof Error ? err.message : 'Lỗi không xác định khi tải mẫu báo cáo';
        setToast({ type: 'error', text: msg });
      } finally {
        setIsLoading(false);
      }
    },
    [dataService, role.docId, role.unit, role.siteManager, currentUserEmail]
  );

  React.useEffect(() => {
    loadData(dateStr).catch(() => {});
  }, [loadData, dateStr]);

  // Date change handler with unsaved changes prompt
  const handleDateChange = (newDate: string): void => {
    if (!newDate || newDate === dateStr) return;
    if (hasUnsavedChanges) {
      const confirmLeave = window.confirm(
        'Bạn có dữ liệu chưa lưu. Đổi ngày sẽ tải dữ liệu của ngày mới và bỏ qua các thay đổi này. Bạn có chắc chắn muốn tiếp tục?'
      );
      if (!confirmLeave) return;
    }
    setDateStr(newDate);
  };

  // Back button handler with unsaved changes prompt
  const handleBack = (): void => {
    if (hasUnsavedChanges) {
      const confirmLeave = window.confirm(
        'Bạn có thay đổi chưa lưu trong báo cáo. Bạn có chắc chắn muốn quay lại danh sách?'
      );
      if (!confirmLeave) return;
    }
    onBack();
  };

  // Generic metadata change handler
  const handleMetadataChange = <K extends keyof ReportMetadata>(
    field: K,
    value: ReportMetadata[K]
  ): void => {
    if (!doc) return;
    setDoc({
      ...doc,
      metadata: {
        ...doc.metadata,
        [field]: value
      }
    });
    setHasUnsavedChanges(true);
  };

  // Dynamic table handlers
  const handleSectionChange = (section: ReportSection): void => {
    if (!doc) return;
    setDoc({ ...doc, sections: doc.sections.map((x) => (x.key === section.key ? section : x)) });
    setHasUnsavedChanges(true);
  };

  // Task change handler (for WeeklyReport or legacy task hierarchy)
  const handleTaskChange = (
    taskId: string,
    field: 'completionRate' | 'quantity',
    value: number | undefined
  ): void => {
    if (!doc) return;

    const updatedTasks = doc.tasks.map((t) => {
      if (t.id === taskId) {
        return {
          ...t,
          [field]: value
        };
      }
      return t;
    });

    const recalculated = recalculateAllRates(updatedTasks);
    setDoc({
      ...doc,
      tasks: recalculated
    });
    setHasUnsavedChanges(true);
  };

  // Upload photos directly to data/<Unit>/<DocID>_<yyyyMMdd>/photos/
  const handleUploadPhotos = async (files: FileList): Promise<void> => {
    if (!doc || files.length === 0) return;
    setIsUploadingPhoto(true);
    try {
      const uploadedList = [...(doc.photos || [])];
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const buf = await file.arrayBuffer();
        const uploaded = await dataService.uploadPhoto(
          role.unit,
          role.docId,
          dateStr,
          file.name,
          buf
        );
        const existingIdx = uploadedList.findIndex(
          (p) => p.name.toLowerCase() === uploaded.name.toLowerCase()
        );
        if (existingIdx >= 0) {
          uploadedList[existingIdx] = uploaded;
        } else {
          uploadedList.push(uploaded);
        }
      }

      setDoc((prev) =>
        prev
          ? {
              ...prev,
              photos: uploadedList
            }
          : null
      );
      setHasUnsavedChanges(true);
      setToast({
        type: 'success',
        text: `Đã tải lên ${files.length} ảnh vào thư mục data/${role.unit}/${role.docId}_${dateStr.replace(/-/g, '')}/photos/!`
      });
    } catch (err: unknown) {
      console.error('Error uploading photo:', err);
      const msg = err instanceof Error ? err.message : 'Lỗi khi tải ảnh lên SharePoint';
      setToast({ type: 'error', text: msg });
    } finally {
      setIsUploadingPhoto(false);
    }
  };

  // Delete an uploaded photo
  const handleDeletePhoto = async (serverRelativeUrl: string, name: string): Promise<void> => {
    if (!doc) return;
    if (!window.confirm(`Bạn có chắc chắn muốn xoá ảnh "${name}" khỏi SharePoint?`)) return;
    setIsUploadingPhoto(true);
    try {
      await dataService.deletePhoto(serverRelativeUrl);
      const remaining = (doc.photos || []).filter((p) => p.serverRelativeUrl !== serverRelativeUrl);
      setDoc((prev) =>
        prev
          ? {
              ...prev,
              photos: remaining
            }
          : null
      );
      setHasUnsavedChanges(true);
      setToast({
        type: 'success',
        text: `Đã xoá ảnh "${name}" khỏi thư mục photos.`
      });
    } catch (err: unknown) {
      console.error('Error deleting photo:', err);
      const msg = err instanceof Error ? err.message : 'Lỗi khi xoá ảnh trên SharePoint';
      setToast({ type: 'error', text: msg });
    } finally {
      setIsUploadingPhoto(false);
    }
  };

  // Load new template when existing / previous-day report uses an outdated template
  const handleLoadLatestTemplate = async (): Promise<void> => {
    if (!doc) return;
    setIsLoading(true);
    try {
      const result = await dataService.loadLatestTemplateForReport(
        role.docId,
        role.unit,
        dateStr,
        currentUserEmail,
        role.siteManager,
        doc
      );
      setDoc(result.doc);
      setTemplateBuffer(result.templateBuffer);
      setHasUnsavedChanges(true);
      setToast({
        type: 'success',
        text: `Đã nạp biểu mẫu mới nhất (${formatTemplateVersionDisplay(
          result.doc.metadata.templateVersion
        )}). Bạn có thể tiếp tục nhập thông tin báo cáo!`
      });
    } catch (err: unknown) {
      console.error('Error loading latest template:', err);
      const msg = err instanceof Error ? err.message : 'Lỗi khi tải template mới từ SharePoint';
      setToast({ type: 'error', text: msg });
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeepCurrentTemplate = (): void => {
    if (!doc) return;
    setDoc({
      ...doc,
      templateUpdateRequired: false
    });
  };

  /** Returns true when all required fields are filled; otherwise highlights them and warns. */
  const ensureRequiredFilled = (action: string): boolean => {
    if (!doc) return false;
    const missing = getMissingRequiredFields(doc.metadata);
    const missingForm = getMissingFormFields(doc);
    if (missing.length === 0 && missingForm.length === 0) return true;
    setShowRequiredErrors(true);
    const labels = missing.map(getFieldLabel).concat(missingForm.map((m) => m.label));
    setToast({
      type: 'error',
      text: `Không thể ${action}. Vui lòng nhập: ${labels.join(', ')}.`
    });
    window.setTimeout(() => {
      const el = document.querySelector<HTMLElement>('[aria-invalid="true"]');
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.focus({ preventScroll: true });
      }
    }, 0);
    return false;
  };

  const handleOpenPrint = (): void => {
    if (ensureRequiredFilled('xuất báo cáo')) setShowPrintPreview(true);
  };

  // Save report (6-sheet workbook)
  const handleSave = async (): Promise<void> => {
    if (!doc || !ensureRequiredFilled('lưu báo cáo')) return;
    setIsSaving(true);
    try {
      const res = await dataService.saveReport(doc, templateBuffer);
      setDoc((prev) => (prev ? { ...prev, isExisting: true, eTag: res.eTag } : null));
      setHasUnsavedChanges(false);
      setToast({
        type: 'success',
        text: `Đã lưu báo cáo 6 sheets thành công vào thư viện SharePoint (${role.unit} - ${dateStr})!`
      });
    } catch (err: unknown) {
      console.error('Error saving report:', err);
      const msg = err instanceof Error ? err.message : 'Lỗi khi lưu báo cáo lên SharePoint';
      setToast({ type: 'error', text: msg });
    } finally {
      setIsSaving(false);
    }
  };

  const scrollToSection = (sectionId: string): void => {
    const el = document.getElementById(sectionId);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  };

  const isWeekly = role.docId.toLowerCase().includes('week');
  const reportMainTitle = doc?.metadata.docTitle || (isWeekly ? 'Báo cáo tuần' : 'Báo cáo ngày');

  return (
    <div className="report-entry-container">
      {/* 1. Header (thành phần chung PageHeader) */}
      <PageHeader
        divider
        leading={
          <button
            type="button"
            className="btn btn--secondary btn--icon"
            onClick={handleBack}
            disabled={isSaving}
            title="Quay lại danh sách báo cáo"
            aria-label="Quay lại"
          >
            <Icon name="arrow-left" />
          </button>
        }
        title={reportMainTitle}
        badge={<span className="badge badge--info">{role.unit}</span>}
        meta={
          <>
            {doc?.isExisting ? (
              <span className="badge badge--success">
                <Icon name="check-circle" size={16} />
                Đã lưu trên SharePoint
              </span>
            ) : (
              <span className="badge badge--info">
                <Icon name="info" size={16} />
                {doc?.copiedFromDate
                  ? `Kế thừa từ ngày ${doc.copiedFromDate.split('-').reverse().join('/')}`
                  : 'Tạo mới từ template (6 sheets)'}
              </span>
            )}
            {doc && (
              <span
                className={`badge ${doc.templateUpdateRequired ? 'badge--warning' : 'badge--success'}`}
                title="Phiên bản biểu mẫu (Template Version)"
              >
                Biểu mẫu (Template): {formatTemplateVersionDisplay(doc.metadata.templateVersion)}
              </span>
            )}
          </>
        }
        actions={
          <>
            {hasUnsavedChanges && (
              <span className="unsaved-badge" title="Có dữ liệu chưa lưu">
                Chưa lưu
              </span>
            )}
            <button
              type="button"
              className="btn btn--secondary"
              onClick={handleOpenPrint}
              disabled={isLoading || !doc || Boolean(doc.templateUpdateRequired)}
              title="Xem bản in 2 trang A4 và xuất PDF"
            >
              <Icon name="print" />
              <span>Xem bản in / Xuất PDF</span>
            </button>
            <button
              type="button"
              className="btn btn--secondary"
              onClick={() => loadData(dateStr)}
              disabled={isLoading || isSaving}
              title="Tải lại dữ liệu"
              aria-label="Làm mới"
            >
              <Icon name="refresh" spin={isLoading} />
              <span>{isLoading ? 'Đang tải...' : 'Làm mới'}</span>
            </button>
            <button
              type="button"
              className="btn btn--primary"
              onClick={handleSave}
              disabled={isLoading || isSaving || !doc || Boolean(doc.templateUpdateRequired)}
            >
              {isSaving ? (
                <>
                  <div className="loading-spinner-sm" />
                  <span>Đang lưu...</span>
                </>
              ) : (
                <>
                  <Icon name="save" />
                  <span>Lưu báo cáo</span>
                </>
              )}
            </button>
          </>
        }
      />

      {/* 2. Quick Section Navigation Pills */}
      {!isLoading && doc && !doc.templateUpdateRequired && (
        <div className="chips card report-sec-jump-bar" role="navigation" aria-label="Mục báo cáo">
          {buildSectionNumbering(doc).visible.map((item) => (
            <button
              key={item.id}
              type="button"
              className="chip"
              onClick={() => scrollToSection(item.anchor)}
            >
              {item.roman}. {item.label}
            </button>
          ))}
        </div>
      )}

      {/* 3. Toast Alert */}
      {toast && (
        <div style={{ marginBottom: '16px' }}>
          <ToastAlert toast={toast} onClose={() => setToast(null)} />
        </div>
      )}

      {/* 4. Main Content */}
      {isLoading ? (
        <LoadingState message="Đang kiểm tra phiên bản biểu mẫu (template) và nạp dữ liệu..." />
      ) : doc && doc.templateUpdateRequired ? (
        <div className="card report-template-upgrade-card">
          <div className="report-template-upgrade-header">
            <div className="report-template-upgrade-icon">
              <Icon name="file-up" size={24} />
            </div>
            <div className="report-template-upgrade-title-wrap">
              <h3 className="report-template-upgrade-title">
                Yêu cầu cập nhật Biểu mẫu (Template) mới cho trạm {role.unit}
              </h3>
              <p className="report-template-upgrade-subtitle">
                {doc.outdatedSourceType === 'previous_day'
                  ? `Báo cáo ngày liền trước (${
                      doc.copiedFromDate ? doc.copiedFromDate.split('-').reverse().join('/') : ''
                    }) đang sử dụng phiên bản template cũ. Vui lòng tải template mới nhất từ thư mục _template trên SharePoint để tiếp tục nhập báo cáo ngày ${dateStr
                      .split('-')
                      .reverse()
                      .join('/')}.`
                  : `File báo cáo ngày ${dateStr
                      .split('-')
                      .reverse()
                      .join('/')} hiện tại đang sử dụng phiên bản template cũ. Vui lòng tải template mới nhất từ thư mục _template trên SharePoint trước khi tiếp tục cập nhật.`}
              </p>
            </div>
            <div className="report-info-date-group">
              <label htmlFor="upgrade-date-picker" className="field__label">
                Ngày báo cáo
              </label>
              <div className="report-date-wrap">
                <input
                  id="upgrade-date-picker"
                  type="date"
                  className="input report-date-width"
                  value={dateStr}
                  onChange={(e) => handleDateChange(e.target.value)}
                />
              </div>
            </div>
          </div>

          <div className="report-template-version-compare">
            <div className="report-template-ver-box is-old">
              <span className="report-template-ver-box-label">
                {doc.outdatedSourceType === 'previous_day'
                  ? 'Phiên bản của báo cáo ngày liền trước'
                  : 'Phiên bản của file báo cáo hiện tại'}
              </span>
              <strong className="report-template-ver-box-value">
                {formatTemplateVersionDisplay(doc.outdatedTemplateVersion || doc.metadata.templateVersion)}
              </strong>
            </div>

            <div className="report-template-ver-arrow" aria-hidden="true">
              <Icon name="arrow-right" size={20} />
            </div>

            <div className="report-template-ver-box is-new">
              <span className="report-template-ver-box-label">
                Phiên bản mới nhất trên SharePoint (_template)
              </span>
              <strong className="report-template-ver-box-value">
                {formatTemplateVersionDisplay(doc.latestTemplateInfo?.version)}
              </strong>
              {doc.latestTemplateInfo?.fileName && (
                <span className="report-template-ver-box-file">
                  File: _template/{doc.latestTemplateInfo.fileName}
                </span>
              )}
            </div>
          </div>

          <div className="report-template-upgrade-actions">
            <button
              type="button"
              className="btn btn--primary"
              onClick={handleLoadLatestTemplate}
            >
              <Icon name="download" size={20} />
              <span>
                Tải Template Mới ({doc.latestTemplateInfo?.fileName || doc.latestTemplateInfo?.version}) &amp; Tiếp Tục Nhập
              </span>
            </button>

            {doc.isExisting && (
              <button
                type="button"
                className="btn btn--secondary"
                onClick={handleKeepCurrentTemplate}
              >
                Xem dữ liệu cũ (không đổi template)
              </button>
            )}
          </div>
        </div>
      ) : doc ? (
        <div className="report-entry-content">
          <ReportMetaForm
            metadata={doc.metadata}
            dateStr={dateStr}
            onDateChange={handleDateChange}
            onMetadataChange={handleMetadataChange}
            isExisting={!!doc.isExisting}
            disabled={isSaving}
            showErrors={showRequiredErrors}
          />

          <ReportSectionsEditor
            doc={doc}
            disabled={isSaving}
            isUploadingPhoto={isUploadingPhoto}
            showErrors={showRequiredErrors}
            onSectionChange={handleSectionChange}
            onUploadPhotos={handleUploadPhotos}
            onDeletePhoto={handleDeletePhoto}
          />

          {isWeekly && doc.tasks && doc.tasks.length > 0 && (
            <ReportTaskTable
              tasks={doc.tasks}
              onTaskChange={handleTaskChange}
              disabled={isSaving}
            />
          )}
        </div>
      ) : (
        <EmptyState message="Không thể nạp dữ liệu mẫu báo cáo" />
      )}

      {/* 5. Print / Export PDF A4 Preview Modal */}
      {showPrintPreview && doc && (
        <ReportPrintPreviewModal doc={doc} onClose={() => setShowPrintPreview(false)} />
      )}
    </div>
  );
};
