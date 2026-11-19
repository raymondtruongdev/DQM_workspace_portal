import * as React from 'react';
import { RoleEntry, ToastState } from '../../../core/types';
import { ToastAlert } from '../../../core/components/layout/ToastAlert';
import { PhotoGroup, ReportDoc, ReportMetadata, ReportPhotoSection, ReportSection, SignedDocInfo } from '../types/report.types';
import { ReportDataService, ReportNumberingError } from '../services/ReportDataService';
import { IReportLock, ISignedTrail, ReportLockedError, canUploadSignedWhileLocked } from '../services/ReportRegistryService';
import { extractTemplateVersionDate, formatTemplateVersionDisplay } from '../utils/reportFileNaming';
import { buildSectionNumbering } from '../utils/reportSections';
import { recalculateAllRates } from '../utils/reportTaskTree';
import { getMissingRequiredFields, getFieldLabel, getMissingFormFields, getMissingSelectCells } from '../utils/reportValidation';
import { ReportMetaForm } from '../components/ReportMetaForm';
import { ReportTaskTable } from '../components/ReportTaskTable';
import { AskCatalogLink, CatalogAddResult } from '../components/ReportDynamicTable';
import { ReportCatalog, canAddToTable, catalogOptions, catalogWithLink, catalogWithRow, findDuplicate } from '../utils/reportCatalog';
import { ReportSectionsEditor } from '../components/ReportSectionsEditor';
import { ReportReadOnlyContext } from '../components/ReportReadOnlyContext';
import { ReportSignedDocument } from '../components/ReportSignedDocument';
import { getReportFolderName, isLayoutGroupId, photoSubfolder } from '../utils/reportFileNaming';
import { formatUnlockUntil } from '../utils/reportSettings';
import { ReportPrintPreviewModal } from '../components/ReportPrintPreviewModal';
import { Icon, PageHeader, LoadingState, EmptyState, ConfirmDialog } from '../../../core/components';
import { SelectOpenContext } from '../../../core/components/CreatableSelect';

/** "YYYY-MM-DD" -> "dd/MM" */
const tplDM = (d?: string): string => (d ? d.split('-').reverse().slice(0, 2).join('/') : '');
/** Hậu tố " · dd/MM" từ ngày trong tên phiên bản template ('' nếu không có). */
const tplVerDM = (v?: string): string => {
  const d8 = extractTemplateVersionDate(v);
  return d8.length === 8 ? ` · ${d8.slice(6, 8)}/${d8.slice(4, 6)}` : '';
};

export interface IReportEntryViewProps {
  role: RoleEntry;
  currentUserEmail: string;
  dataService: ReportDataService;
  /** true = portal đang bảo trì, không được ghi lên SharePoint. */
  onCheckMaintenance: () => Promise<boolean>;
  /** Ngày mở sẵn (YYYY-MM-DD); mặc định là hôm nay. */
  initialDate?: string;
  onBack: () => void;
  /** Về trang chủ portal (nút trong thông báo báo cáo bị khoá). */
  onGoHome: () => void;
  /** Được khoá/mở khoá (admin thật và đang xem bằng hồ sơ admin) báo cáo. */
  canManageLock?: boolean;
}

export const ReportEntryView: React.FC<IReportEntryViewProps> = ({
  role,
  currentUserEmail,
  dataService,
  onCheckMaintenance,
  initialDate,
  onBack,
  onGoHome,
  canManageLock = false
}) => {
  const [dateStr, setDateStr] = React.useState<string>(() => initialDate || dataService.today());
  const [doc, setDoc] = React.useState<ReportDoc | null>(null);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [isUploadingPhoto, setIsUploadingPhoto] = React.useState<boolean>(false);
  const [hasUnsavedChanges, setHasUnsavedChanges] = React.useState<boolean>(false);
  const [showRequiredErrors, setShowRequiredErrors] = React.useState<boolean>(false);
  const [showPrintPreview, setShowPrintPreview] = React.useState<boolean>(false);
  const [toast, setToast] = React.useState<ToastState | null>(null);
  const [catalog, setCatalog] = React.useState<ReportCatalog | undefined>(undefined);
  const [signedDoc, setSignedDoc] = React.useState<SignedDocInfo | undefined>(undefined);
  const [isSignedLoading, setIsSignedLoading] = React.useState<boolean>(false);
  const [isUploadingSigned, setIsUploadingSigned] = React.useState<boolean>(false);
  const [isRemovingSigned, setIsRemovingSigned] = React.useState<boolean>(false);
  const [removeSignedOpen, setRemoveSignedOpen] = React.useState<boolean>(false);
  // Khoá báo cáo (đã upload bản ký hoặc admin khoá): người dùng chỉ xem/sao chép, không sửa/lưu; admin bật/tắt
  const [lock, setLock] = React.useState<IReportLock | undefined>(undefined);
  const [lockConfirm, setLockConfirm] = React.useState<'lock' | 'unlock' | null>(null);
  const [isTogglingLock, setIsTogglingLock] = React.useState<boolean>(false);
  /** Hiện hộp thoại "Báo cáo đã bị khoá" (khi đang làm dở mà admin khoá, lúc lưu bị từ chối). */
  const [lockedNotice, setLockedNotice] = React.useState<boolean>(false);
  const isLocked = Boolean(lock && lock.locked);
  // Số báo cáo theo ngày bắt đầu của nhà ga (DRPUnitConfig): chưa cấu hình / ngày trước ngày bắt đầu thì xem được nhưng không lưu được
  const [numberingIssue, setNumberingIssue] = React.useState<ReportNumberingError | undefined>(undefined);
  const [unitGeoTag, setUnitGeoTag] = React.useState<string>('');
  // Khoá nhưng báo cáo đã Lưu và chưa ký: người dùng vẫn upload được bản đã ký (nội dung vẫn chỉ xem)
  const canSignWhileLocked = isLocked && canUploadSignedWhileLocked(lock);

  /** Lỗi khi ghi: báo cáo đã bị khoá (admin vừa khoá hoặc đã ký) thì hiện hộp thoại và chuyển sang chỉ xem; lỗi khác hiện toast. */
  const handleWriteError = (err: unknown, fallback: string): void => {
    if (err instanceof ReportLockedError) {
      setLock((prev) => ({ exists: true, locked: true, canToggle: prev ? prev.canToggle : false }));
      setLockedNotice(true);
      return;
    }
    setToast({ type: 'error', text: err instanceof Error ? err.message : fallback });
  };

  // Load report data for selected date
  const loadData = React.useCallback(
    async (targetDate: string): Promise<void> => {
      setIsLoading(true);
      try {
        const result = await dataService.loadReport(role.docId, role.unit, targetDate, currentUserEmail, role.siteManager);
        setDoc(result.doc);
        setNumberingIssue(result.numberingIssue);
        setUnitGeoTag(result.unitGeoTag || '');
        setShowRequiredErrors(false);
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

  // Danh mục (Activity, cấu kiện, tài nguyên, vật tư…) của đơn vị (cho các cột select:@…); lỗi thì để trống, form vẫn dùng được
  React.useEffect(() => {
    let cancelled = false;
    dataService
      .getCatalog(role.unit)
      .then((c) => {
        if (cancelled) return;
        setCatalog(c);
        if (c.warnings.length > 0) setToast({ type: 'error', text: `Danh mục: ${c.warnings[0]}` });
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, [dataService, role.unit]);

  // Đọc lại danh mục khi người dùng bấm vào một dropdown (thấy ngay giá trị người khác vừa thêm). Tối đa 1 lần / 5 giây,
  // không chạy chồng; chỉ cập nhật khi danh mục thật sự đổi (tránh dựng lại danh sách đang mở); bỏ kết quả cũ nếu giữa chừng có thêm mục.
  const catalogRef = React.useRef<ReportCatalog | undefined>(undefined);
  catalogRef.current = catalog;
  const catalogStamp = React.useRef({ busy: false, last: 0, version: 0 });
  const refreshCatalog = React.useCallback((): void => {
    const st = catalogStamp.current;
    if (st.busy || Date.now() - st.last < 5000 || !catalogRef.current) return;
    st.busy = true;
    st.last = Date.now();
    const version = st.version;
    dataService
      .getCatalog(role.unit)
      .then((c) => {
        if (version !== st.version) return;
        const cur = catalogRef.current;
        if (cur && JSON.stringify(cur.tables) !== JSON.stringify(c.tables)) setCatalog(c);
      })
      .catch(() => {})
      .then(
        () => {
          st.busy = false;
        },
        () => {
          st.busy = false;
        }
      );
  }, [dataService, role.unit]);

  React.useEffect(() => {
    loadData(dateStr).catch(() => {});
  }, [loadData, dateStr]);

  // Còn thay đổi chưa lưu thì trình duyệt hỏi trước khi rời trang (tải lại, đóng tab, bị điều hướng nhầm sang file/ảnh)
  React.useEffect(() => {
    if (!hasUnsavedChanges) return undefined;
    const onBeforeUnload = (e: BeforeUnloadEvent): void => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [hasUnsavedChanges]);

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
    // Dạng hàm: nhiều lần đổi liên tiếp trong một sự kiện (vd thêm đơn vị mới + chọn nó) không ghi đè nhau
    setDoc((prev) => (prev ? { ...prev, metadata: { ...prev.metadata, [field]: value } } : prev));
    setHasUnsavedChanges(true);
  };

  // Dynamic table handlers
  /**
   * Người dùng chọn "Thêm mới…" ở cột danh mục: chống trùng (mục đã có thì chọn mục sẵn có), chưa có thì ghi vào list DRPCatalogItems
   * để mọi người cùng thấy; mục có mã thì hỏi thêm giá trị liên kết. Trả `{ cancelled }` khi người dùng huỷ, undefined khi không ghi được
   * (giá trị chỉ lưu trong báo cáo này).
   */
  const handleAddCatalogItem = async (
    ref: string,
    value: string,
    byLink: boolean,
    askLink?: AskCatalogLink
  ): Promise<CatalogAddResult | undefined> => {
    const clean = value.trim().replace(/\s+/g, ' ');
    if (!clean || !catalog) return undefined;
    const dup = findDuplicate(catalog, ref, byLink, clean);
    if (dup) {
      const existing = byLink ? dup.link : dup.id;
      setToast({ type: 'success', text: `"${existing}" đã có trong danh mục, đã chọn giá trị có sẵn.` });
      return { value: existing, catalog };
    }
    if (!canAddToTable(catalog, ref)) return undefined;
    if (await onCheckMaintenance()) return undefined;
    // Mục mới theo mã: hỏi luôn giá trị liên kết (vd Activity của cấu kiện), bỏ trống được; huỷ thì không thêm
    let link = '';
    let field: 'note' | 'nameVn' | undefined;
    if (!byLink && askLink) {
      const asked = await askLink();
      if (!asked) return { cancelled: true };
      link = asked.link.trim();
      field = asked.field;
    }
    try {
      const res = await dataService.addCatalogItem(role.unit, { table: ref, code: byLink ? '' : clean, link: byLink ? clean : field ? '' : link, ...(field && link ? { [field]: link } : {}) });
      const saved = byLink ? res.row.link : res.row.code;
      const next = catalogWithRow(catalog, role.unit, res.row);
      catalogStamp.current.version++;
      setCatalog(next);
      setToast({
        type: 'success',
        text: res.created ? `Đã thêm "${saved}" vào danh mục dùng chung.` : `"${saved}" đã có trong danh mục, đã chọn giá trị có sẵn.`
      });
      return { value: saved, catalog: next };
    } catch (err: unknown) {
      console.warn('[Report] Could not add catalog item:', err);
      setToast({ type: 'error', text: 'Chưa thêm được vào danh mục dùng chung (list DRPCatalogItems); giá trị chỉ lưu trong báo cáo này.' });
      return undefined;
    }
  };

  /**
   * Người dùng chọn lại Activity cho cấu kiện do người dùng thêm (`AddedByAdmin = 0`): ghi liên kết vào danh mục dùng chung
   * để mọi người sau thấy cấu kiện đã liên kết. Mục admin tạo sẵn thì không cập nhật. Lỗi (thiếu quyền sửa…) thì chỉ cảnh báo,
   * giá trị vẫn lưu trong báo cáo này.
   */
  const handleUpdateCatalogLink = async (ref: string, code: string, link: string): Promise<void> => {
    if (!catalog) return;
    if (await onCheckMaintenance()) return;
    try {
      const ok = await dataService.updateCatalogLink(role.unit, ref, code, link);
      if (!ok) return;
      catalogStamp.current.version++;
      setCatalog(catalogWithLink(catalog, ref, code, link));
      setToast({ type: 'success', text: `Đã cập nhật liên kết của "${code}" trong danh mục dùng chung.` });
    } catch (err: unknown) {
      console.warn('[Report] Could not update catalog link:', err);
      setToast({ type: 'error', text: 'Chưa cập nhật được liên kết trong danh mục dùng chung (cần quyền sửa mục do bạn thêm); giá trị chỉ lưu trong báo cáo này.' });
    }
  };

  /** Thêm đơn vị thi công mới vào danh mục dùng chung `subcontractor`; trả về giá trị để chọn. */
  // Tọa độ của nhà ga: lấy từ danh mục `location` (Code = tọa độ, NameVN = mã nhà ga), người dùng không chọn/sửa

  const handleAddSubcontractor = async (value: string): Promise<string | undefined> => {
    const res = await handleAddCatalogItem('subcontractor', value, false);
    return res && !res.cancelled ? res.value : undefined;
  };

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

  const isWeekly = role.docId.toLowerCase().includes('week');

  // Tự điền Tọa độ theo nhà ga (cột GeoTag của DRPUnitConfig; không đánh dấu "chưa lưu": giá trị luôn suy ra từ cấu hình); báo cáo đã khoá giữ giá trị đã lưu
  React.useEffect(() => {
    if (!doc || isLocked || !unitGeoTag || doc.metadata.geoTag === unitGeoTag) return;
    setDoc((prev) => (prev ? { ...prev, metadata: { ...prev.metadata, geoTag: unitGeoTag } } : prev));
  }, [doc, isLocked, unitGeoTag]);

  // File PDF đã ký của ngày đang mở (chỉ khi báo cáo đã lưu trên SharePoint)
  const isReportSaved = Boolean(doc && doc.isExisting);
  React.useEffect(() => {
    if (isWeekly || !isReportSaved) {
      setSignedDoc(undefined);
      return;
    }
    let alive = true;
    setIsSignedLoading(true);
    dataService
      .getSignedDocument(role.unit, role.docId, dateStr)
      .then((info) => {
        if (alive) {
          setSignedDoc(info);
          setIsSignedLoading(false);
        }
      })
      .catch((err) => {
        console.warn('Could not read signed document:', err);
        if (alive) setIsSignedLoading(false);
      });
    return () => {
      alive = false;
    };
  }, [dataService, role.unit, role.docId, dateStr, isReportSaved, isWeekly]);

  // Upload file PDF đã ký vào thư mục báo cáo của ngày (cùng thư mục với file .json)
  const handleUploadSigned = async (file: File): Promise<void> => {
    if (!doc || !doc.isExisting || hasUnsavedChanges) return;
    if (await onCheckMaintenance()) return;
    setIsUploadingSigned(true);
    try {
      const buf = await file.arrayBuffer();
      const res = await dataService.uploadSignedDocument(
        role.unit,
        role.docId,
        dateStr,
        buf,
        doc.metadata.reportNo,
        doc.metadata.email,
        doc.metadata.docCode,
        currentUserEmail
      );
      setSignedDoc(res.info);
      setToast(
        res.registryWarning
          ? { type: 'error', text: res.registryWarning }
          : { type: 'success', text: 'Đã upload bản đã ký. Báo cáo ngày này đã hoàn tất.' }
      );
    } catch (err: unknown) {
      console.error('Error uploading signed document:', err);
      handleWriteError(err, 'Lỗi khi upload file đã ký lên SharePoint');
    } finally {
      setIsUploadingSigned(false);
    }
  };

  // Trạng thái khoá đọc từ sổ DRPRegistry mỗi lần mở ngày, và đọc lại sau khi lưu/upload/xoá bản ký/bật tắt khoá
  React.useEffect(() => {
    if (isWeekly) {
      setLock(undefined);
      return;
    }
    let alive = true;
    dataService
      .getLockState(role.unit, role.docId, dateStr)
      .then((l) => {
        if (alive) setLock(l);
      })
      .catch(() => undefined);
    return () => {
      alive = false;
    };
  }, [dataService, role.unit, role.docId, dateStr, isReportSaved, isWeekly, signedDoc, isRemovingSigned, isUploadingSigned, isTogglingLock]);

  // Mở khoá có hạn: đếm ngược, cảnh báo khi còn ≤ 10 phút, và đọc lại khoá đúng lúc hết hạn để form chuyển sang chỉ xem
  const unlockUntil = !isLocked && lock && lock.unlockUntil ? lock.unlockUntil : undefined;
  const [nowMs, setNowMs] = React.useState<number>(() => dataService.now());
  React.useEffect(() => {
    if (!unlockUntil) return;
    setNowMs(dataService.now());
    const expiry = Date.parse(unlockUntil);
    const tick = window.setInterval(() => setNowMs(dataService.now()), 30000);
    const wait = Math.max(0, Math.min(expiry - dataService.now() + 1000, 2147483000));
    const expire = window.setTimeout(() => {
      dataService
        .getLockState(role.unit, role.docId, dateStr)
        .then((l) => setLock(l))
        .catch(() => undefined);
    }, wait);
    return () => {
      window.clearInterval(tick);
      window.clearTimeout(expire);
    };
  }, [unlockUntil, dataService, role.unit, role.docId, dateStr]);
  const unlockMinutesLeft = unlockUntil ? Math.max(0, Math.ceil((Date.parse(unlockUntil) - nowMs) / 60000)) : 0;

  // Admin bật/tắt khoá báo cáo của ngày đang mở
  const handleToggleLock = async (): Promise<void> => {
    const target = lockConfirm === 'lock';
    setLockConfirm(null);
    if (await onCheckMaintenance()) return;
    setIsTogglingLock(true);
    try {
      await dataService.setReportLocked(role.unit, role.docId, dateStr, target);
      setToast({
        type: 'success',
        text: target ? 'Đã khoá báo cáo: người dùng chỉ xem được.' : 'Đã mở khoá báo cáo: người dùng có thể chỉnh sửa lại.'
      });
    } catch (err: unknown) {
      console.error('Error toggling report lock:', err);
      setToast({ type: 'error', text: err instanceof Error ? err.message : 'Chưa bật/tắt được khoá báo cáo.' });
    } finally {
      setIsTogglingLock(false);
    }
  };

  // Dấu vết bản đã ký (người upload, lần xoá gần nhất) đọc từ sổ DRPRegistry; đọc lại sau mỗi lần upload/xoá
  const [signedTrail, setSignedTrail] = React.useState<ISignedTrail | undefined>(undefined);
  React.useEffect(() => {
    if (isWeekly || !isReportSaved) {
      setSignedTrail(undefined);
      return;
    }
    let alive = true;
    dataService
      .getSignedTrail(role.unit, role.docId, dateStr)
      .then((r) => {
        if (alive) setSignedTrail(r);
      })
      .catch(() => undefined);
    return () => {
      alive = false;
    };
  }, [dataService, role.unit, role.docId, dateStr, isReportSaved, isWeekly, signedDoc, isRemovingSigned, isUploadingSigned]);

  // Xoá file PDF đã ký: xoá file, đưa ngày về "Chưa upload bản ký", ghi người xoá + thời điểm gần nhất vào sổ
  const handleRemoveSigned = async (): Promise<void> => {
    setRemoveSignedOpen(false);
    if (!doc || !doc.isExisting) return;
    if (await onCheckMaintenance()) return;
    setIsRemovingSigned(true);
    try {
      const res = await dataService.removeSignedDocument(role.unit, role.docId, dateStr, currentUserEmail);
      setSignedDoc(undefined);
      setToast(
        res.warning
          ? { type: 'error', text: res.warning }
          : { type: 'success', text: 'Đã xoá file đã ký. Báo cáo ngày này quay về "Chưa upload bản ký".' }
      );
    } catch (err: unknown) {
      console.error('Error removing signed document:', err);
      handleWriteError(err, 'Lỗi khi xoá file đã ký trên SharePoint');
    } finally {
      setIsRemovingSigned(false);
    }
  };

  // Thư mục ảnh của báo cáo: data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/photos (mỗi bộ phận một thư mục con)
  const photosFolder = dataService.getPhotosFolderPath(role.unit, role.docId, dateStr);

  const updatePhotoSection = (
    sectionKey: string,
    update: (s: ReportPhotoSection) => ReportPhotoSection
  ): void => {
    setDoc((prev) =>
      prev
        ? {
            ...prev,
            sections: prev.sections.map((x) => (x.kind === 'photo' && x.key === sectionKey ? update(x) : x))
          }
        : null
    );
    setHasUnsavedChanges(true);
  };

  const findPhotoGroup = (sectionKey: string, groupId: string): PhotoGroup | undefined => {
    const sec = doc?.sections.filter((x) => x.kind === 'photo' && x.key === sectionKey)[0];
    return sec && sec.kind === 'photo' ? sec.groups.filter((g) => g.id === groupId)[0] : undefined;
  };

  // Hình Layout dùng chung thư mục photos/layout/: hàng khác còn dùng cùng tên file thì chỉ gỡ khỏi hàng này, giữ file trên SharePoint
  const isLayoutPhotoUsedElsewhere = (sectionKey: string, groupId: string, name: string): boolean => {
    if (!isLayoutGroupId(groupId)) return false;
    const sec = doc?.sections.filter((x) => x.kind === 'photo' && x.key === sectionKey)[0];
    const k = name.toLowerCase();
    return !!sec && sec.kind === 'photo' && sec.groups.some((g) => g.id !== groupId && g.files.some((f) => f.toLowerCase() === k));
  };

  // Hộp thoại hình Layout trùng tên: Thay thế hoặc Huỷ (hỏi lần lượt từng file trùng)
  const [photoReplaceAsk, setPhotoReplaceAsk] = React.useState<{ name: string; resolve: (replace: boolean) => void } | null>(null);
  const askReplacePhoto = (name: string): Promise<boolean> => new Promise<boolean>((resolve) => setPhotoReplaceAsk({ name, resolve }));

  // Tải ảnh của một bộ phận lên data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/photos/<groupId>/
  const handleUploadGroupPhotos = async (sectionKey: string, groupId: string, files: FileList): Promise<void> => {
    if (!doc || files.length === 0) return;
    if (await onCheckMaintenance()) return;
    setIsUploadingPhoto(true);
    try {
      const names: string[] = [];
      let skipped = 0;
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        // Hình Layout giữ nguyên tên và nằm chung photos/layout/: trùng tên file đã có thì hỏi Thay thế / Huỷ
        if (isLayoutGroupId(groupId)) {
          const chk = await dataService.checkPhoto(role.unit, role.docId, dateStr, groupId, file.name);
          if (chk.exists && !(await askReplacePhoto(chk.name))) {
            skipped += 1;
            continue;
          }
        }
        const buf = await file.arrayBuffer();
        names.push(await dataService.uploadPhoto(role.unit, role.docId, dateStr, groupId, file.name, buf));
      }
      updatePhotoSection(sectionKey, (s) => ({
        ...s,
        groups: s.groups.map((g) => {
          if (g.id !== groupId) return g;
          const lower = g.files.map((f) => f.toLowerCase());
          return { ...g, files: g.files.concat(names.filter((n) => lower.indexOf(n.toLowerCase()) === -1)) };
        })
      }));
      setToast({
        type: 'success',
        text:
          names.length === 0
            ? 'Đã huỷ, không có hình nào được tải lên.'
            : `Đã tải lên ${names.length} ảnh${skipped > 0 ? ` (bỏ qua ${skipped} hình trùng tên)` : ''}.`
      });
    } catch (err: unknown) {
      console.error('Error uploading photo:', err);
      handleWriteError(err, 'Lỗi khi tải ảnh lên SharePoint');
    } finally {
      setIsUploadingPhoto(false);
    }
  };

  // Xoá một ảnh của bộ phận (xoá cả file trên SharePoint)
  const handleDeleteGroupPhoto = async (sectionKey: string, groupId: string, name: string): Promise<void> => {
    if (!doc) return;
    if (!window.confirm(`Bạn có chắc chắn muốn xoá ảnh "${name}" khỏi SharePoint?`)) return;
    if (await onCheckMaintenance()) return;
    setIsUploadingPhoto(true);
    try {
      if (!isLayoutPhotoUsedElsewhere(sectionKey, groupId, name)) await dataService.deletePhoto(`${photosFolder}/${photoSubfolder(groupId)}/${name}`);
      updatePhotoSection(sectionKey, (s) => ({
        ...s,
        groups: s.groups.map((g) => (g.id === groupId ? { ...g, files: g.files.filter((f) => f !== name) } : g))
      }));
      setToast({ type: 'success', text: `Đã xoá ảnh "${name}".` });
    } catch (err: unknown) {
      console.error('Error deleting photo:', err);
      const msg = err instanceof Error ? err.message : 'Lỗi khi xoá ảnh trên SharePoint';
      setToast({ type: 'error', text: msg });
    } finally {
      setIsUploadingPhoto(false);
    }
  };

  // Xoá cả bộ phận: xoá các ảnh của bộ phận trên SharePoint rồi bỏ nhóm khỏi báo cáo
  const handleDeletePhotoGroup = async (sectionKey: string, groupId: string): Promise<void> => {
    const group = findPhotoGroup(sectionKey, groupId);
    if (!group) return;
    const label = group.department || 'chưa đặt tên';
    const msg =
      group.files.length > 0
        ? `Xoá bộ phận "${label}" và ${group.files.length} ảnh của bộ phận này khỏi SharePoint?`
        : `Xoá bộ phận "${label}"?`;
    if (!window.confirm(msg)) return;
    if (await onCheckMaintenance()) return;
    setIsUploadingPhoto(true);
    try {
      for (let i = 0; i < group.files.length; i++) {
        if (!isLayoutPhotoUsedElsewhere(sectionKey, groupId, group.files[i])) await dataService.deletePhoto(`${photosFolder}/${photoSubfolder(groupId)}/${group.files[i]}`);
      }
      updatePhotoSection(sectionKey, (s) => ({ ...s, groups: s.groups.filter((g) => g.id !== groupId) }));
    } catch (err: unknown) {
      console.error('Error deleting photo group:', err);
      const m = err instanceof Error ? err.message : 'Lỗi khi xoá ảnh trên SharePoint';
      setToast({ type: 'error', text: m });
    } finally {
      setIsUploadingPhoto(false);
    }
  };

  // PDF đính kèm của dòng sự kiện: pdf/eventlog/<mã sự kiện>_<tên file>.pdf; danh sách tên lưu trong `files` của dòng khi bấm Lưu
  const eventPdfFolder = dataService.getEventPdfFolderPath(role.unit, role.docId, dateStr);

  const updateRowFiles = (sectionKey: string, rowId: string, update: (files: string[]) => string[]): void => {
    setDoc((prev) =>
      prev
        ? {
            ...prev,
            sections: prev.sections.map((x) =>
              x.kind === 'table' && x.key === sectionKey
                ? { ...x, rows: x.rows.map((r) => (r.id === rowId ? { ...r, files: update(r.files || []) } : r)) }
                : x
            )
          }
        : null
    );
    setHasUnsavedChanges(true);
  };

  // Số dòng khác (kể cả mục khác) còn dùng file PDF này: file giữ tên gốc nên có thể dùng chung, chỉ xoá file khi không ai còn dùng
  const isPdfUsedElsewhere = (sectionKey: string, rowId: string, name: string): boolean => {
    const lower = name.toLowerCase();
    return Boolean(
      doc &&
        doc.sections.some(
          (x) =>
            x.kind === 'table' &&
            x.rows.some((r) => !(x.key === sectionKey && r.id === rowId) && (r.files || []).some((f) => f.toLowerCase() === lower))
        )
    );
  };

  // Hộp thoại "file đã tồn tại": Thay thế hoặc Huỷ (hỏi lần lượt từng file trùng tên)
  const [pdfReplaceAsk, setPdfReplaceAsk] = React.useState<{ name: string; resolve: (replace: boolean) => void } | null>(null);
  const askReplacePdf = (name: string): Promise<boolean> => new Promise<boolean>((resolve) => setPdfReplaceAsk({ name, resolve }));

  const handleUploadRowPdf = async (sectionKey: string, rowId: string, files: FileList): Promise<void> => {
    if (!doc || files.length === 0) return;
    const picked: File[] = [];
    for (let i = 0; i < files.length; i++) picked.push(files[i]);
    const bad = picked.filter((f) => !/\.pdf$/i.test(f.name) && f.type !== 'application/pdf');
    if (bad.length > 0) {
      setToast({ type: 'error', text: `Chỉ nhận file PDF (bỏ qua: ${bad.map((f) => f.name).join(', ')}).` });
      return;
    }
    if (await onCheckMaintenance()) return;
    try {
      const names: string[] = [];
      let skipped = 0;
      for (const file of picked) {
        // Giữ nguyên tên file; trùng tên file đã có trong pdf/eventlog thì hỏi Thay thế / Huỷ
        const chk = await dataService.checkEventPdf(role.unit, role.docId, dateStr, file.name);
        if (chk.exists && !(await askReplacePdf(chk.name))) {
          skipped += 1;
          continue;
        }
        const buf = await file.arrayBuffer();
        names.push(await dataService.uploadEventPdf(role.unit, role.docId, dateStr, file.name, buf));
      }
      if (names.length > 0) {
        updateRowFiles(sectionKey, rowId, (cur) => {
          const lower = cur.map((f) => f.toLowerCase());
          return cur.concat(names.filter((n) => lower.indexOf(n.toLowerCase()) === -1));
        });
      }
      if (names.length === 0) setToast({ type: 'success', text: 'Đã huỷ, không có file nào được đính kèm.' });
      else
        setToast({
          type: 'success',
          text: `Đã đính kèm ${names.length} file PDF${skipped > 0 ? ` (bỏ qua ${skipped} file trùng tên)` : ''}. Bấm Lưu báo cáo để ghi vào báo cáo.`
        });
    } catch (err: unknown) {
      console.error('Error uploading event PDF:', err);
      handleWriteError(err, 'Lỗi khi tải file PDF lên SharePoint');
    }
  };

  const handleDeleteRowPdf = async (sectionKey: string, rowId: string, name: string): Promise<void> => {
    if (await onCheckMaintenance()) return;
    try {
      // Dòng khác còn dùng file này thì chỉ gỡ khỏi dòng, giữ file
      if (!isPdfUsedElsewhere(sectionKey, rowId, name)) await dataService.deleteEventPdf(role.unit, role.docId, dateStr, name);
      updateRowFiles(sectionKey, rowId, (cur) => cur.filter((f) => f !== name));
      setToast({ type: 'success', text: `Đã xoá file "${name}". Bấm Lưu báo cáo để ghi vào báo cáo.` });
    } catch (err: unknown) {
      console.error('Error deleting event PDF:', err);
      handleWriteError(err, 'Lỗi khi xoá file PDF trên SharePoint');
    }
  };

  // Xoá cả dòng sự kiện: xoá luôn các PDF của dòng (trừ file dòng khác còn dùng); lỗi thì ném ra để bảng giữ lại dòng
  const handlePurgeRowPdf = async (sectionKey: string, rowId: string, names: string[]): Promise<void> => {
    if (await onCheckMaintenance()) throw new Error('Portal đang bảo trì');
    try {
      for (const name of names) {
        if (!isPdfUsedElsewhere(sectionKey, rowId, name)) await dataService.deleteEventPdf(role.unit, role.docId, dateStr, name);
      }
    } catch (err: unknown) {
      console.error('Error deleting event PDFs:', err);
      handleWriteError(err, 'Lỗi khi xoá file PDF trên SharePoint');
      throw err;
    }
  };

  // Load new template when existing / previous-day report uses an outdated template
  const handleLoadLatestTemplate = async (): Promise<void> => {
    if (!doc) return;
    // Chế độ bảo trì (Shutdown Server): không cho áp dụng template mới
    if (await onCheckMaintenance()) return;
    setIsLoading(true);
    try {
      const next = await dataService.applyLatestTemplate(doc, currentUserEmail, role.siteManager);
      setDoc(next);
      setHasUnsavedChanges(true);
      setToast({
        type: 'success',
        text: `Đã nạp biểu mẫu mới nhất (${formatTemplateVersionDisplay(
          next.metadata.templateVersion
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
  const ensureRequiredFilled = (action: string, requireSelects = false): boolean => {
    if (!doc) return false;
    const missing = getMissingRequiredFields(doc.metadata);
    const missingForm = getMissingFormFields(doc);
    const missingSelects = requireSelects ? getMissingSelectCells(doc) : [];
    if (missing.length === 0 && missingForm.length === 0 && missingSelects.length === 0) return true;
    setShowRequiredErrors(true);
    const labels = missing.map(getFieldLabel).concat(missingForm.map((m) => m.label));
    const shownSelects = missingSelects.slice(0, 5).map((m) => m.label);
    const moreSelects = missingSelects.length > 5 ? ` và ${missingSelects.length - 5} ô khác` : '';
    const parts: string[] = [];
    if (labels.length > 0) parts.push(`Vui lòng nhập: ${labels.join(', ')}.`);
    if (missingSelects.length > 0) parts.push(`Cần chọn đủ các ô còn "Chọn…": ${shownSelects.join('; ')}${moreSelects}.`);
    setToast({
      type: 'error',
      text: `Không thể ${action}. ${parts.join(' ')}`
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
    if (ensureRequiredFilled('xem bản in / xuất PDF', true)) setShowPrintPreview(true);
  };

  // Lưu báo cáo: ghi .json + .html độc lập
  const handleSave = async (): Promise<void> => {
    if (!doc || numberingIssue || !ensureRequiredFilled('lưu báo cáo')) return;
    if (await onCheckMaintenance()) return;
    setIsSaving(true);
    try {
      const res = await dataService.saveReport(doc, currentUserEmail);
      // Số báo cáo/mã tài liệu có thể đã được cấp lại theo sổ DRPRegistry lúc lưu
      // Nền bộ đếm mã tự tăng cũng lấy lại theo bộ đếm vừa lưu (hàng đã lưu rồi xoá không bị cấp lại trong cùng lần mở)
      setDoc((prev) =>
        prev ? { ...prev, isExisting: true, eTag: res.eTag, metadata: res.metadata, idCounters: res.idCounters || prev.idCounters } : null
      );
      setHasUnsavedChanges(false);
      const warning = res.registryWarning || res.htmlWarning;
      setToast(
        warning
          ? { type: 'error', text: warning }
          : {
              type: 'success',
              text: `Đã lưu báo cáo số ${res.metadata.reportNo} vào thư viện SharePoint (${role.unit} - ${dateStr})!`
            }
      );
    } catch (err: unknown) {
      console.error('Error saving report:', err);
      handleWriteError(err, 'Lỗi khi lưu báo cáo lên SharePoint');
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

  // Báo cáo đang khoá thì không bắt Apply template (chỉ xem)
  const tplRequired = Boolean(doc && doc.templateUpdateRequired) && !isLocked;
  const editDisabled = isSaving || isLocked;

  const reportMainTitle = doc?.metadata.docTitle || (isWeekly ? 'Báo cáo tuần' : 'Báo cáo ngày');

  return (
    <SelectOpenContext.Provider value={refreshCatalog}>
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
                  : 'Tạo mới từ template'}
              </span>
            )}
            {!isWeekly && doc?.isExisting && !isSignedLoading && (
              <span className={`badge ${signedDoc ? 'badge--success' : 'badge--warning'}`}>
                <Icon name={signedDoc ? 'check-circle' : 'file-up'} size={16} />
                {signedDoc ? 'Hoàn tất (đã có bản ký)' : 'Chưa upload bản ký'}
              </span>
            )}
            {isLocked && (
              <span
                className={`badge ${canSignWhileLocked ? 'badge--info' : 'badge--danger'}`}
                title={canSignWhileLocked ? 'Nội dung đã khoá (chỉ xem) nhưng bạn vẫn upload được bản đã ký' : 'Báo cáo đã khoá: chỉ xem, không chỉnh sửa được'}
              >
                <Icon name="lock" size={16} />
                {canSignWhileLocked ? 'Khoá nội dung · vẫn upload bản ký' : 'Đã khoá'}
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
          doc && tplRequired ? undefined : (
          <>
            {hasUnsavedChanges && (
              <span className="unsaved-badge" title="Có dữ liệu chưa lưu">
                Chưa lưu
              </span>
            )}
            {canManageLock && lock && !isWeekly && (
              <button
                type="button"
                className={`btn btn--secondary report-lock-toggle report-lock-toggle--${isLocked ? 'locked' : 'open'}`}
                onClick={() => setLockConfirm(isLocked ? 'unlock' : 'lock')}
                disabled={isLoading || isSaving || isTogglingLock}
                title={isLocked ? 'Đang khoá (người dùng chỉ xem). Bấm để mở khoá cho phép chỉnh sửa lại' : 'Đang mở (người dùng sửa được). Bấm để khoá: người dùng chỉ xem được'}
                aria-label={isLocked ? 'Trạng thái LOCKED, bấm để mở khoá' : 'Trạng thái OPEN, bấm để khoá'}
              >
                <Icon name={isLocked ? 'lock' : 'lock-open'} />
                <span>{isLocked ? 'LOCKED' : lock && lock.unlockUntil ? `OPEN · đến ${formatUnlockUntil(lock.unlockUntil)}` : 'OPEN'}</span>
              </button>
            )}
            <button
              type="button"
              className="btn btn--secondary"
              onClick={handleOpenPrint}
              disabled={isLoading || !doc || tplRequired}
              title="Xem bản in A4 và xuất PDF"
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
            {!isLocked && (
              <button
                type="button"
                className="btn btn--primary"
                onClick={handleSave}
                disabled={isLoading || isSaving || !doc || tplRequired || isLocked || Boolean(numberingIssue)}
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
            )}
          </>
          )
        }
      />

      {/* 2. Quick Section Navigation Pills */}
      {!isLoading && doc && !tplRequired && (
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
          {!isWeekly && (
            <button type="button" className="chip" onClick={() => scrollToSection('report-sec-signed')}>
              Bản đã ký
            </button>
          )}
        </div>
      )}

      {numberingIssue && !isLoading && (
        <div className="card report-lock-note" role="alert">
          <Icon name="warning" size={20} />
          <div className="report-lock-note__text">
            <strong>
              {numberingIssue.reason === 'beforeStart' ? 'Ngày này chưa được bắt đầu báo cáo, không lưu được.' : 'Chưa đánh được số báo cáo, không lưu được.'}
            </strong>
            <span> {numberingIssue.message}</span>
          </div>
        </div>
      )}

      {isLocked && !isLoading && (
        <div className={`card report-lock-note${canSignWhileLocked ? ' report-lock-note--sign' : ''}`} role="status">
          <Icon name={canSignWhileLocked ? 'file-up' : 'lock'} size={20} />
          <div className="report-lock-note__text">
            {canSignWhileLocked ? (
              <>
                <strong>Nội dung báo cáo đã khoá, chỉ xem. Bạn vẫn upload được bản đã ký.</strong>
                <span>
                  {' '}
                  Không sửa được dữ liệu đã lưu. Nếu báo cáo đã được in và ký, hãy upload file PDF đã ký ở mục BẢN ĐÃ KÝ (cuối trang).
                  {canManageLock ? ' Admin: bấm nút LOCKED ở trên để mở khoá nếu cần sửa nội dung.' : ' Cần sửa nội dung thì liên hệ admin để mở khoá.'}
                </span>
              </>
            ) : (
              <>
                <strong>Báo cáo này đã khoá, chỉ xem.</strong>
                <span>
                  {' '}
                  Bạn có thể xem và sao chép nội dung nhưng không chỉnh sửa được.
                  {canManageLock ? ' Bấm nút LOCKED ở trên để mở khoá, cho phép chỉnh sửa lại.' : ' Cần chỉnh sửa thì liên hệ admin để mở khoá.'}
                </span>
              </>
            )}
          </div>
          {canSignWhileLocked && (
            <button type="button" className="btn btn--primary btn--sm report-lock-note__action" onClick={() => scrollToSection('report-sec-signed')}>
              <Icon name="upload" />
              <span>Đến mục Bản đã ký</span>
            </button>
          )}
        </div>
      )}

      {unlockUntil && !isLoading && (
        <div className={`card report-lock-note report-lock-note--open${unlockMinutesLeft <= 10 ? ' report-lock-note--soon' : ''}`} role="status">
          <Icon name={unlockMinutesLeft <= 10 ? 'warning' : 'lock-open'} size={20} />
          <div className="report-lock-note__text">
            <strong>Báo cáo đang được mở khoá đến {formatUnlockUntil(unlockUntil)}</strong>
            <span>
              {' '}
              (còn {unlockMinutesLeft >= 60 ? `${Math.floor(unlockMinutesLeft / 60)} giờ ${unlockMinutesLeft % 60} phút` : `${unlockMinutesLeft} phút`}), sau đó tự khoá lại.
              Hãy Lưu trước khi hết giờ{unlockMinutesLeft <= 10 ? ': sắp hết hạn!' : '.'}
            </span>
          </div>
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
      ) : doc && tplRequired ? (
        <div className="report-tpl-notice" role="alert">
          <div className="report-tpl-notice__top">
            <div className="report-tpl-notice__icon" aria-hidden="true">
              <Icon name="warning" size={24} />
            </div>
            <div className="report-tpl-notice__text">
              <h3>Cần cập nhật template mới</h3>
              <p>
                {doc.outdatedSourceType === 'previous_day'
                  ? `Báo cáo ngày ${tplDM(doc.copiedFromDate)} dùng template cũ. Áp dụng bản mới nhất từ SharePoint `
                  : `Báo cáo ngày ${tplDM(dateStr)} dùng template cũ. Áp dụng bản mới nhất từ SharePoint `}
                <code>_template</code>
                {doc.outdatedSourceType === 'previous_day'
                  ? ` để tiếp tục nhập báo cáo ngày ${dateStr.split('-').reverse().join('/')}.`
                  : ' để tiếp tục cập nhật báo cáo.'}
              </p>
              <p>
                <strong>Lưu ý:</strong> dữ liệu đã nhập được giữ lại; mục, cột hoặc trường mới của template sẽ để trống
                hoặc điền giá trị mặc định, mục/cột đã bị bỏ khỏi template sẽ không còn trong báo cáo.
              </p>
            </div>
          </div>

          <div className="report-tpl-notice__compare">
            <div className="report-tpl-ver report-tpl-ver--old">
              <span className="report-tpl-ver__label">
                Hiện tại{tplVerDM(doc.outdatedTemplateVersion || doc.metadata.templateVersion)}
              </span>
              <strong className="report-tpl-ver__value">
                {(doc.outdatedTemplateVersion || doc.metadata.templateVersion || 'Chưa gắn phiên bản').replace(/\.json$/i, '')}
              </strong>
            </div>
            <span className="report-tpl-notice__arrow" aria-hidden="true">
              <Icon name="arrow-right" size={20} />
            </span>
            <div className="report-tpl-ver report-tpl-ver--new">
              <span className="report-tpl-ver__label">
                Mới nhất{tplVerDM(doc.latestTemplateInfo?.version)}
              </span>
              <strong className="report-tpl-ver__value">
                {doc.latestTemplateInfo?.fileName || doc.latestTemplateInfo?.version || '—'}
              </strong>
            </div>
          </div>

          <div className="report-tpl-notice__actions">
            <button type="button" className="btn btn--primary" onClick={handleLoadLatestTemplate}>
              <Icon name="check" size={20} />
              <span>Apply template mới để tiếp tục</span>
            </button>
            {doc.isExisting && (
              <button type="button" className="btn btn--secondary" onClick={handleKeepCurrentTemplate}>
                Xem dữ liệu cũ (không đổi template)
              </button>
            )}
          </div>
        </div>
      ) : doc ? (
        <ReportReadOnlyContext.Provider value={isLocked}>
        <div className="report-entry-content">
          <ReportMetaForm
            metadata={doc.metadata}
            dateStr={dateStr}
            onDateChange={handleDateChange}
            dateLocked={!isWeekly}
            onMetadataChange={handleMetadataChange}
            isExisting={!!doc.isExisting}
            subcontractorOptions={catalogOptions(catalog, 'subcontractor', false, true)}
            onAddSubcontractor={handleAddSubcontractor}
            disabled={editDisabled}
            showErrors={showRequiredErrors}
          />

          <ReportSectionsEditor
            doc={doc}
            disabled={editDisabled}
            isUploadingPhoto={isUploadingPhoto}
            catalog={catalog}
            onAddCatalogItem={handleAddCatalogItem}
            onUpdateCatalogLink={handleUpdateCatalogLink}
            onReserveAutoId={
              dataService.usesAutoIds(doc.docId)
                ? async (key, atLeast) => {
                    try {
                      return await dataService.reserveAutoId(doc.unit, doc.docId, key, atLeast);
                    } catch (err: unknown) {
                      // Dự phòng cục bộ ở bảng; báo rõ để admin kiểm tra list DRPCounters (mã có thể trùng với người khác)
                      setToast({
                        type: 'error',
                        text: `Chưa đặt chỗ được mã tự tăng ở list DRPCounters nên mã mới chỉ tính trên máy này và có thể trùng với người khác. Hãy báo Admin kiểm tra list DRPCounters (đã tạo chưa, quyền thêm/sửa, tên cột Title/Counters). ${err instanceof Error ? err.message.slice(0, 160) : ''}`
                      });
                      throw err;
                    }
                  }
                : undefined
            }
            eventPdfFolder={eventPdfFolder}
            onUploadRowPdf={handleUploadRowPdf}
            onDeleteRowPdf={handleDeleteRowPdf}
            onPurgeRowPdf={handlePurgeRowPdf}
            showErrors={showRequiredErrors}
            onSectionChange={handleSectionChange}
            photosFolder={photosFolder}
            onUploadGroupPhotos={handleUploadGroupPhotos}
            onDeleteGroupPhoto={handleDeleteGroupPhoto}
            onDeletePhotoGroup={handleDeletePhotoGroup}
          />

          {!isWeekly && (
            <ReportSignedDocument
              folderLabel={`data/${role.unit}/${getReportFolderName(role.docId, role.unit, dateStr)}/`}
              highlight={canSignWhileLocked && !signedDoc}
              canUpload={Boolean(doc.isExisting) && !hasUnsavedChanges && (!isLocked || canSignWhileLocked)}
              blockedReason={
                isLocked
                  ? 'Báo cáo đã có bản ký và đang khoá. Liên hệ admin để mở khoá nếu cần thay bản đã ký.'
                  : doc.isExisting
                  ? 'Có thay đổi chưa lưu. Hãy lưu báo cáo trước khi upload bản đã ký để file ký khớp với nội dung.'
                  : 'Lưu báo cáo trước, sau đó mới upload được bản đã ký.'
              }
              signed={signedDoc}
              isLoading={isSignedLoading}
              isUploading={isUploadingSigned}
              disabled={isSaving}
              onUpload={handleUploadSigned}
              onRemove={isLocked ? undefined : () => setRemoveSignedOpen(true)}
              trail={signedTrail}
              isRemoving={isRemovingSigned}
            />
          )}

          {isWeekly && doc.tasks && doc.tasks.length > 0 && (
            <ReportTaskTable
              tasks={doc.tasks}
              onTaskChange={handleTaskChange}
              disabled={editDisabled}
            />
          )}
        </div>
        </ReportReadOnlyContext.Provider>
      ) : (
        <EmptyState message="Không thể nạp dữ liệu mẫu báo cáo" />
      )}

      {lockConfirm && (
        <ConfirmDialog
          danger={lockConfirm === 'lock'}
          title={lockConfirm === 'lock' ? 'Khoá báo cáo?' : 'Mở khoá báo cáo?'}
          message={
            lockConfirm === 'lock'
              ? `Báo cáo ${role.unit} ngày ${dateStr} sẽ chuyển sang chỉ xem: người dùng không sửa/lưu/upload được cho tới khi admin mở khoá.${
                  hasUnsavedChanges ? ' Các thay đổi chưa lưu trên màn hình này sẽ không lưu được.' : ''
                }`
              : `Người dùng sẽ chỉnh sửa lại được báo cáo ${role.unit} ngày ${dateStr}. Báo cáo tự khoá lại khi có bản đã ký mới được upload.`
          }
          confirmLabel={lockConfirm === 'lock' ? 'Khoá' : 'Mở khoá'}
          onConfirm={() => {
            handleToggleLock().catch((err) => console.error(err));
          }}
          onCancel={() => setLockConfirm(null)}
        />
      )}

      {lockedNotice && (
        <ConfirmDialog
          hideCancel
          danger
          title="Báo cáo đã bị khoá"
          message={`Báo cáo ${role.unit} ngày ${dateStr} đã bị khoá nên không lưu được thay đổi của bạn. Vui lòng liên hệ admin để được mở khoá.`}
          confirmLabel="Trở về Home"
          onConfirm={() => {
            setLockedNotice(false);
            onGoHome();
          }}
          onCancel={() => undefined}
        />
      )}

      {photoReplaceAsk && (
        <ConfirmDialog
          danger
          title="Hình đã tồn tại"
          message={`Hình "${photoReplaceAsk.name}" đã có trong thư mục photos/layout của báo cáo này. Thay thế sẽ ghi đè hình cũ (hàng Layout nào đang dùng hình này cũng thấy hình mới). Huỷ thì giữ hình cũ và không tải hình này lên.`}
          confirmLabel="Thay thế"
          onConfirm={() => {
            photoReplaceAsk.resolve(true);
            setPhotoReplaceAsk(null);
          }}
          onCancel={() => {
            photoReplaceAsk.resolve(false);
            setPhotoReplaceAsk(null);
          }}
        />
      )}

      {pdfReplaceAsk && (
        <ConfirmDialog
          danger
          title="File PDF đã tồn tại"
          message={`File "${pdfReplaceAsk.name}" đã có trong thư mục pdf/eventlog của báo cáo này. Thay thế sẽ ghi đè file cũ (dòng sự kiện nào đang dùng file này cũng thấy nội dung mới). Huỷ thì giữ file cũ và không đính kèm file này.`}
          confirmLabel="Thay thế"
          onConfirm={() => {
            pdfReplaceAsk.resolve(true);
            setPdfReplaceAsk(null);
          }}
          onCancel={() => {
            pdfReplaceAsk.resolve(false);
            setPdfReplaceAsk(null);
          }}
        />
      )}

      {removeSignedOpen && doc && (
        <ConfirmDialog
          danger
          title="Xoá file đã ký?"
          message={`File PDF đã ký của ngày ${dateStr} sẽ bị xoá và ngày này quay về "Chưa upload bản ký". Hệ thống ghi lại người xoá và thời điểm xoá.`}
          confirmLabel="Xoá file đã ký"
          onConfirm={() => {
            handleRemoveSigned().catch((err) => console.error(err));
          }}
          onCancel={() => setRemoveSignedOpen(false)}
        />
      )}

      {/* 5. Print / Export PDF A4 Preview Modal */}
      {showPrintPreview && doc && (
        <ReportPrintPreviewModal
          doc={doc}
          photosFolder={photosFolder}
          resourceFolder={dataService.getResourceFolderPath()}
          onClose={() => setShowPrintPreview(false)}
        />
      )}
    </div>
    </SelectOpenContext.Provider>
  );
};
