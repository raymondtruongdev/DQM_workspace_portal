import * as React from 'react';
import { ModuleProps, RoleEntry } from '../../core/types';
import { SharePointClient } from '../../core/services';
import { REPORT_STORAGE_KEYS } from './constants/reportStorageKeys';
import {
  ReportDataService,
  ReportFileService,
  ReportRegistryService,
  ReportRoleService,
  ReportTemplateService
} from './services';
import { ReportLandingView } from './views/ReportLandingView';
import { ReportEntryView } from './views/ReportEntryView';
import { ReportCalendarView } from './views/ReportCalendarView';
import './report.css';

/** Báo cáo ngày (có lịch 14 ngày); báo cáo tuần hiện là placeholder nên vào thẳng form. */
const isDailyReport = (role: RoleEntry): boolean => !role.docId.toLowerCase().includes('week');

export const ReportModule: React.FC<ModuleProps> = (props) => {
  const {
    currentUser,
    spHttpClient,
    siteUrl,
    onGoHome,
    checkMaintenance
  } = props;

  const isAdmin = !!currentUser?.isAdmin;
  const effectiveEmail = currentUser?.email || currentUser?.username || '';

  // Initialize services
  const sharePointClient = React.useMemo(() => new SharePointClient(spHttpClient), [spHttpClient]);
  const fileService = React.useMemo(() => new ReportFileService(sharePointClient, siteUrl), [sharePointClient, siteUrl]);
  const roleService = React.useMemo(() => new ReportRoleService(fileService), [fileService]);
  const templateService = React.useMemo(() => new ReportTemplateService(fileService), [fileService]);
  const registryService = React.useMemo(
    () => new ReportRegistryService(sharePointClient, siteUrl, () => fileService.getLibraryPath()),
    [sharePointClient, siteUrl, fileService]
  );
  const dataService = React.useMemo(
    () => new ReportDataService(fileService, templateService, registryService),
    [fileService, templateService, registryService]
  );

  // View state
  const [currentView, setCurrentView] = React.useState<'landing' | 'calendar' | 'entry'>('landing');
  const [entryDate, setEntryDate] = React.useState<string | undefined>(undefined);
  // Tab đang xem ở trang danh sách (admin: Tổng quan / Báo cáo của tôi, mặc định Tổng quan); giữ khi quay lại từ lịch/form
  const [landingView, setLandingView] = React.useState<'mine' | 'overview'>(isAdmin ? 'overview' : 'mine');
  // Form được mở thẳng từ tab Tổng quan: Back quay về Tổng quan chứ không qua lịch
  const [entryFromOverview, setEntryFromOverview] = React.useState<boolean>(false);
  const [selectedRole, setSelectedRole] = React.useState<RoleEntry | null>(null);

  // Unit filter state
  const [selectedUnit, setSelectedUnit] = React.useState<string>(() => {
    try {
      return localStorage.getItem(REPORT_STORAGE_KEYS.LAST_UNIT) || 'ALL';
    } catch {
      return 'ALL';
    }
  });

  const [isLoadingRoles, setIsLoadingRoles] = React.useState<boolean>(true);
  const [assignedRoles, setAssignedRoles] = React.useState<RoleEntry[]>([]);
  const [allRoles, setAllRoles] = React.useState<RoleEntry[]>([]);

  // Load and resolve roles
  const loadRoles = React.useCallback(async (): Promise<void> => {
    setIsLoadingRoles(true);
    try {
      const resolution = await roleService.resolveUserRoles(effectiveEmail, isAdmin);
      setAssignedRoles(resolution.assignedRoles);
      setAllRoles(resolution.allRoles);
    } catch (err: unknown) {
      console.error('Error loading report roles:', err);
    } finally {
      setIsLoadingRoles(false);
    }
  }, [roleService, effectiveEmail, isAdmin]);

  React.useEffect(() => {
    loadRoles().catch(() => {});
  }, [loadRoles]);

  const handleSelectUnit = (unit: string): void => {
    setSelectedUnit(unit);
    try {
      localStorage.setItem(REPORT_STORAGE_KEYS.LAST_UNIT, unit);
    } catch {
      // Ignore storage errors
    }
  };

  const handleSelectReport = (role: RoleEntry): void => {
    setSelectedRole(role);
    setEntryDate(undefined);
    setEntryFromOverview(false);
    // Báo cáo ngày: xem lịch 14 ngày trước rồi mới vào form; báo cáo tuần (placeholder) vào thẳng form
    setCurrentView(isDailyReport(role) ? 'calendar' : 'entry');
    try {
      localStorage.setItem(REPORT_STORAGE_KEYS.LAST_DOC_ID, role.docId);
      localStorage.setItem(REPORT_STORAGE_KEYS.LAST_UNIT, role.unit);
    } catch {
      // Ignore storage errors
    }
  };

  const handleBackToLanding = (): void => {
    setCurrentView('landing');
    setSelectedRole(null);
    setEntryDate(undefined);
  };

  // Từ tab Tổng quan: mở thẳng form báo cáo của một đơn vị ở một ngày
  const handleOpenReportDate = (role: RoleEntry, dateStr: string): void => {
    setSelectedRole(role);
    setEntryDate(dateStr);
    setEntryFromOverview(true);
    setCurrentView('entry');
  };

  const handleOpenDate = (dateStr: string): void => {
    setEntryDate(dateStr);
    setEntryFromOverview(false);
    setCurrentView('entry');
  };

  // Từ form nhập quay lại: báo cáo ngày về lịch (đọc lại trạng thái), báo cáo tuần về danh sách
  const handleBackFromEntry = (): void => {
    if (selectedRole && isDailyReport(selectedRole) && !entryFromOverview) {
      setCurrentView('calendar');
    } else {
      handleBackToLanding();
    }
  };

  return (
    <div className="report-module-root">
      {currentView === 'calendar' && selectedRole ? (
        <ReportCalendarView
          role={selectedRole}
          dataService={dataService}
          onCheckMaintenance={checkMaintenance}
          onBack={handleBackToLanding}
          onOpenDate={handleOpenDate}
        />
      ) : currentView === 'entry' && selectedRole ? (
        <ReportEntryView
          role={selectedRole}
          currentUserEmail={effectiveEmail}
          dataService={dataService}
          onCheckMaintenance={checkMaintenance}
          initialDate={entryDate}
          onBack={handleBackFromEntry}
        />
      ) : (
        <ReportLandingView
          assignments={assignedRoles}
          allRoles={allRoles}
          userEmail={effectiveEmail}
          selectedUnit={selectedUnit}
          onSelectUnit={handleSelectUnit}
          onSelectReport={handleSelectReport}
          onRefresh={loadRoles}
          onGoHome={onGoHome}
          isLoading={isLoadingRoles}
          isAdmin={isAdmin}
          dataService={dataService}
          view={landingView}
          onChangeView={setLandingView}
          onOpenReportDate={handleOpenReportDate}
        />
      )}
    </div>
  );
};

