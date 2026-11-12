import * as React from 'react';
import { RoleEntry } from '../../../core/types';
import { ReportTile } from '../components/ReportTile';
import { ReportOverviewView } from './ReportOverviewView';
import { ReportDataService } from '../services/ReportDataService';
import { NoAccessState } from '../components/NoAccessState';
import { Icon, PageHeader, FilterChips, SearchInput, LoadingState, EmptyState } from '../../../core/components';

export interface IReportLandingViewProps {
  assignments: RoleEntry[];
  allRoles?: RoleEntry[];
  userEmail: string;
  selectedUnit: string;
  onSelectUnit: (unit: string) => void;
  onSelectReport: (role: RoleEntry) => void;
  onRefresh: () => void;
  onGoHome: () => void;
  isLoading: boolean;
  /** Admin thấy thêm tab "Tổng quan" (theo dõi trạng thái mọi đơn vị). */
  isAdmin: boolean;
  dataService: ReportDataService;
  view: 'mine' | 'overview';
  onChangeView: (view: 'mine' | 'overview') => void;
  /** Mở form báo cáo của một đơn vị ở một ngày (từ tab Tổng quan). */
  onOpenReportDate: (role: RoleEntry, dateStr: string) => void;
}

export const ReportLandingView: React.FC<IReportLandingViewProps> = ({
  assignments,
  allRoles = [],
  userEmail,
  selectedUnit,
  onSelectUnit,
  onSelectReport,
  onRefresh,
  onGoHome,
  isLoading,
  isAdmin,
  dataService,
  view,
  onChangeView,
  onOpenReportDate
}) => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  // Extract list of unique units
  const units = React.useMemo(() => {
    const set = new Set<string>();
    assignments.forEach((a) => {
      if (a.unit) set.add(a.unit.trim());
    });
    return Array.from(set).sort();
  }, [assignments]);

  // Filter assignments by unit and search term
  const filteredAssignments = React.useMemo(() => {
    return assignments.filter((role) => {
      // Unit filter
      if (selectedUnit && selectedUnit !== 'ALL') {
        if (role.unit.trim().toLowerCase() !== selectedUnit.trim().toLowerCase()) {
          return false;
        }
      }

      // Search term
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase();
        const matchesUnit = role.unit.toLowerCase().includes(term);
        const matchesDocId = role.docId.toLowerCase().includes(term);
        const matchesTitle = (role.docTitle || '').toLowerCase().includes(term);
        const matchesManager = (role.siteManager || '').toLowerCase().includes(term);
        if (!matchesUnit && !matchesDocId && !matchesTitle && !matchesManager) {
          return false;
        }
      }

      return true;
    });
  }, [assignments, selectedUnit, searchTerm]);

  const header = (
    <PageHeader
      leading={
        <button
          type="button"
          className="btn btn--secondary"
          onClick={onGoHome}
          title="Quay về trang chủ Portal"
          aria-label="Về Trang Chủ"
        >
          <Icon name="home" />
          <span>Trang Chủ</span>
        </button>
      }
      title="Quản lý Báo cáo thi công"
      actions={
        <button
          type="button"
          className="btn btn--secondary"
          onClick={onRefresh}
          disabled={isLoading}
          title="Tải lại danh sách phân quyền"
          aria-label="Làm mới"
        >
          <Icon name="refresh" spin={isLoading} />
          <span>{isLoading ? 'Đang tải...' : 'Làm mới'}</span>
        </button>
      }
    />
  );

  if (!isLoading && assignments.length === 0) {
    return (
      <div className="report-landing-container">
        {header}
        <NoAccessState userEmail={userEmail} allRoles={allRoles} onRefresh={onRefresh} />
      </div>
    );
  }

  const unitOptions = [
    { value: 'ALL', label: `Tất cả (${assignments.length})` },
    ...units.map((u) => ({
      value: u,
      label: `${u} (${assignments.filter((a) => a.unit.trim().toLowerCase() === u.toLowerCase()).length})`
    }))
  ];
  const activeUnit = units.find((u) => u.toLowerCase() === selectedUnit.toLowerCase()) || 'ALL';

  const viewTabs = isAdmin ? (
    <FilterChips
      ariaLabel="Chế độ xem"
      value={view}
      onChange={(v) => onChangeView(v as 'mine' | 'overview')}
      options={[
        { value: 'overview', label: 'Tổng quan' },
        { value: 'mine', label: 'Báo cáo của tôi' }
      ]}
    />
  ) : null;

  if (isAdmin && view === 'overview') {
    return (
      <div className="report-landing-container">
        {header}
        {viewTabs}
        <ReportOverviewView
          assignments={assignments}
          dataService={dataService}
          onOpenUnit={onSelectReport}
          onOpenDate={onOpenReportDate}
        />
      </div>
    );
  }

  return (
    <div className="report-landing-container">
      {header}
      {viewTabs}

      <div className="report-landing-controls">
        <FilterChips
          options={unitOptions}
          value={activeUnit}
          onChange={onSelectUnit}
          ariaLabel="Lọc theo trạm"
        />
        <SearchInput
          value={searchTerm}
          onChange={setSearchTerm}
          placeholder="Tìm theo mã ga, loại báo cáo, người quản lý..."
        />
      </div>

      {isLoading ? (
        <LoadingState message="Đang tải dữ liệu phân quyền báo cáo..." />
      ) : filteredAssignments.length === 0 ? (
        <EmptyState message="Không tìm thấy báo cáo phù hợp với điều kiện lọc" />
      ) : (
        <div className="report-tiles-grid">
          {filteredAssignments.map((role) => (
            <ReportTile
              key={`${role.unit}_${role.docId}_${role.id}`}
              role={role}
              onClick={onSelectReport}
            />
          ))}
        </div>
      )}
    </div>
  );
};
