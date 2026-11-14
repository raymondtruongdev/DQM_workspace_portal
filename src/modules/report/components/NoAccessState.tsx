import * as React from 'react';
import { RoleEntry } from '../../../core/types';
import { Icon, EmptyState } from '../../../core/components';

export interface INoAccessStateProps {
  userEmail: string;
  allRoles?: RoleEntry[];
  onRefresh?: () => void;
}

export const NoAccessState: React.FC<INoAccessStateProps> = ({ userEmail, onRefresh }) => (
  <EmptyState
    narrow
    tone="danger"
    icon="error-circle"
    title="Chưa được phân quyền quản lý báo cáo"
    message={
      <>
        Tài khoản <strong>{userEmail || 'của bạn'}</strong> chưa được phân công báo cáo (DailyReport / WeeklyReport) trong list <code>ReportRole</code>, hoặc dòng phân công đang để Active = No.
      </>
    }
    hint="Vui lòng liên hệ Quản trị viên để phân quyền trong SharePoint"
    action={
      onRefresh && (
        <button type="button" onClick={onRefresh} className="btn btn--primary">
          <Icon name="refresh" />
          <span>Kiểm tra lại quyền</span>
        </button>
      )
    }
  />
);
