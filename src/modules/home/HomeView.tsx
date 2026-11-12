import React, { useState, useEffect } from 'react';
import type { UserSession, UserProfileRecord } from '../../core/auth';
import { UserProfileService, extractCleanEmail, checkIsAdmin } from '../../core/auth';
import type { PortalModule } from '../../core/types';
import { Icon } from '../../core/components/Icon';

export interface HomeViewProps {
  currentUser?: UserSession;
  modules: PortalModule[];
  onOpenModule: (moduleId: string) => void;
  onSelectProfile?: (profile: UserSession) => void;
  profileService: UserProfileService;
  siteUrl: string;
  userDisplayName?: string;
  userEmail?: string;
  userLoginName?: string;
}

export const HomeView: React.FC<HomeViewProps> = ({
  currentUser,
  modules,
  onOpenModule,
  onSelectProfile,
  profileService,
  siteUrl,
  userDisplayName = 'User',
  userEmail = '',
  userLoginName = '',
}) => {
  const [profiles, setProfiles] = useState<UserProfileRecord[]>([]);
  const [showRoleSwitcher, setShowRoleSwitcher] = useState(false);

  const cleanCurrentEmail = extractCleanEmail(userEmail || userLoginName);
  const activeDisplayName = currentUser?.displayName || userDisplayName;
  const userInitials = activeDisplayName.slice(0, 2).toUpperCase();

  // Load UserProfiles list to provide role switching / diagnostics
  useEffect(() => {
    let isMounted = true;
    const loadProfiles = async (): Promise<void> => {
      try {
        const list = await profileService.getUserProfiles(siteUrl);
        if (isMounted) setProfiles(list);
      } catch (err: unknown) {
        console.warn('[HomeView] Could not load UserProfiles:', err);
      }
    };

    loadProfiles().catch((e) => console.error(e));
    return () => {
      isMounted = false;
    };
  }, [profileService, siteUrl]);

  const handleSwitch = (p: UserProfileRecord): void => {
    if (!onSelectProfile) return;
    const session: UserSession = {
      email: p.email || cleanCurrentEmail,
      username: p.username || p.email,
      displayName: p.displayName || p.username || userDisplayName,
      department: p.departmentName || p.departmentCode,
      departmentCode: p.departmentCode,
      role: p.role,
      isAdmin: checkIsAdmin(p.role),
    };
    onSelectProfile(session);
  };

  const isAuthorized = Boolean(currentUser);
  const visibleModules = modules.filter((m) => !m.adminOnly || currentUser?.isAdmin);

  return (
    <div className="landing-wrapper">
      <div className="landing-container">
        {/* Main Hero Card */}
        <section className="landing-hero-card">
          {/* Brand Badge */}
          <div className="landing-brand-badge">
            <Icon name="apps" size={16} />
            <span>Workspace Portal</span>
          </div>

          {/* User Identity & Authorization Profile Box */}
          <div className="landing-profile-box">
            {/* Centered Avatar */}
            <div className="landing-avatar-wrapper">
              <div className="landing-user-avatar">{userInitials}</div>
              {isAuthorized && <span className="landing-avatar-status" title="Đã xác thực Microsoft 365" />}
            </div>

            {/* Display Name */}
            <h2 className="landing-user-name">{activeDisplayName}</h2>

            {/* Email */}
            <div className="landing-user-email">
              <Icon name="mail" size={16} />
              <span>{currentUser?.email || cleanCurrentEmail || 'Microsoft 365 Account'}</span>
            </div>

            {/* Badges Row: Role Badge + Department Badge */}
            <div className="landing-badges-row">
              {isAuthorized ? (
                <span className={currentUser?.isAdmin ? 'sso-badge-admin' : 'sso-badge-user'}>
                  {currentUser?.isAdmin ? 'ADMIN' : (currentUser?.role?.toUpperCase() || 'USER')}
                </span>
              ) : (
                <span className="badge badge--danger">
                  Chưa gán quyền
                </span>
              )}

              {currentUser?.department && (
                <span className="landing-dept-badge">
                  <Icon name="building" size={16} />
                  <span>{currentUser.department}</span>
                </span>
              )}
            </div>
          </div>

          {/* Divider */}
          <div className="landing-divider" />

          {/* Primary Management Modules Grid (SharePoint Tile Style) */}
          {isAuthorized ? (
            <div className="landing-modules-container">
              <div className="sp-tiles-grid">
                {visibleModules.map((module) => (
                  <button
                    key={module.id}
                    type="button"
                    className="landing-grid-btn"
                    onClick={() => onOpenModule(module.id)}
                    title={`Nhấn để truy cập ${module.title}`}
                  >
                    <div className="landing-grid-btn-icon">
                      {/* style-lint: allow-svg – icon thương hiệu do module cung cấp (khung 32x32) */}
                      <svg width="26" height="26" viewBox="0 0 32 32" fill="none">
                        {module.icon}
                      </svg>
                    </div>
                    <span className="landing-grid-btn-text">{module.title}</span>
                    <Icon name="chevron-right" size={20} className="landing-grid-btn-arrow" />
                  </button>
                ))}
              </div>
            </div>
          ) : (
            <div className="landing-unauthorized">
              <div className="landing-alert-notice" role="status">
                <Icon name="info" size={20} />
                <span>
                  Tài khoản của bạn chưa có trong danh sách phân quyền <strong>UserProfiles</strong>. Vui lòng liên hệ Admin để được cấp quyền.
                </span>
              </div>
            </div>
          )}
        </section>

        {/* Developer & QA Role Switcher */}
        {profiles.length > 0 && (
          <div className="landing-switcher">
            <button
              type="button"
              className="btn btn--secondary btn--sm"
              onClick={() => setShowRoleSwitcher(!showRoleSwitcher)}
            >
              <Icon name={showRoleSwitcher ? 'chevron-up' : 'chevron-down'} size={16} />
              <span>{showRoleSwitcher ? 'Ẩn bộ chuyển đổi vai trò thử nghiệm' : 'Chế độ Thử nghiệm Vai trò (Developer / QA)'}</span>
            </button>

            {showRoleSwitcher && (
              <div className="sso-profiles-list">
                <p className="landing-switcher-hint">
                  Bấm vào hồ sơ bên dưới để thử nghiệm giao diện với vai trò tương ứng:
                </p>
                {profiles.map((p, idx) => (
                  <button
                    key={`${p.email}-${idx}`}
                    type="button"
                    className="sso-profile-item"
                    onClick={() => handleSwitch(p)}
                  >
                    <div className="sso-profile-item-info">
                      <span className="sso-profile-item-name">
                        {p.displayName || p.username} • {p.departmentName || p.departmentCode}
                      </span>
                      <span className="sso-profile-item-email">{p.email || p.username}</span>
                    </div>
                    <span className={checkIsAdmin(p.role) ? 'sso-badge-admin' : 'sso-badge-user'}>
                      {p.role}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default HomeView;
