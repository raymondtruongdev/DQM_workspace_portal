import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import '../core/styles';
import type { IPortalProps } from './IPortalProps';
import type { UserSession } from '../core/auth';
import type { PortalConfig } from '../core/types';
import { DEFAULT_PORTAL_CONFIG } from '../core/types';
import { UserProfileService, extractCleanEmail, checkIsAdmin } from '../core/auth';
import { PortalConfigService, SharePointClient } from '../core/services';
import { HomeView } from '../modules/home';
import { PORTAL_MODULES } from './moduleRegistry';
import { EmptyState, Icon } from '../core/components';

/** Chu kỳ kiểm tra chế độ bảo trì khi người dùng đang mở portal. */
const MAINTENANCE_POLL_MS = 15000;

export const PortalShell: React.FC<IPortalProps> = (props) => {
  const { spHttpClient, siteUrl, userDisplayName, userEmail, userLoginName } = props;

  const profileService = useMemo(() => new UserProfileService(new SharePointClient(spHttpClient)), [spHttpClient]);

  // Cấu hình portal (chế độ bảo trì): đọc lại mỗi khi quay về trang chủ / đổi module
  const configService = useMemo(() => new PortalConfigService(new SharePointClient(spHttpClient), siteUrl), [spHttpClient, siteUrl]);
  const [portalConfig, setPortalConfig] = useState<PortalConfig>(DEFAULT_PORTAL_CONFIG);
  const [isConfigLoaded, setIsConfigLoaded] = useState<boolean>(false);

  // Active module: null = Home (module launcher)
  const [activeModuleId, setActiveModuleId] = useState<string | null>(null);

  // Danh tính THẬT của người đang mở web part: luôn đọc lại từ list UserProfiles mỗi lần mở (không lưu/không tin sessionStorage)
  const [realUser, setRealUser] = useState<UserSession | null>(null);
  // Chỉ admin thật mới được "xem với vai trò khác" (Developer / QA); chỉ giữ trong bộ nhớ, tải lại trang là mất
  const [viewAsUser, setViewAsUser] = useState<UserSession | null>(null);
  const currentUser = viewAsUser ?? realUser;
  const isRealAdmin = Boolean(realUser?.isAdmin);

  // State to track SSO resolution on startup
  const [isResolvingAuth, setIsResolvingAuth] = useState<boolean>(true);

  useEffect(() => {
    let alive = true;
    configService
      .getConfig()
      .then((c) => {
        if (!alive) return;
        setPortalConfig(c);
        setIsConfigLoaded(true);
      })
      .catch((e) => console.error(e));
    return () => {
      alive = false;
    };
  }, [configService, activeModuleId]);

  // Module gọi trước mỗi thao tác ghi: đọc lại cấu hình ngay lúc đó (true = bị chặn)
  const checkMaintenance = useCallback(async (): Promise<boolean> => {
    const c = await configService.getConfig();
    setPortalConfig(c);
    return c.maintenance && !isRealAdmin;
  }, [configService, isRealAdmin]);

  // Người dùng đang ở trong module khi admin bật bảo trì: kiểm tra định kỳ và khi quay lại tab
  useEffect(() => {
    let alive = true;
    const refresh = (): void => {
      configService
        .getConfig()
        .then((c) => {
          if (alive) setPortalConfig(c);
        })
        .catch((e) => console.error(e));
    };
    const onVisible = (): void => {
      if (document.visibilityState === 'visible') refresh();
    };
    const timer = window.setInterval(refresh, MAINTENANCE_POLL_MS);
    document.addEventListener('visibilitychange', onVisible);
    window.addEventListener('focus', refresh);
    return () => {
      alive = false;
      window.clearInterval(timer);
      document.removeEventListener('visibilitychange', onVisible);
      window.removeEventListener('focus', refresh);
    };
  }, [configService]);

  // Xác định danh tính từ list UserProfiles khi mở web part
  useEffect(() => {
    let isMounted = true;
    const autoResolveSSO = async (): Promise<void> => {
      const cleanEmail = extractCleanEmail(userEmail || userLoginName);
      if (!cleanEmail) {
        if (isMounted) setIsResolvingAuth(false);
        return;
      }

      try {
        console.log(`[PortalShell] Attempting auto-SSO resolution for: ${cleanEmail}`);
        const profiles = await profileService.getUserProfiles(siteUrl);
        const match = profiles.find((p) => {
          return (
            p.email.toLowerCase() === cleanEmail ||
            p.username.toLowerCase() === cleanEmail ||
            extractCleanEmail(p.email) === cleanEmail
          );
        });

        if (!match) {
          console.warn(
            `[PortalShell] Không khớp hồ sơ nào: đọc được ${profiles.length} dòng UserProfiles, email đăng nhập "${cleanEmail}", email trong list:`,
            profiles.map((p) => p.email)
          );
        }
        if (match && isMounted) {
          const session: UserSession = {
            email: match.email || cleanEmail,
            username: match.username || match.email,
            displayName: match.displayName || match.username || userDisplayName,
            department: match.departmentName || match.departmentCode,
            departmentCode: match.departmentCode,
            role: match.role,
            isAdmin: checkIsAdmin(match.role),
          };
          console.log('[PortalShell] Auto-SSO success:', session);
          setRealUser(session);
        }
      } catch (err: unknown) {
        console.warn('[PortalShell] Auto-SSO resolution check:', err);
      } finally {
        if (isMounted) {
          setIsResolvingAuth(false);
        }
      }
    };

    autoResolveSSO().catch((err) => {
      console.error('[PortalShell] autoResolveSSO unhandled error:', err);
      if (isMounted) setIsResolvingAuth(false);
    });

    return () => {
      isMounted = false;
    };
  }, [userEmail, userLoginName, userDisplayName, siteUrl, profileService]);

  // Xem với vai trò khác (Developer / QA): chỉ admin thật được gọi, kiểm tra lại ở đây
  const handleSelectProfile = (user: UserSession): void => {
    if (!isRealAdmin) return;
    setViewAsUser(user);
  };

  // Loading state while checking M365 SSO
  if (isResolvingAuth || !isConfigLoaded) {
    return (
      <div className="app-layout">
        <main className="main-content landing-wrapper">
          <div className="portal-loading">
            <Icon name="refresh" size={32} spin />
            <p>Đang xác thực danh tính Microsoft 365...</p>
          </div>
        </main>
      </div>
    );
  }

  // Chế độ bảo trì: chỉ admin còn truy cập được
  const isBlocked = portalConfig.maintenance && !isRealAdmin;
  const maintenanceScreen = (
    <main className="main-content landing-wrapper">
      <EmptyState
        narrow
        icon="lock"
        title="Hệ thống đang bảo trì"
        message={portalConfig.message || 'Vui lòng quay lại sau.'}
      />
    </main>
  );

  const activeModule = PORTAL_MODULES.find((m) => m.id === activeModuleId && m.component);
  if (activeModule && activeModule.component) {
    const ModuleComponent = activeModule.component;
    // Khung chuẩn cho MỌI module: .app-layout (token + base) + .main-content (lề, max-width).
    // Khi bị khoá bảo trì: ẩn module (không tương tác được: lưu, Back, Home…) nhưng vẫn giữ trạng thái
    // đang nhập để người dùng tiếp tục khi bảo trì kết thúc.
    return (
      <div className="app-layout">
        {isBlocked && maintenanceScreen}
        <main className="main-content" style={isBlocked ? { display: 'none' } : undefined} aria-hidden={isBlocked}>
          <ModuleComponent
            currentUser={currentUser ?? undefined}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            onGoHome={() => setActiveModuleId(null)}
            checkMaintenance={checkMaintenance}
          />
        </main>
      </div>
    );
  }

  if (isBlocked) {
    return <div className="app-layout">{maintenanceScreen}</div>;
  }

  return (
    <div className="app-layout">
      <HomeView
        currentUser={currentUser ?? undefined}
        modules={PORTAL_MODULES}
        maintenance={portalConfig.maintenance}
        onOpenModule={setActiveModuleId}
        onSelectProfile={isRealAdmin ? handleSelectProfile : undefined}
        profileService={profileService}
        siteUrl={siteUrl}
        userDisplayName={userDisplayName}
        userEmail={userEmail}
        userLoginName={userLoginName}
      />
    </div>
  );
};

export default PortalShell;
