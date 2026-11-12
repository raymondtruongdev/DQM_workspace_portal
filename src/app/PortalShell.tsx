import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import '../core/styles';
import type { IPortalProps } from './IPortalProps';
import type { UserSession } from '../core/auth';
import type { PortalConfig } from '../core/types';
import { DEFAULT_PORTAL_CONFIG } from '../core/types';
import { UserProfileService, extractCleanEmail, checkIsAdmin } from '../core/auth';
import { PortalConfigService, SharePointClient } from '../core/services';
import { SESSION_STORAGE_KEY } from '../core/constants';
import { HomeView } from '../modules/home';
import { PORTAL_MODULES } from './moduleRegistry';
import { EmptyState, Icon } from '../core/components';

export const PortalShell: React.FC<IPortalProps> = (props) => {
  const { spHttpClient, siteUrl, userDisplayName, userEmail, userLoginName } = props;

  const profileService = useMemo(() => new UserProfileService(new SharePointClient(spHttpClient)), [spHttpClient]);

  // Cấu hình portal (chế độ bảo trì): đọc lại mỗi khi quay về trang chủ / đổi module
  const configService = useMemo(() => new PortalConfigService(new SharePointClient(spHttpClient), siteUrl), [spHttpClient, siteUrl]);
  const [portalConfig, setPortalConfig] = useState<PortalConfig>(DEFAULT_PORTAL_CONFIG);
  const [isConfigLoaded, setIsConfigLoaded] = useState<boolean>(false);

  // Active module: null = Home (module launcher)
  const [activeModuleId, setActiveModuleId] = useState<string | null>(null);

  // Session state
  const [currentUser, setCurrentUser] = useState<UserSession | null>(() => {
    try {
      const saved = sessionStorage.getItem(SESSION_STORAGE_KEY);
      if (saved) return JSON.parse(saved);
    } catch {
      // Ignore
    }
    return null;
  });

  // State to track auto-SSO resolution on startup
  const [isResolvingAuth, setIsResolvingAuth] = useState<boolean>(() => {
    try {
      const saved = sessionStorage.getItem(SESSION_STORAGE_KEY);
      if (saved) return false;
    } catch {
      // Ignore
    }
    return true;
  });

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

  // Attempt automatic SSO resolution on component mount
  useEffect(() => {
    if (currentUser) {
      setIsResolvingAuth(false);
      return;
    }

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
          setCurrentUser(session);
          try {
            sessionStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(session));
          } catch {
            // Ignore
          }
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
  }, [userEmail, userLoginName, userDisplayName, siteUrl, profileService, currentUser]);

  // Profile selection (SSO dev/qa switcher)
  const handleSelectProfile = (user: UserSession): void => {
    setCurrentUser(user);
    try {
      sessionStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(user));
    } catch {
      // Ignore
    }
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
  if (portalConfig.maintenance && !currentUser?.isAdmin) {
    return (
      <div className="app-layout">
        <main className="main-content landing-wrapper">
          <EmptyState
            narrow
            icon="lock"
            title="Hệ thống đang bảo trì"
            message={portalConfig.message || 'Vui lòng quay lại sau.'}
          />
        </main>
      </div>
    );
  }

  const activeModule = PORTAL_MODULES.find((m) => m.id === activeModuleId && m.component);
  if (activeModule && activeModule.component) {
    const ModuleComponent = activeModule.component;
    // Khung chuẩn cho MỌI module: .app-layout (token + base) + .main-content (lề, max-width)
    return (
      <div className="app-layout">
        <main className="main-content">
          <ModuleComponent
            currentUser={currentUser ?? undefined}
            spHttpClient={spHttpClient}
            siteUrl={siteUrl}
            onGoHome={() => setActiveModuleId(null)}
          />
        </main>
      </div>
    );
  }

  return (
    <div className="app-layout">
      <HomeView
        currentUser={currentUser ?? undefined}
        modules={PORTAL_MODULES}
        maintenance={portalConfig.maintenance}
        onOpenModule={setActiveModuleId}
        onSelectProfile={handleSelectProfile}
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
