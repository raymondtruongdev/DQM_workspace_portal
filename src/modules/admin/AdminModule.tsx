import * as React from 'react';
import type { ModuleProps } from '../../core/types';
import { Icon, PageHeader, TileGrid } from '../../core/components';
import { PortalConfigService, SharePointClient } from '../../core/services';
import { StyleGuideView } from './views/StyleGuideView';
import { ShutdownView } from './views/ShutdownView';
import { DailyReportSetupView } from './views/DailyReportSetupView';
import './admin.css';

type AdminView = 'menu' | 'styleguide' | 'shutdown' | 'dailyreport';

/** Admin Config: menu các công cụ quản trị (chỉ admin thấy). */
export const AdminModule: React.FC<ModuleProps> = ({ currentUser, spHttpClient, siteUrl, onGoHome }) => {
  const [view, setView] = React.useState<AdminView>('menu');
  const configService = React.useMemo(
    () => new PortalConfigService(new SharePointClient(spHttpClient), siteUrl),
    [spHttpClient, siteUrl]
  );

  if (view === 'styleguide') {
    return <StyleGuideView onBack={() => setView('menu')} />;
  }

  if (view === 'shutdown') {
    return (
      <ShutdownView
        configService={configService}
        adminName={currentUser?.email || currentUser?.username || 'admin'}
        onBack={() => setView('menu')}
      />
    );
  }

  if (view === 'dailyreport') {
    return (
      <DailyReportSetupView
        configService={configService}
        adminName={currentUser?.email || currentUser?.username || 'admin'}
        onBack={() => setView('menu')}
      />
    );
  }

  return (
    <div className="admin-module-root">
      <PageHeader
        leading={
          <button type="button" className="btn btn--secondary" onClick={onGoHome}>
            <Icon name="home" />
            <span>Trang Chủ</span>
          </button>
        }
        title="Admin Config"
      />
      <TileGrid
        items={[
          {
            id: 'styleguide',
            title: 'UI Kit (Style Guide)',
            icon: <Icon name="paint" size={26} />,
            onClick: () => setView('styleguide')
          },
          {
            id: 'dailyreport',
            title: 'Setup DailyReport',
            icon: <Icon name="calendar" size={26} />,
            onClick: () => setView('dailyreport')
          },
          {
            id: 'shutdown',
            title: 'Shutdown Server',
            icon: <Icon name="power" size={26} />,
            onClick: () => setView('shutdown')
          }
        ]}
      />
    </div>
  );
};
