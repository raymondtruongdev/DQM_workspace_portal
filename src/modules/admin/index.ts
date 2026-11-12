import * as React from 'react';
import type { PortalModule } from '../../core/types';
import { AdminModule } from './AdminModule';

export const adminModule: PortalModule = {
  id: 'admin',
  title: 'Admin Config',
  adminOnly: true,
  icon: React.createElement(
    React.Fragment,
    null,
    React.createElement('circle', { cx: '16', cy: '16', r: '4.5', stroke: '#ffffff', strokeWidth: '2', fill: 'none' }),
    React.createElement('path', {
      d: 'M16 4v4M16 24v4M4 16h4M24 16h4M7.5 7.5l2.8 2.8M21.7 21.7l2.8 2.8M7.5 24.5l2.8-2.8M21.7 10.3l2.8-2.8',
      stroke: '#ffffff',
      strokeWidth: '2',
      strokeLinecap: 'round'
    })
  ),
  component: AdminModule
};

export * from './AdminModule';
