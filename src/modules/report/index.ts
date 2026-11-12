import * as React from 'react';
import type { PortalModule } from '../../core/types';
import { ReportModule } from './ReportModule';

export const reportModule: PortalModule = {
  id: 'report',
  title: 'Quản lý Report',
  icon: React.createElement(
    React.Fragment,
    null,
    React.createElement('path', {
      d: 'M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z',
      stroke: '#ffffff',
      strokeWidth: '2',
      fill: 'none',
      strokeLinejoin: 'round'
    }),
    React.createElement('polyline', {
      points: '14 2 14 8 20 8',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '2',
      strokeLinejoin: 'round'
    }),
    React.createElement('line', {
      x1: '16',
      y1: '13',
      x2: '8',
      y2: '13',
      stroke: '#ffffff',
      strokeWidth: '2'
    }),
    React.createElement('line', {
      x1: '16',
      y1: '17',
      x2: '8',
      y2: '17',
      stroke: '#ffffff',
      strokeWidth: '2'
    }),
    React.createElement('line', {
      x1: '10',
      y1: '9',
      x2: '8',
      y2: '9',
      stroke: '#ffffff',
      strokeWidth: '2'
    })
  ),
  component: ReportModule
};

export * from './ReportModule';

