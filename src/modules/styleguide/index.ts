import * as React from 'react';
import type { PortalModule } from '../../core/types';
import { StyleGuideModule } from './StyleGuideModule';

export const styleGuideModule: PortalModule = {
  id: 'styleguide',
  title: 'UI Kit (Style Guide)',
  adminOnly: true,
  icon: React.createElement(
    React.Fragment,
    null,
    React.createElement('rect', { x: '5', y: '5', width: '9', height: '9', rx: '2', stroke: '#ffffff', strokeWidth: '2', fill: 'none' }),
    React.createElement('rect', { x: '18', y: '5', width: '9', height: '9', rx: '2', stroke: '#ffffff', strokeWidth: '2', fill: 'none' }),
    React.createElement('rect', { x: '5', y: '18', width: '9', height: '9', rx: '2', stroke: '#ffffff', strokeWidth: '2', fill: 'none' }),
    React.createElement('rect', { x: '18', y: '18', width: '9', height: '9', rx: '2', stroke: '#ffffff', strokeWidth: '2', fill: 'rgba(255, 255, 255, 0.22)' })
  ),
  component: StyleGuideModule
};

export * from './StyleGuideModule';
