import type { PortalModule } from '../core/types';
import { reportModule } from '../modules/report';
import { styleGuideModule } from '../modules/styleguide';

/**
 * Registry of all portal modules, in display order.
 * To add a module: create `src/modules/<name>/`, export a PortalModule, and add it here.
 */
export const PORTAL_MODULES: PortalModule[] = [
  reportModule,
  styleGuideModule, // UI Kit (chỉ admin) – tài liệu sống của các thành phần dùng chung
];
