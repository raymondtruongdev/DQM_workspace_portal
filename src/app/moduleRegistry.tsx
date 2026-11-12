import type { PortalModule } from '../core/types';
import { reportModule } from '../modules/report';
import { adminModule } from '../modules/admin';

/**
 * Registry of all portal modules, in display order.
 * To add a module: create `src/modules/<name>/`, export a PortalModule, and add it here.
 */
export const PORTAL_MODULES: PortalModule[] = [
  reportModule,
  adminModule, // Admin Config (chỉ admin): UI Kit, Shutdown Server
];
