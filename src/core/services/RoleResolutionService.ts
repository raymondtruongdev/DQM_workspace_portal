import { extractCleanEmail } from '../auth';
import { RoleEntry, UserRoleResolution } from '../types/role.types';

export class RoleResolutionService {
  /**
   * Filter and resolve role assignments for a given user email and admin flag.
   */
  public static resolveRoles(
    allRoles: RoleEntry[],
    rawUserEmail: string,
    isAdmin: boolean
  ): UserRoleResolution {
    const cleanEmail = extractCleanEmail(rawUserEmail).toLowerCase();
    const activeRoles = allRoles.filter((r) => r.status);

    let assignedRoles: RoleEntry[];
    if (isAdmin) {
      assignedRoles = activeRoles;
    } else {
      assignedRoles = activeRoles.filter((r) => {
        const entryEmail = extractCleanEmail(r.email).toLowerCase();
        return entryEmail === cleanEmail;
      });
    }

    return {
      userEmail: cleanEmail,
      isAdmin,
      assignedRoles,
      allRoles
    };
  }
}
