import type { UserProfileRecord } from './auth.types';

/**
 * Extracts a clean email from raw email or SharePoint UPN/loginName
 * (Strips claims prefixes like "i:0#.f|membership|" if present)
 */
export function extractCleanEmail(rawEmailOrLoginName?: string): string {
  if (!rawEmailOrLoginName) return '';
  const normalized = String(rawEmailOrLoginName)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ')
    .trim();
  const parts = normalized.split('|');
  const candidate = parts[parts.length - 1].trim();
  const emailMatch = candidate.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
  if (emailMatch) {
    return emailMatch[1].toLowerCase();
  }
  return candidate.replace(/^mailto:/i, '').trim().toLowerCase();
}

/**
 * Checks whether a given role string corresponds to an administrator
 */
export function checkIsAdmin(role?: string): boolean {
  if (!role) return false;
  const r = role.toLowerCase().trim();
  return r === 'admin' || r === 'administrator' || r === 'quantri' || r === 'quản trị';
}

/**
 * Maps raw SharePoint UserProfiles list items (or imported Excel table) to clean UserProfileRecord
 *
 * Cột đọc theo tên hiển thị (Email, DisplayName, Department, DepartmentName, Role); các tên field_N chỉ là dự phòng cho list
 * cũ từng có thêm cột số thứ tự ở đầu (nay đã bỏ):
 * - Col 1: Email / field_1 / Title / Mail / UserEmail / UserPrincipalName
 * - Col 2: Username / field_2 / Name / DisplayName / Title
 * - Col 3: DepartmentCode / Department / field_3
 * - Col 4: DepartmentName / field_4 / Department
 * - Col 5: Role / field_5
 */
/** Cột Active: Yes/No (true/false) hoặc chữ yes/y/1/active…; để trống hoặc không có cột = đang hoạt động. */
export function parseProfileActive(raw: unknown): boolean {
  if (raw === undefined || raw === null || String(raw).trim() === '') return true;
  if (typeof raw === 'boolean') return raw;
  if (typeof raw === 'number') return raw !== 0;
  return ['true', '1', 'yes', 'y', 'active', 'x', 'hoạt động'].indexOf(String(raw).trim().toLowerCase()) !== -1;
}

export function mapSharePointUserProfile(raw: Record<string, unknown>): UserProfileRecord {
  // Extract email candidate
  const rawEmail = String(
    raw.Email ?? raw.email ?? raw.Mail ?? raw.UserEmail ?? raw.UserPrincipalName ?? raw.field_1 ?? raw.Title ?? ''
  ).trim();
  const cleanEmail = extractCleanEmail(rawEmail);

  // Extract username / code
  const rawUsername = String(
    raw.Username ?? raw.username ?? raw.field_2 ?? raw.field_1 ?? raw.Title ?? ''
  ).trim();

  // Extract display name (fall back to username or email prefix)
  const rawDisplayName = String(
    raw.DisplayName ?? raw.Name ?? raw.field_2 ?? rawUsername ?? ''
  ).trim();

  // Department code & name
  const deptCode = String(raw.DepartmentCode ?? raw.Department ?? raw.field_3 ?? '').trim();
  const deptName = String(raw.DepartmentName ?? raw.field_4 ?? deptCode ?? '').trim();

  // Role
  const role = String(raw.Role ?? raw.role ?? raw.field_5 ?? 'user').trim().toLowerCase();

  return {
    email: cleanEmail,
    username: rawUsername || cleanEmail.split('@')[0],
    displayName: rawDisplayName || rawUsername || cleanEmail.split('@')[0],
    departmentCode: deptCode,
    departmentName: deptName || deptCode,
    role: role || 'user',
    active: parseProfileActive(raw.Active ?? raw.active),
  };
}
