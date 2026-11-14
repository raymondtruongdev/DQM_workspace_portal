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
 * Cột đọc theo TÊN HIỂN THỊ (Email, DisplayName, Department, DepartmentName, Role, Active), đã được `UserProfileService` đổi sang
 * các khoá này từ tên nội bộ thật. KHÔNG đọc theo tên nội bộ `field_N`: SharePoint tự đặt `field_N` theo thứ tự cột lúc
 * import nên không liên quan tới ý nghĩa cột (từng lấy nhầm cột Email làm tên hiển thị). Chỉ còn `Title` làm dự phòng cho email
 * của list cũ để cột Email nằm ở cột Title.
 */
/** Cột Active: Yes/No (true/false) hoặc chữ yes/y/1/active…; để trống hoặc không có cột = đang hoạt động. */
export function parseProfileActive(raw: unknown): boolean {
  if (raw === undefined || raw === null || String(raw).trim() === '') return true;
  if (typeof raw === 'boolean') return raw;
  if (typeof raw === 'number') return raw !== 0;
  return ['true', '1', 'yes', 'y', 'active', 'x', 'hoạt động'].indexOf(String(raw).trim().toLowerCase()) !== -1;
}

export function mapSharePointUserProfile(raw: Record<string, unknown>): UserProfileRecord {
  const text = (v: unknown): string => (v === undefined || v === null ? '' : String(v).trim());

  // Email: cột Email (dự phòng cho list cũ: cột Title)
  const rawEmail = text(raw.Email ?? raw.email ?? raw.Mail ?? raw.UserEmail ?? raw.UserPrincipalName ?? raw.Title);
  const cleanEmail = extractCleanEmail(rawEmail);
  const emailPrefix = cleanEmail.split('@')[0];

  // Username: cột Username, không có thì phần trước @ của email
  const rawUsername = text(raw.Username ?? raw.username);

  // Tên hiển thị: cột DisplayName, không có thì Username, rồi phần trước @ của email
  const rawDisplayName = text(raw.DisplayName ?? raw.Name);

  // Phòng ban và vai trò
  const deptCode = text(raw.DepartmentCode ?? raw.Department);
  const deptName = text(raw.DepartmentName);
  const role = text(raw.Role ?? raw.role).toLowerCase();

  return {
    email: cleanEmail,
    username: rawUsername || emailPrefix,
    displayName: rawDisplayName || rawUsername || emailPrefix,
    departmentCode: deptCode,
    departmentName: deptName || deptCode,
    role: role || 'user',
    active: parseProfileActive(raw.Active ?? raw.active),
  };
}
