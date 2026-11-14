export interface UserSession {
  username: string;
  displayName: string;
  email?: string;
  department?: string;
  departmentCode?: string;
  role?: string;
  isAdmin?: boolean;
}

export interface UserProfileRecord {
  email: string;
  username: string;
  displayName?: string;
  departmentCode?: string;
  departmentName?: string;
  role: string;
}
