export const REPORT_CONFIG = {
  LIBRARY_NAME: 'CapNhatCongViec',
  ROLE_FILE_NAME: 'Report_Role.xlsx',
  /** Danh mục mã Activity (P6) và cấu kiện, mỗi đơn vị một sheet; nằm cùng cấp Report_Role.xlsx. */
  ACTIVITY_FILE_NAME: 'ActivityP6.xlsx',
  TEMPLATE_FOLDER: '_template',
  VALID_DOC_IDS: ['DailyReport', 'WeeklyReport'] as const
} as const;

