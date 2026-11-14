export const REPORT_CONFIG = {
  LIBRARY_NAME: 'CapNhatCongViec',
  /** Danh mục Activity (P6): một sheet ActivityID | Unit | Description; nằm ở gốc thư viện (cùng chỗ với thư mục _template). */
  ACTIVITY_FILE_NAME: 'ActivityP6.xlsx',
  TEMPLATE_FOLDER: '_template',
  VALID_DOC_IDS: ['DailyReport', 'WeeklyReport'] as const
} as const;
