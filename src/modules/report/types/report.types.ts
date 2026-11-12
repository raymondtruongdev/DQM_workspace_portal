export interface TemplateVersionInfo {
  fileName: string;          // e.g. "template_ST01_20261109.xlsx"
  version: string;           // e.g. "template_ST01_20261109"
  versionDate: string;       // e.g. "20261109"
  serverRelativeUrl: string;
  timeLastModified?: string;
}

/** Metadata fields a template may flag as required (column `Required` of sheet 01_General). */
export type RequirableField = 'location' | 'reportNo' | 'shift';

/** Thông tin chung của báo cáo (sheet 01_General). Nội dung báo cáo nằm trong các mục (sections). */
export interface ReportMetadata {
  /** Fields that must be filled before saving/printing (from template flag). */
  requiredFields?: RequirableField[];
  docTitle: string;             // e.g. "BÁO CÁO NGÀY" or "BÁO CÁO TUẦN"
  docTitleEn: string;           // e.g. "DAILY REPORT"
  subTitle: string;             // e.g. "Tuyến Metro số 2 – Gói thầu M2.1..."
  docCode: string;              // e.g. "HCM2_1-23_CS-MAUR_TCJV-CPZZ-ST06-CWS-RP-041-S3-C01_BCN"
  templateVersion?: string;     // e.g. "template_ST01_20261109"
  location: string;             // e.g. "ST06 – Ga ST06"
  reportNo: string;             // e.g. "041"
  shift: string;                // e.g. "Cả ngày (2 ca)"
  email: string;                // User email
  updateTime: string;           // ISO 8601 or formatted string
}

export type CellValue = string | number | undefined;
export type ColumnType = 'text' | 'number' | 'select' | 'calc';
export type ColumnTone = 'danger' | 'warning';

/** Cột của bảng động, khai báo trong sheet (dòng mã cột + nhãn + spec). */
export interface TableColumn {
  code: string;                       // Mã cột, ví dụ "Nos"
  label: string;                      // "Giờ công | Man-hrs"
  spec: string;                       // Nguyên văn spec trong sheet, ví dụ "number;total;w=14"
  type: ColumnType;
  options?: string[];                 // select:A/B/C
  formula?: string;                   // calc:Nos*Hrs
  key?: boolean;                      // dòng IsLocked => ô này chỉ đọc
  total?: boolean;                    // có dòng Tổng
  align?: 'left' | 'center' | 'right';
  width?: number;                     // % bề rộng (tuỳ chọn)
  defaultValue?: string;              // giá trị khi thêm dòng; {unit}, {n3}
  tones?: Record<string, ColumnTone>; // tô màu theo giá trị (select)
}

export interface TableRow {
  id: string;
  isLocked?: boolean;
  values: Record<string, CellValue>;
}

/** Một mục dạng bảng, mỗi mục = một sheet 0N_<Tên> trong template/báo cáo. */
export interface ReportTable {
  kind: 'table';
  key: string;        // tên sheet (không tiền tố số), chữ thường: "manpower"
  sheetBase: string;  // tên sheet gốc không tiền tố: "Manpower"
  title: string;      // "NHÂN LỰC"
  titleEn: string;    // "MANPOWER"
  shortTitle: string; // nhãn thanh nhảy nhanh
  summary: string;    // Tóm tắt | Summary
  columns: TableColumn[];
  rows: TableRow[];
}

export type FieldType = 'text' | 'textarea' | 'number' | 'select';

/** Trường của mục dạng form (một dòng trong sheet form). */
export interface FormField {
  key: string;
  label: string;                // "Sáng | AM"
  spec: string;                 // Nguyên văn spec trong sheet
  type: FieldType;
  options?: string[];           // select:A/B/C
  full?: boolean;               // trải hết hàng
  span?: number;                // số cột lưới chiếm (mặc định 1)
  rows?: number;                // số dòng của textarea
  placeholder?: string;
  defaultValue?: string;        // giá trị khi trống; {user} {siteManager} {unit} được điền tự động
  required?: boolean;           // cột Required của sheet
  value: CellValue;
}

/** Một mục dạng form (thời tiết, công việc, HSE, ký duyệt…): mỗi sheet form = một mục. */
export interface ReportForm {
  kind: 'form';
  key: string;
  sheetBase: string;
  title: string;
  titleEn: string;
  shortTitle: string;
  columns: number;              // số cột lưới (và số cột bảng khi in)
  sign: boolean;                // in dạng khung chữ ký
  summary: string;              // Tóm tắt | Summary (dòng 2 của sheet)
  showSummary: boolean;         // có hiện ô Tóm tắt không
  fields: FormField[];
}

/** Nhóm ảnh của một bộ phận: tên bộ phận, tóm tắt và các file ảnh trong photos/<groupId>/. */
export interface PhotoGroup {
  id: string;           // tên thư mục con chứa ảnh
  department: string;   // "DCEM", "HSE"…
  summary: string;
  files: string[];      // tên file ảnh
}

/** Mục ảnh hiện trường theo bộ phận: mỗi sheet Photo = một mục. */
export interface ReportPhotoSection {
  kind: 'photo';
  key: string;
  sheetBase: string;
  title: string;
  titleEn: string;
  shortTitle: string;
  summary: string;
  showSummary: boolean;
  columns: number;      // số ảnh mỗi hàng
  groups: PhotoGroup[];
}

export type ReportSection = ReportTable | ReportForm | ReportPhotoSection;


export interface ReportTaskRow {
  id: string; // "1", "1.1", "1.2", "2", etc.
  taskName: string;
  completionRate?: number; // 0 - 100
  quantity?: number;
  unit: string;
  parentId?: string;
  level: number;
  isLeaf: boolean;
}

export interface ReportTaskNode extends ReportTaskRow {
  children: ReportTaskNode[];
}

export interface ReportDoc {
  docId: string;          // "DailyReport" | "WeeklyReport"
  unit: string;           // e.g. "ST01", "ST06"
  dateStr: string;        // "YYYY-MM-DD"
  metadata: ReportMetadata;
  /** Các mục của báo cáo (bảng hoặc form), theo thứ tự sheet trong template. */
  sections: ReportSection[];
  tasks: ReportTaskRow[]; // Used for WeeklyReport or legacy tree-table
  eTag?: string;
  isExisting?: boolean;
  copiedFromDate?: string; // e.g. "2026-11-10" when pre-filled from previous day
  latestTemplateInfo?: TemplateVersionInfo;
  templateUpdateRequired?: boolean;
  outdatedSourceType?: 'existing' | 'previous_day';
  outdatedTemplateVersion?: string;
}
