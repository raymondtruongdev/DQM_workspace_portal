export interface TemplateVersionInfo {
  fileName: string;          // e.g. "template_DRP_ST01_20261114.json"
  version: string;           // e.g. "template_DRP_ST01_20261114"
  versionDate: string;       // e.g. "20261114"
  serverRelativeUrl: string;
  timeLastModified?: string;
}

/** Metadata fields that must be filled before saving/printing. */
export type RequirableField = 'location' | 'reportNo' | 'subcontractor';

/** Dấu vết tạo/sửa báo cáo: email người dùng (do trình duyệt ghi, chỉ để hiển thị) và thời điểm ISO 8601. */
export interface ReportStamp {
  by: string;
  at: string;
}

/** Thông tin chung của báo cáo. Nội dung báo cáo nằm trong các mục (sections). */
export interface ReportMetadata {
  docTitle: string;             // e.g. "BÁO CÁO NGÀY" or "BÁO CÁO TUẦN"
  docTitleEn: string;           // e.g. "DAILY REPORT"
  subTitle: string;             // Tên dự án (một chuỗi, bản in tự xuống dòng)
  docCode: string;              // Mã tài liệu đã thay {unit} {reportNo}, e.g. "HCM2_1-23_CS-THAC_DCEM-CPZZ-ST01-CWS-DRP-00001"
  docCodePattern?: string;      // Mẫu mã tài liệu có {unit} và {reportNo} (từ template, không lưu trong file báo cáo)
  templateVersion?: string;     // e.g. "template_DRP_ST01_20261114"
  location: string;             // Luôn = mã đơn vị (ST01), tự điền, không sửa được
  reportNo: string;             // Số báo cáo 5 chữ số, tự tăng độc lập theo từng đơn vị, e.g. "00001"
  subcontractors: string[];     // Các đơn vị thi công được chọn trong báo cáo này
  email: string;                // Người lưu gần nhất (updated.by)
  updateTime: string;           // Thời điểm lưu gần nhất (updated.at, ISO 8601)
  created?: ReportStamp;        // Người/thời điểm lưu lần đầu (không đổi ở các lần lưu sau)
}

export type CellValue = string | number | undefined;
export type ColumnType = 'text' | 'number' | 'select' | 'calc' | 'date';
export type ColumnTone = 'danger' | 'warning' | 'success' | 'info' | 'muted';

/** Cột của bảng động = định nghĩa trong template ghép với nhãn của báo cáo (xem reportJson.ts). */
export interface TableColumn {
  code: string;                       // `id` cột, ví dụ "nos"
  label: string;                      // "Giờ công | Man-hrs"
  type: ColumnType;
  options?: string[];                 // danh sách chọn cố định (nếu không dùng danh mục)
  formula?: string;                   // calc: nos*hrs
  key?: boolean;                      // dòng isLocked => ô này chỉ đọc
  readonly?: boolean;                 // ô chỉ đọc ở mọi dòng
  total?: boolean;                    // có dòng Tổng
  align?: 'left' | 'center' | 'right';
  width?: number;                     // % bề rộng (tuỳ chọn)
  defaultValue?: string;              // giá trị khi thêm dòng; {unit}, {n5}
  tones?: Record<string, ColumnTone>; // tô màu theo giá trị (select)
  carry?: string;                     // công thức tính trên dòng của ngày trước (luỹ kế)
  reset?: boolean;                    // khi tạo từ ngày trước: về giá trị mặc định
  optionsRef?: string;                // tên danh mục (chữ thường): danh sách lấy từ danh mục lúc chạy
  optionsBy?: 'link';                 // dropdown liệt kê giá trị ràng buộc (vd tên vật tư); chọn thì cột from tự điền mã
  noNone?: boolean;                   // danh mục cố định: không có mục "---" đầu danh sách
  from?: string;                      // chọn mã ở cột này thì ô này tự điền giá trị ràng buộc của mã
}

export interface TableRow {
  id: string;
  isLocked?: boolean;
  values: Record<string, CellValue>;
}

export type FieldType = 'text' | 'textarea' | 'number' | 'select';

/** Trường văn bản đơn lẻ của một mục (ghi chú HSE, ký duyệt, chỉ đạo…). */
export interface FormField {
  key: string;
  label: string;                // "Sáng | AM"
  type: FieldType;
  options?: string[];
  full?: boolean;               // trải hết hàng, in dạng "Nhãn: giá trị"
  span?: number;                // số cột lưới chiếm (mặc định 1)
  rows?: number;                // số dòng của textarea
  lines?: number;               // số dòng kẻ của textarea khi in (UI vẫn hiện `rows` dòng)
  placeholder?: string;
  defaultValue?: string;        // giá trị khi trống; {user} {siteManager} {unit} được điền tự động
  required?: boolean;
  value: CellValue;
}

/** Mục dạng bảng (có thể kèm các trường văn bản `fields`, vd HSE). */
export interface ReportTable {
  kind: 'table';
  key: string;        // `id` mục: "manpower"
  title: string;      // "NHÂN LỰC"
  titleEn: string;    // "MANPOWER"
  shortTitle: string; // nhãn thanh nhảy nhanh
  summary: string;    // Tóm tắt | Summary
  showSummary: boolean; // có hiện ô Tóm tắt trên form nhập không
  single: boolean;    // bảng một dòng cố định: không thêm/xoá dòng
  columns: TableColumn[];
  rows: TableRow[];
  /** Trường văn bản của mục (in dạng "Nhãn: giá trị" dưới bảng). */
  fields?: FormField[];
  /** Số cột lưới của `fields`. */
  fieldColumns?: number;
}

/** Mục chỉ có các trường văn bản (ký duyệt, chỉ đạo…). */
export interface ReportForm {
  kind: 'form';
  key: string;
  title: string;
  titleEn: string;
  shortTitle: string;
  columns: number;              // số cột lưới (và số cột bảng khi in)
  sign: boolean;                // in dạng khung chữ ký
  summary: string;
  showSummary: boolean;
  fields: FormField[];
}

/** Nhóm ảnh của một bộ phận: tên bộ phận, tóm tắt và các file ảnh trong photos/<id>/. */
export interface PhotoGroup {
  id: string;           // tên thư mục con chứa ảnh
  department: string;   // "DCEM", "HSE"…
  summary: string;
  files: string[];      // tên file ảnh
}

/** Mục ảnh hiện trường theo bộ phận. */
export interface ReportPhotoSection {
  kind: 'photo';
  key: string;
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
  /** Các mục của báo cáo, theo thứ tự hiển thị. */
  sections: ReportSection[];
  tasks: ReportTaskRow[]; // Chỉ báo cáo tuần (placeholder)
  eTag?: string;
  isExisting?: boolean;
  copiedFromDate?: string; // e.g. "2026-11-10" when pre-filled from previous day
  latestTemplateInfo?: TemplateVersionInfo;
  templateUpdateRequired?: boolean;
  outdatedSourceType?: 'existing' | 'previous_day';
  outdatedTemplateVersion?: string;
  /** Số lớn nhất đã dùng của các mã tự tăng (Event ID, Issue ID…) ở các báo cáo khác của đơn vị, đọc từ sổ ReportRegistry; khoá "events.eventid". */
  idCounters?: Record<string, number>;
}

/** File PDF đã ký nằm trong thư mục báo cáo của một ngày. */
export interface SignedDocInfo {
  name: string;
  serverRelativeUrl: string;
  length: number;
  timeLastModified: string;
}

/** Trạng thái của một ngày: chưa lưu báo cáo / đã lưu nhưng thiếu PDF đã ký / hoàn tất. */
export type ReportDayStatus = 'missing' | 'unsigned' | 'done';

export interface ReportDayInfo {
  dateStr: string;            // YYYY-MM-DD
  status: ReportDayStatus;
  savedAt?: string;           // thời điểm lưu file báo cáo (.json)
  signed?: SignedDocInfo;     // file PDF đã ký (nếu có)
  reporter?: string;          // người lập (email) – từ sổ ReportRegistry
  reportNo?: string;          // số báo cáo – từ sổ ReportRegistry
}
