/**
 * Định dạng file lưu báo cáo (`<DRP|WRP>_<yyyyMMdd>_<Unit>.json`) và file template (`template_DRP_<Unit>_<yyyyMMdd>.json`).
 * Hợp đồng đầy đủ: docs/REPORT_JSON_PLAN.md và mẫu docs/sharepoint_setup/DailyReportLibrary/_template/template_DRP_ST01_<yyyyMMdd>.json.
 *
 * Nguyên tắc: file báo cáo chỉ chứa định danh, nhãn và GIÁ TRỊ THẬT (kể cả ô đã tính và dòng Tổng) để mở một mình đọc được.
 * Cách NHẬP/TRÌNH BÀY (kiểu ô, danh mục dropdown, công thức, mặc định, căn lề, độ rộng, bố cục) do CODE quyết định
 * (utils/reportUiSchema.ts), ghép với báo cáo theo `id` mục / `id` cột / `id` trường. File template JSON có CÙNG định dạng
 * với file báo cáo (`ReportJson`), chỉ khác dữ liệu: `unit`, `date`, `reportNo`, `docCode` chỉ là chỗ giữ (cấp lại khi import), `docCodeTemplate`
 * là mẫu mã tài liệu có {unit} {reportNo}, trường có thể chứa token {user} {siteManager} {unit}. `ReportTemplateJson` bên dưới là dạng đã mở rộng
 * trong bộ nhớ (template + bảng UI) mà phần còn lại của code dùng.
 */

export const REPORT_JSON_SCHEMA = 1;

export interface LabelJson {
  vi: string;
  en?: string;
}

export interface SectionTitleJson extends LabelJson {
  /** Nhãn ngắn cho thanh nhảy nhanh trên web. */
  short?: string;
}

/* ------------------------------ File báo cáo ------------------------------ */

export type JsonCellValue = string | number;

export interface ReportJsonColumn {
  id: string;
  label: LabelJson;
}

export interface ReportJsonRow {
  /** Duy nhất trong từng mục (r1, r2…), không bao giờ đổi. */
  id: string;
  /** Chỉ ghi khi true: dòng mẫu của template, không được xoá. */
  locked?: true;
  cells: Record<string, JsonCellValue>;
  /** Tên file PDF đính kèm của dòng (mục có `rowPdf`); chỉ ghi khi có file. Không đặt tên `files` vì dòng ảnh dùng khoá đó để phân biệt. */
  pdf?: string[];
}

/** Ô văn bản đơn lẻ gắn với cả mục (ghi chú HSE, ký duyệt, chỉ đạo…). */
export interface ReportJsonField {
  id: string;
  label: LabelJson;
  value: JsonCellValue;
}

/** Một bộ phận trong mục ảnh; `id` đồng thời là tên thư mục `photos/<id>/`. */
export interface ReportJsonPhotoRow {
  id: string;
  department: string;
  summary: string;
  files: string[];
}

export interface ReportJsonSection {
  id: string;
  title: SectionTitleJson;
  /** Chỉ có `"photo"` cho mục ảnh; mục khác không ghi. */
  type?: 'photo';
  /** Luôn có; trống thì "". */
  summary: string;
  columns?: ReportJsonColumn[];
  rows?: Array<ReportJsonRow | ReportJsonPhotoRow>;
  /** Giá trị dòng Tổng theo `column.id` (có thì in dòng Tổng). */
  totals?: Record<string, number>;
  fields?: ReportJsonField[];
}

export interface ReportJsonStamp {
  by: string;
  at: string;
}

export interface ReportJson {
  schema: number;
  unit: string;
  docId: string;
  /** yyyy-MM-dd */
  date: string;
  reportNo: number;
  /** Mẫu mã tài liệu có chỗ giữ {unit} {reportNo}; docCode = mẫu này đã thay. Có trong file lưu để file tự dùng làm template được. */
  docCodeTemplate: string;
  docCode: string;
  title: { vi: string; en: string; sub: string };
  /** Tên template đã dùng (vd template_DRP_ST01_20261114): so với bản mới nhất để báo cập nhật. */
  template: string;
  created: ReportJsonStamp;
  updated: ReportJsonStamp;
  subcontractors: string[];
  sections: ReportJsonSection[];
  /** Chỉ báo cáo tuần (placeholder). */
  tasks?: Array<{ id: string; taskName: string; completionRate?: number; quantity?: number; unit: string; parentId?: string; level: number; isLeaf: boolean }>;
}

/* ------------- Template đã mở rộng (file template + bảng UI trong code) ------------- */

export type TemplateColumnType = 'text' | 'number' | 'date' | 'select' | 'calc';
export type TemplateTone = 'danger' | 'warning' | 'success' | 'info' | 'muted';

export interface TemplateColumnJson {
  id: string;
  label: LabelJson;
  /** Mặc định `text`. */
  type?: TemplateColumnType;
  /** Tên danh mục (Element_ID, weather…) cho cột select: danh sách lấy lúc chạy từ list DRPCatalogItems. */
  catalog?: string;
  /** `link`: dropdown liệt kê giá trị ràng buộc (tên vật tư…) thay vì mã; chọn thì cột `from` tự điền mã. */
  catalogBy?: 'link';
  /** Chọn mã ở cột này thì ô này tự điền giá trị ràng buộc của mã và bị khoá. */
  from?: string;
  /** Với `from`: lấy cột này của mã (mặc định `link`) và ghi giá trị người dùng nhập cho mã mới vào cột đó. */
  fromField?: 'note' | 'nameVn';
  /** Danh sách chọn cố định (nếu không dùng danh mục). */
  options?: string[];
  /** true = không có mục "---" ở đầu danh sách (danh mục cố định như thời tiết). */
  noNone?: boolean;
  formula?: string;
  readonly?: boolean;
  /** Ô chỉ đọc khi dòng bị khoá (cờ `key` trước đây). */
  key?: boolean;
  total?: boolean;
  align?: 'left' | 'center' | 'right';
  width?: number;
  default?: string | number;
  /** Khi tạo từ báo cáo ngày trước: về giá trị mặc định. */
  reset?: boolean;
  /** Khi tạo từ báo cáo ngày trước: giá trị = công thức tính trên dòng ngày trước (vd daily+prevCum). */
  carry?: string;
  tones?: Record<string, TemplateTone>;
}

export interface TemplateFieldJson {
  id: string;
  label: LabelJson;
  /** Mặc định `text`. */
  type?: 'text' | 'textarea' | 'number' | 'select';
  options?: string[];
  full?: boolean;
  span?: number;
  rows?: number;
  lines?: number;
  placeholder?: string;
  /** Có thể dùng {user} {siteManager} {unit}; chỉ điền khi trường đang trống. */
  default?: string | number;
  /**
   * Token (vd {user}, {siteManager}) luôn tính lại cho báo cáo MỚI (từ template hoặc kế thừa ngày trước), bỏ qua giá trị đã
   * lưu trong template/báo cáo trước. Mở lại báo cáo đã lưu thì vẫn giữ giá trị đã lưu. Khai báo trong utils/reportUiSchema.ts.
   */
  auto?: string;
}

export interface TemplateSectionJson {
  id: string;
  title: SectionTitleJson;
  kind?: 'photo';
  /** Bảng một dòng cố định: không thêm/xoá dòng. */
  single?: boolean;
  /** true = mỗi dòng đính kèm được file PDF (thuộc tính UI do code quyết định, không nằm trong file template). */
  rowPdf?: boolean;
  /** false = ẩn ô Tóm tắt trên form nhập (mặc định hiện). */
  showSummary?: boolean;
  /** Bố cục các `fields`: số cột lưới; `sign` = in dạng khung chữ ký. */
  layout?: { columns: number; sign?: boolean };
  /** Số ảnh mỗi hàng (mục ảnh). */
  perRow?: number;
  columns?: TemplateColumnJson[];
  /** Dòng mặc định khi tạo báo cáo từ template. */
  rows?: ReportJsonRow[];
  fields?: TemplateFieldJson[];
}

export interface ReportTemplateJson {
  schema: number;
  docId: string;
  title: { vi: string; en: string; sub: string };
  /** Mẫu mã tài liệu: {unit} = mã đơn vị, {reportNo} = số báo cáo 5 chữ số. */
  docCodePattern: string;
  sections: TemplateSectionJson[];
}
