import { RoleEntry } from '../../../core/types';

export type ReportRoleAssignment = RoleEntry;

export interface TemplateVersionInfo {
  fileName: string;          // e.g. "template_ST01_20261109.xlsx"
  version: string;           // e.g. "template_ST01_20261109"
  versionDate: string;       // e.g. "20261109"
  serverRelativeUrl: string;
  timeLastModified?: string;
}

/** Metadata fields a template may flag as required (column `Required` of sheet 01_General). */
export type RequirableField = 'location' | 'reportNo' | 'shift' | 'maxTemp' | 'summary';

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
  // I. General & Weather
  weatherAM: string;            // "Nắng"
  weatherPM: string;            // "Mưa to"
  maxTemp: string;              // "33 °C"
  hrsLost?: number;             // 1.5
  progressEval: string;         // "Đạt" | "Chậm" | "Vượt"
  summary: string;              // I. Tóm tắt | Summary
  manpowerSummary?: string;     // II. Tóm tắt nhân lực
  equipmentSummary?: string;    // III. Tóm tắt thiết bị
  workQtySummary?: string;      // IV. Tóm tắt khối lượng
  materialsSummary?: string;    // V. Tóm tắt vật tư
  worksPlanSummary?: string;    // VI. Tóm tắt công việc & kế hoạch
  issuesSummary?: string;       // VII. Tóm tắt vướng mắc
  // VI. Works Today & Plan for Tomorrow
  worksToday: string;
  planTomorrow: string;
  // VIII. HSE – Quality – Engineer's Instructions
  hseIncidents?: number;
  hseInspections?: number;
  inspectionRequests?: number;
  newNCR?: number;
  photosCount?: number;
  photosFolder?: string;
  incidentDetails: string;
  engineerInstructions: string;
  planSummary: string;
  // Sign-off & Data Checks
  preparedBy: string;
  siteManager: string;
  tcjvRep: string;
  engineerAck: string;
  approvalComment: string;
}

export interface ManpowerRow {
  id: string;
  org: string;        // Đơn vị | Org.
  category: string;   // Loại | Category
  nos?: number;       // Số người | Nos.
  hrs?: number;       // Giờ/người | Hrs
  manHrs?: number;    // Giờ công | Man-hrs
  isLocked?: boolean; // Khoá không cho xoá / sửa tên đơn vị
}

export interface EquipmentRow {
  id: string;
  equipment: string;  // Thiết bị | Equipment
  qty?: number;       // SL | Qty
  hrs?: number;       // Giờ HĐ | Hrs
  status: string;     // Tình trạng | Status
  remark: string;     // Ghi chú | Remark
  isLocked?: boolean; // Khoá không cho xoá
}

export interface WorkQtyRow {
  id: string;
  activityId: string; // Activity ID
  item: string;       // Hạng mục | Item
  unit: string;       // ĐV | Unit
  plan?: number;      // KH ngày | Plan
  actual?: number;    // Thực hiện | Actual
  percent?: number;   // %
  remark: string;     // Ghi chú | Remark
  isLocked?: boolean; // Khoá không cho xoá / sửa mã Activity ID
}

export interface MaterialRow {
  id: string;
  material: string;   // Vật tư | Material
  unit: string;       // ĐV | Unit
  recv?: number;      // Nhập về | Recv.
  used?: number;      // Sử dụng | Used
  mirMar: string;     // Phiếu MIR/MAR
  isLocked?: boolean; // Khoá không cho xoá
}

export interface IssueRow {
  id: string;
  issueId: string;      // Mã | ID
  description: string;  // Nội dung | Description
  severity: string;     // Mức | Sev. ("Cao", "Trung bình", "Thấp")
  criticalPath: string; // Găng | CP ("Có", "–")
  owner: string;        // Bên xử lý | By
  dueDate: string;      // Hạn | Due
  status: string;       // Trạng thái ("Mở", "Đang xử lý", "Đóng")
  isLocked?: boolean;   // Khoá không cho xoá
}

export interface ReportPhotoFile {
  name: string;
  serverRelativeUrl: string;
  timeLastModified: string;
  length: number;
}

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
  manpower: ManpowerRow[];
  equipment: EquipmentRow[];
  workItems: WorkQtyRow[];
  materials: MaterialRow[];
  issues: IssueRow[];
  tasks: ReportTaskRow[]; // Used for WeeklyReport or legacy tree-table
  photos?: ReportPhotoFile[];
  eTag?: string;
  isExisting?: boolean;
  copiedFromDate?: string; // e.g. "2026-11-10" when pre-filled from previous day
  latestTemplateInfo?: TemplateVersionInfo;
  templateUpdateRequired?: boolean;
  outdatedSourceType?: 'existing' | 'previous_day';
  outdatedTemplateVersion?: string;
}
