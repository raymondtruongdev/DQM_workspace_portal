import { TemplateColumnJson, TemplateFieldJson, TemplateSectionJson } from '../types/reportJson.types';
import { REPORT_FEATURES } from '../constants/reportFeatures';

/**
 * Cách NHẬP/TRÌNH BÀY của báo cáo do CODE quyết định (file này), không nằm trong file template JSON.
 * Template JSON (`_template/template_*.json`) và file báo cáo chỉ chứa định danh, nhãn và giá trị:
 * template = khung báo cáo rỗng (mục, cột, nhãn, dòng mẫu, giá trị khởi tạo của trường).
 *
 * Bảng dưới khai báo theo `id` mục → `id` cột / `id` trường: kiểu ô, danh mục dropdown, công thức, mặc định, luỹ kế, căn lề,
 * độ rộng, bố cục. Mục/cột/trường có trong template mà không khai báo ở đây hiển thị như văn bản thường.
 * Thêm cột chỉ-văn-bản chỉ cần sửa template; cột cần dropdown/số/công thức thì thêm một dòng ở đây (và build lại).
 * Template được mở rộng thành `ReportTemplateJson` bằng `expandTemplate` (reportJson.ts).
 */

export type ColumnUi = Omit<TemplateColumnJson, 'id' | 'label'>;
export type FieldUi = Omit<TemplateFieldJson, 'id' | 'label' | 'default'>;

export interface SectionUi extends Pick<TemplateSectionJson, 'kind' | 'variant' | 'single' | 'rowPdf' | 'showSummary' | 'layout' | 'perRow'> {
  columns?: Record<string, ColumnUi>;
  fields?: Record<string, FieldUi>;
}

/** Bề rộng tối thiểu (px) của cột Activity ID (P6): đủ hiện trọn mã dài nhất như CON-ST01-ARC-TBD-00010 cùng mũi tên dropdown. */
const ACTIVITY_COL_PX = 270;

/** Daily Report (`docId = DailyReport`). */
const DAILY_REPORT_UI: Record<string, SectionUi> = {
  weather: {
    single: true,
    showSummary: false,
    columns: {
      am: { type: 'select', catalog: 'weather', reset: true, default: '---', align: 'center', minPx: 110 },
      pm: { type: 'select', catalog: 'weather', reset: true, default: '---', align: 'center', minPx: 110 },
      night: { type: 'select', catalog: 'weather', reset: true, default: '---', align: 'center', minPx: 110 },
      hrsLost: { type: 'number', reset: true, default: 0, align: 'center', minPx: 110 },
      value: { reset: true, default: 'SV | EV | AC', align: 'left', width: 40 },
    },
  },
  hse: {
    single: true,
    showSummary: false,
    layout: { columns: 1 },
    columns: {
      hseIncidents: { type: 'number', align: 'center' },
      envIncidents: { type: 'number', align: 'center' },
      hseInspections: { type: 'number', align: 'center' },
      inspectionRequests: { type: 'number', align: 'center' },
      newNcr: { type: 'number', align: 'center' },
    },
    fields: {
      incidentDetails: { full: true },
      engineerInstructions: { type: 'textarea', full: true, rows: 2 },
    },
  },
  events: {
    rowPdf: REPORT_FEATURES.EVENT_PDF, // cờ trong constants/reportFeatures.ts
    columns: {
      eventId: { readonly: true, default: 'EVT-{unit}-{n5}', align: 'center', minPx: 150, fit: true },
      type: { type: 'select', catalog: 'event_type', minPx: 180, fit: true, default: '---' },
      description: { width: 40 },
      due: { type: 'date', width: 10 },
      status: { type: 'select', catalog: 'event_status', default: '---', align: 'center', minPx: 140, fit: true },
      remark: { width: 20 },
    },
  },
  workQty: {
    columns: {
      activityId: { type: 'select', catalog: 'ActivityP6_ID', from: 'element', default: '---', align: 'left', minPx: ACTIVITY_COL_PX },
      element: { type: 'select', catalog: 'Element_ID', noNone: true, align: 'center', width: 18 },
      step: { width: 22 },
      unit: { default: 'm³', align: 'center', width: 6 },
      start: { type: 'date', width: 10 },
      finish: { type: 'date', width: 10 },
      daily: { type: 'number', reset: true, default: 0, width: 9 },
      prevCum: { type: 'number', readonly: true, default: 0, carry: 'daily+prevCum', width: 9 },
      planned: { type: 'number', key: true, default: 0, width: 9 },
      progress: { type: 'calc', formula: 'round((daily+prevCum)/planned*100)', width: 8 },
      remark: { width: 16 },
    },
  },
  materials: {
    columns: {
      matId: { type: 'select', catalog: 'Material_ID', noNone: true, align: 'left', askName: 'nameVn', minPx: 290, fit: true },
      unit: { default: 'm³', align: 'center', minPx: 70, fit: true },
      recv: { type: 'number', default: 0, minPx: 140, fit: true },
      supplier: { type: 'select', catalog: 'supplier', default: '---', minPx: 130, fit: true },
      remark: { minPx: 240 },
    },
  },
  manpower: {
    columns: {
      resId: { type: 'select', catalog: 'ResHuman_ID', noNone: true, align: 'center', minPx: 100, fit: true },
      category: { readonly: true, from: 'resId', fromField: 'nameVn', minPx: 240, fit: true },
      daily: { type: 'number', reset: true, total: true, default: 0, minPx: 130, fit: true },
      prevCum: { type: 'number', readonly: true, total: true, default: 0, carry: 'daily+prevCum', minPx: 130, fit: true },
      cumTotal: { type: 'calc', formula: 'daily+prevCum', total: true, minPx: 130, fit: true },
      remark: { minPx: 240 },
    },
  },
  equipment: {
    columns: {
      activityId: { type: 'select', catalog: 'ActivityP6_ID', default: '---', align: 'left', minPx: ACTIVITY_COL_PX, fit: true },
      resId: { type: 'select', catalog: 'ResEquip_ID', noNone: true, align: 'center', minPx: 220, fit: true },
      equipment: { readonly: true, from: 'resId', fromField: 'nameVn', minPx: 260, fit: true },
      plate: { minPx: 170, placeholder: '51H12345, 62A11223' }, // Biển số: người dùng tự nhập (gợi ý viết liền, không dấu; nhiều biển cách nhau dấu phẩy)
      qty: { type: 'number', default: 1, minPx: 90, fit: true },
      status: { type: 'select', catalog: 'equipment_status', default: '---', align: 'center', minPx: 130, fit: true },
    },
  },
  issues: {
    columns: {
      activityId: { type: 'select', catalog: 'ActivityP6_ID', default: '---', align: 'left', minPx: ACTIVITY_COL_PX },
      issueId: { readonly: true, default: 'ISS-{unit}-{n5}', align: 'center', width: 20 },
      description: { width: 34 },
      severity: { type: 'select', catalog: 'issue_severity', default: '---', align: 'center', width: 9 },
      owner: { width: 13 },
      raised: { type: 'date', width: 10 },
      due: { type: 'date', width: 10 },
      status: { type: 'select', catalog: 'issue_status', default: '---', align: 'center', width: 10 },
      todayUpdate: { default: '---', width: 22 },
    },
  },
  layout: {
    kind: 'photo',
    variant: 'layout',
    showSummary: false,
    perRow: 1,
  },
  photos: {
    kind: 'photo',
    showSummary: false,
    perRow: 2,
  },
  signOff: {
    showSummary: false,
    layout: { columns: 2, sign: true },
    // Luôn là người đang lập / trưởng ga hiện tại khi tạo báo cáo mới, kể cả khi dùng báo cáo đã lưu làm template
    fields: {
      preparedBy: { auto: '{user}' },
      siteManager: { auto: '{siteManager}' },
    },
  },
  directives: {
    showSummary: false,
    layout: { columns: 1 },
    fields: {
      directives: { type: 'textarea', full: true, rows: 3, lines: 15, placeholder: 'Cấp trên nhập chỉ đạo (bản in kẻ 15 dòng)' },
    },
  },
};

/** Báo cáo tuần chưa có template (placeholder). */
export const REPORT_UI: Record<string, Record<string, SectionUi>> = {
  DailyReport: DAILY_REPORT_UI
};
