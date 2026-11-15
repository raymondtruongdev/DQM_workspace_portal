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

export interface SectionUi extends Pick<TemplateSectionJson, 'kind' | 'single' | 'rowPdf' | 'showSummary' | 'layout' | 'perRow'> {
  columns?: Record<string, ColumnUi>;
  fields?: Record<string, FieldUi>;
}

/** Daily Report (`docId = DailyReport`). */
const DAILY_REPORT_UI: Record<string, SectionUi> = {
  weather: {
    single: true,
    showSummary: false,
    columns: {
      am: { type: 'select', catalog: 'weather', noNone: true, reset: true, default: 'Nắng', align: 'center', width: 13 },
      pm: { type: 'select', catalog: 'weather', noNone: true, reset: true, default: 'Nắng', align: 'center', width: 13 },
      night: { type: 'select', catalog: 'weather', noNone: true, reset: true, default: 'Nắng', align: 'center', width: 13 },
      hrsLost: { type: 'number', reset: true, default: 0, align: 'center', width: 15 },
      value: { reset: true, default: 'SV | EV | AC', align: 'center', width: 30 },
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
      eventId: { readonly: true, default: 'EVT-{unit}-{n5}', align: 'center', width: 20 },
      type: { type: 'select', catalog: 'event_type', noNone: true, width: 18, default: 'KIEM_TRA_ATLD' },
      description: { width: 40 },
      due: { type: 'date', width: 10 },
      status: { type: 'select', catalog: 'event_status', noNone: true, default: 'Không Đạt', align: 'center', width: 10 },
      remark: { width: 20 },
    },
  },
  workQty: {
    columns: {
      activityId: { type: 'select', catalog: 'ActivityP6_ID', from: 'element', default: '---', align: 'center', width: 14 },
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
      matId: { type: 'select', catalog: 'Material_ID', noNone: true, align: 'center', width: 16 },
      material: { readonly: true, from: 'matId', fromField: 'nameVn', width: 30 },
      unit: { default: 'm³', align: 'center', width: 8 },
      recv: { type: 'number', default: 0, width: 14 },
      supplier: { type: 'select', catalog: 'supplier', noNone: true, width: 20, default: 'Thiên An' },
      remark: { width: 20 },
    },
  },
  manpower: {
    columns: {
      resId: { type: 'select', catalog: 'ResHuman_ID', noNone: true, align: 'center', width: 16 },
      category: { readonly: true, from: 'resId', fromField: 'nameVn', width: 28 },
      daily: { type: 'number', reset: true, total: true, default: 0, width: 13 },
      prevCum: { type: 'number', readonly: true, total: true, default: 0, carry: 'daily+prevCum', width: 13 },
      cumTotal: { type: 'calc', formula: 'daily+prevCum', total: true, width: 13 },
      remark: { width: 20 },
    },
  },
  equipment: {
    columns: {
      activityId: { type: 'select', catalog: 'ActivityP6_ID', default: '---', align: 'center', width: 16 },
      resId: { type: 'select', catalog: 'ResEquip_ID', noNone: true, align: 'center', width: 22 },
      equipment: { readonly: true, from: 'resId', fromField: 'nameVn', width: 34 },
      qty: { type: 'number', default: 1, width: 10 },
      status: { type: 'select', catalog: 'equipment_status', noNone: true, default: 'Active', align: 'center', width: 14 },
    },
  },
  issues: {
    columns: {
      activityId: { type: 'select', catalog: 'ActivityP6_ID', default: '---', align: 'center', width: 14 },
      issueId: { readonly: true, default: 'ISS-{unit}-{n5}', align: 'center', width: 20 },
      description: { width: 34 },
      severity: { type: 'select', catalog: 'issue_severity', noNone: true, default: 'Trung bình', align: 'center', width: 9 },
      owner: { width: 13 },
      raised: { type: 'date', width: 10 },
      due: { type: 'date', width: 10 },
      status: { type: 'select', catalog: 'issue_status', noNone: true, default: 'Mở', align: 'center', width: 10 },
      todayUpdate: { default: '---', width: 22 },
    },
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
