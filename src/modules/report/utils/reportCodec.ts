import type * as XLSXTypes from 'xlsx';
import {
  EquipmentRow,
  IssueRow,
  ManpowerRow,
  MaterialRow,
  ReportDoc,
  ReportMetadata,
  ReportSectionKey,
  ReportTaskRow,
  RequirableField,
  WorkQtyRow
} from '../types/report.types';
import { recalculateAllRates } from './reportTaskTree';
import { TABLE_SECTIONS, SECTION_SHEET_NAME, findSectionSheet, hasSection } from './reportSections';
import { REQUIRABLE_FIELDS, getRequiredFields } from './reportValidation';

function getSectionSheet(wb: XLSXTypes.WorkBook, key: ReportSectionKey): XLSXTypes.WorkSheet | undefined {
  const name = findSectionSheet(wb.SheetNames, key);
  return name ? wb.Sheets[name] : undefined;
}

async function getXlsx(): Promise<typeof XLSXTypes> {
  return await import(/* webpackChunkName: 'xlsx' */ 'xlsx');
}

/**
 * Default seed data for a new Daily Report (based on Mau_BCN_ST06_20261025.pdf)
 */
export function getDefaultDailyReportSections(unit: string): {
  manpower: ManpowerRow[];
  equipment: EquipmentRow[];
  workItems: WorkQtyRow[];
  materials: MaterialRow[];
  issues: IssueRow[];
} {
  const cleanUnit = (unit || 'ST01').toUpperCase();

  if (cleanUnit === 'ST02') {
    return {
      manpower: [
        { id: 'mp-1', org: 'TCJV', category: 'Ban chỉ huy Ga ST02', nos: 7, hrs: 10, manHrs: 70, isLocked: true },
        { id: 'mp-2', org: 'TCJV', category: 'Kỹ sư Địa kỹ thuật & Quan trắc', nos: 12, hrs: 10, manHrs: 120, isLocked: true },
        { id: 'mp-3', org: 'NTP-TV01', category: 'Công nhân khoan cọc & kết cấu ST02', nos: 38, hrs: 10, manHrs: 380, isLocked: true },
        { id: 'mp-4', org: 'NTP-TV01', category: 'Lao động phổ thông', nos: 20, hrs: 10, manHrs: 200, isLocked: true },
        { id: 'mp-5', org: 'TCJV', category: 'An toàn viên (HSE)', nos: 4, hrs: 10, manHrs: 40, isLocked: true },
        { id: 'mp-6', org: 'NTP-CD02', category: 'Đội di dời điện/viễn thông', nos: 8, hrs: 8, manHrs: 64, isLocked: false }
      ],
      equipment: [
        { id: 'eq-1', equipment: 'Máy khoan cọc nhồi Bauer BG28', qty: 1, hrs: 10, status: 'Hoạt động', remark: 'Khoan cọc KP-08', isLocked: false },
        { id: 'eq-2', equipment: 'Cẩu phục vụ 55T', qty: 1, hrs: 9, status: 'Hoạt động', remark: '', isLocked: false },
        { id: 'eq-3', equipment: 'Máy đào bánh xích 1,2 m³', qty: 2, hrs: 8.5, status: 'Hoạt động', remark: 'Đào hố móng trục 1–3', isLocked: false },
        { id: 'eq-4', equipment: 'Máy nén khí & rửa lỗ khoan', qty: 1, hrs: 6, status: 'Hoạt động', remark: '', isLocked: false }
      ],
      workItems: [
        { id: 'wq-1', activityId: 'ST02-DW-1120', item: 'Cọc khoan nhồi D1500 – khoan tạo lỗ', unit: 'm', plan: 42, actual: 42, percent: 100, remark: 'Cọc KP-08', isLocked: true },
        { id: 'wq-2', activityId: 'ST02-DW-1130', item: 'Gia công & hạ lồng thép cọc KP-08', unit: 'tấn', plan: 18, actual: 18, percent: 100, remark: '', isLocked: true },
        { id: 'wq-3', activityId: 'ST02-GW-2010', item: 'Đổ bê tông tường dẫn trục 6–8', unit: 'm3', plan: 65, actual: 65, percent: 100, remark: '', isLocked: false },
        { id: 'wq-4', activityId: 'ST02-EX-3010', item: 'Đào đất mở lớp 1 (trục 1–3)', unit: 'm3', plan: 320, actual: 290, percent: 91, remark: '', isLocked: false }
      ],
      materials: [
        { id: 'mat-1', material: 'Bê tông C30/37 cọc khoan nhồi', unit: 'm3', recv: 82, used: 82, mirMar: 'MIR-ST02-0064' },
        { id: 'mat-2', material: 'Thép D28/D25 cọc nhồi', unit: 'tấn', recv: 22, used: 18, mirMar: 'MIR-ST02-0065' }
      ],
      issues: [
        {
          id: 'iss-1',
          issueId: 'VM-ST02-004',
          description: 'Chờ cắt điện tuyến cáp trung thế 22kV trục 7 Ga ST02',
          severity: 'Cao',
          criticalPath: 'Có',
          owner: 'EVNHCMC / TCJV',
          dueDate: '14/11',
          status: 'Đang xử lý'
        }
      ]
    };
  }

  return {
    manpower: [
      { id: 'mp-1', org: 'TCJV', category: `Quản lý hiện trường ${cleanUnit}`, nos: 8, hrs: 10, manHrs: 80, isLocked: true },
      { id: 'mp-2', org: 'TCJV', category: 'Kỹ sư/Kỹ thuật', nos: 14, hrs: 10, manHrs: 140, isLocked: true },
      { id: 'mp-3', org: 'NTP-TV01', category: `Công nhân kỹ thuật ${cleanUnit}`, nos: 52, hrs: 10, manHrs: 520, isLocked: true },
      { id: 'mp-4', org: 'NTP-TV01', category: 'Lao động phổ thông', nos: 30, hrs: 10, manHrs: 300, isLocked: true },
      { id: 'mp-5', org: 'TCJV', category: 'An toàn viên', nos: 6, hrs: 10, manHrs: 60, isLocked: true },
      { id: 'mp-6', org: 'NTP-TB02', category: 'Vận hành máy', nos: 4, hrs: 8, manHrs: 32, isLocked: false }
    ],
    equipment: [
      { id: 'eq-1', equipment: 'Cẩu bánh xích 80T', qty: 1, hrs: 10, status: 'Hoạt động', remark: '', isLocked: false },
      { id: 'eq-2', equipment: 'Gầu ngoạm tường vây (grab)', qty: 1, hrs: 8.5, status: 'Hoạt động', remark: 'Dừng 1,5h do mưa', isLocked: false },
      { id: 'eq-3', equipment: 'Trạm trộn/tái sinh bentonite', qty: 1, hrs: 12, status: 'Hoạt động', remark: '', isLocked: false },
      { id: 'eq-4', equipment: 'Máy đào 0,7 m³', qty: 2, hrs: 9, status: 'Hoạt động', remark: '', isLocked: false },
      { id: 'eq-5', equipment: 'Xe bồn bê tông 7 m³', qty: 6, hrs: 4, status: 'Hoạt động', remark: 'Đổ BT tấm P-12', isLocked: false },
      { id: 'eq-6', equipment: 'Máy phát điện 250 kVA', qty: 1, hrs: 0, status: 'Chờ việc', remark: 'Dự phòng', isLocked: false }
    ],
    workItems: [
      { id: 'wq-1', activityId: `${cleanUnit}-DW-1120`, item: 'Tường vây – đào tấm P-13', unit: 'm3', plan: 160, actual: 142, percent: 89, remark: 'Mưa to 14:30–16:00', isLocked: true },
      { id: 'wq-2', activityId: `${cleanUnit}-DW-1130`, item: 'Tường vây – lồng thép tấm P-12', unit: 'tấn', plan: 28, actual: 28, percent: 100, remark: '', isLocked: true },
      { id: 'wq-3', activityId: `${cleanUnit}-DW-1140`, item: 'Tường vây – bê tông tấm P-12', unit: 'm3', plan: 210, actual: 214, percent: 102, remark: '', isLocked: false },
      { id: 'wq-4', activityId: `${cleanUnit}-GW-1010`, item: 'Tường dẫn trục 8–10', unit: 'm', plan: 12, actual: 9, percent: 75, remark: 'Chờ ván khuôn', isLocked: false }
    ],
    materials: [
      { id: 'mat-1', material: 'Bê tông C35/45 tường vây', unit: 'm3', recv: 214, used: 214, mirMar: `MIR-${cleanUnit}-0187` },
      { id: 'mat-2', material: 'Thép D32 (lồng P-14)', unit: 'tấn', recv: 32, used: 0, mirMar: `MIR-${cleanUnit}-0188` },
      { id: 'mat-3', material: 'Bentonite', unit: 'tấn', recv: 4, used: 3.5, mirMar: 'MAR chờ PICC – NCC mới' }
    ],
    issues: [
      {
        id: 'iss-1',
        issueId: `VM-${cleanUnit}-007`,
        description: 'Cọc secant lấn ranh 300–450 mm tại cạnh Đông – chờ phương án xử lý của thiết kế',
        severity: 'Cao',
        criticalPath: 'Có',
        owner: 'Thiết kế TCJV / TVTK',
        dueDate: '28/10',
        status: 'Đang xử lý'
      },
      {
        id: 'iss-2',
        issueId: `VM-${cleanUnit}-009`,
        description: 'Phê duyệt NCC bentonite mới chưa có – còn tồn 4 ngày sử dụng',
        severity: 'Cao',
        criticalPath: 'Có',
        owner: 'Mua sắm / PICC',
        dueDate: '27/10',
        status: 'Mở'
      },
      {
        id: 'iss-3',
        issueId: `VM-${cleanUnit}-010`,
        description: 'Cáp ngầm chưa xác định tại trục 9 – cần dò tìm trước khi đào tường dẫn',
        severity: 'Trung bình',
        criticalPath: '–',
        owner: 'Site + EVNHCMC',
        dueDate: '30/10',
        status: 'Mở'
      },
      {
        id: 'iss-4',
        issueId: `VM-${cleanUnit}-011`,
        description: 'Thiếu mặt bằng tập kết lồng thép do phân luồng giao thông giai đoạn 2',
        severity: 'Trung bình',
        criticalPath: '–',
        owner: 'Ban QLĐSĐT / CSGT',
        dueDate: '05/11',
        status: 'Mở'
      }
    ]
  };
}

function parseOptionalNum(val: unknown): number | undefined {
  if (val === undefined || val === null || String(val).trim() === '') return undefined;
  const cleaned = String(val).replace(',', '.').replace(/[^0-9.-]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? undefined : n;
}

function parseIsLocked(val: unknown, fallback: boolean): boolean {
  if (val !== undefined && val !== null && String(val).trim() !== '') {
    const s = String(val).trim().toLowerCase();
    if (s === 'true' || s === '1' || s === 'yes' || s === 'x' || s === 'khoá' || s === 'khóa' || s === 'locked') {
      return true;
    }
    if (s === 'false' || s === '0' || s === 'no') {
      return false;
    }
  }
  return fallback;
}

/**
 * Parse an Excel buffer (either 6-sheet template/report or legacy 1-sheet template) into ReportDoc
 */
export async function parseReportDocFromBuffer(
  buffer: ArrayBuffer,
  docId: string,
  unit: string,
  dateStr: string
): Promise<ReportDoc> {
  const XLSX = await getXlsx();
  const wb = XLSX.read(buffer, { type: 'array' });

  const defaults = getDefaultDailyReportSections(unit);

  const metadata: ReportMetadata = {
    docTitle: docId === 'WeeklyReport' ? 'BÁO CÁO TUẦN' : 'BÁO CÁO NGÀY',
    docTitleEn: docId === 'WeeklyReport' ? 'WEEKLY REPORT' : 'DAILY REPORT',
    subTitle: 'Tuyến Metro số 2 – Gói thầu M2.1 Tham Lương – Bến Thành | HCMC Metro Line 2 – Package M2.1',
    docCode: `HCM2_1-23_CS-MAUR_TCJV-CPZZ-${unit}-CWS-RP-041-S3-C01_BCN`,
    templateVersion: '',
    location: `${unit} – Ga ${unit}`,
    reportNo: '001',
    shift: 'Cả ngày (2 ca)',
    email: '',
    updateTime: '',
    weatherAM: 'Nắng',
    weatherPM: 'Nắng',
    maxTemp: '33 °C',
    hrsLost: 0,
    progressEval: 'Đạt',
    summary: '',
    manpowerSummary: '',
    equipmentSummary: '',
    workQtySummary: '',
    materialsSummary: '',
    worksPlanSummary: '',
    issuesSummary: '',
    worksToday: '',
    planTomorrow: '',
    hseIncidents: 0,
    hseInspections: 0,
    inspectionRequests: 0,
    newNCR: 0,
    photosCount: 0,
    incidentDetails: 'Không có | None',
    engineerInstructions: '',
    planSummary: '',
    preparedBy: '',
    siteManager: '',
    tcjvRep: '',
    engineerAck: '',
    approvalComment: '',
  };

  const hasMultiSheet = wb.SheetNames.some((s) => s === '01_General' || s === '02_Manpower');

  if (hasMultiSheet) {
    // 1. Parse 01_General
    const wsGen = wb.Sheets['01_General'] || wb.Sheets[wb.SheetNames[0]];
    const requiredFields: RequirableField[] = [];
    let hasRequiredColumn = false;
    if (wsGen) {
      const rows: unknown[][] = XLSX.utils.sheet_to_json(wsGen, { header: 1, defval: '' });
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;
        const key = String(row[0] || '').trim().toLowerCase();
        const val = String(row[1] ?? '').trim();
        if (!key || key === 'key') continue;
        const reqEntry = REQUIRABLE_FIELDS[key];
        if (reqEntry && row[3] !== undefined && String(row[3]).trim() !== '') {
          hasRequiredColumn = true;
          if (parseIsLocked(row[3], false)) requiredFields.push(reqEntry.field);
        }

        if (key === 'doctile' || key === 'doctitle') metadata.docTitle = val || metadata.docTitle;
        else if (key === 'doctileen' || key === 'doctitleen') metadata.docTitleEn = val || metadata.docTitleEn;
        else if (key === 'subtile' || key === 'subtitle') metadata.subTitle = val || metadata.subTitle;
        else if (key === 'doccode') metadata.docCode = val || metadata.docCode;
        else if (key === 'templateversion' || key === 'template_version' || key === 'templatefile') metadata.templateVersion = val;
        else if (key === 'location') metadata.location = val || metadata.location;
        else if (key === 'reportno') metadata.reportNo = val;
        else if (key === 'shift') metadata.shift = val || metadata.shift;
        else if (key === 'email') metadata.email = val;
        else if (key === 'updatetime') metadata.updateTime = val;
        else if (key === 'weatheram') metadata.weatherAM = val;
        else if (key === 'weatherpm') metadata.weatherPM = val;
        else if (key === 'maxtemp') metadata.maxTemp = val;
        else if (key === 'hrslost') metadata.hrsLost = parseOptionalNum(row[1]);
        else if (key === 'progresseval') metadata.progressEval = val;
        else if (key === 'summary') metadata.summary = val;
        else if (key === 'manpowersummary') metadata.manpowerSummary = val;
        else if (key === 'equipmentsummary') metadata.equipmentSummary = val;
        else if (key === 'workqtysummary') metadata.workQtySummary = val;
        else if (key === 'materialssummary') metadata.materialsSummary = val;
        else if (key === 'worksplansummary') metadata.worksPlanSummary = val;
        else if (key === 'issuessummary') metadata.issuesSummary = val;
        else if (key === 'workstoday') metadata.worksToday = val;
        else if (key === 'plantomorrow') metadata.planTomorrow = val;
        else if (key === 'hseincidents') metadata.hseIncidents = parseOptionalNum(row[1]) ?? 0;
        else if (key === 'hseinspections') metadata.hseInspections = parseOptionalNum(row[1]) ?? 0;
        else if (key === 'inspectionrequests') metadata.inspectionRequests = parseOptionalNum(row[1]) ?? 0;
        else if (key === 'newncr') metadata.newNCR = parseOptionalNum(row[1]) ?? 0;
        else if (key === 'photoscount') metadata.photosCount = parseOptionalNum(row[1]) ?? 0;
        else if (key === 'photosfolder') metadata.photosFolder = val || metadata.photosFolder;
        else if (key === 'incidentdetails') metadata.incidentDetails = val;
        else if (key === 'engineerinstructions') metadata.engineerInstructions = val;
        else if (key === 'plansummary') metadata.planSummary = val;
        else if (key === 'preparedby') metadata.preparedBy = val;
        else if (key === 'sitemanager') metadata.siteManager = val;
        else if (key === 'tcjvrep') metadata.tcjvRep = val;
        else if (key === 'engineerack') metadata.engineerAck = val;
        else if (key === 'approvalcomment') metadata.approvalComment = val;
      }
      if (hasRequiredColumn) metadata.requiredFields = requiredFields;
    }

    // 2. Parse 02_Manpower
    const manpower: ManpowerRow[] = [];
    const wsMp = getSectionSheet(wb, 'manpower');
    if (wsMp) {
      const rows: unknown[][] = XLSX.utils.sheet_to_json(wsMp, { header: 1, defval: '' });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || (!row[0] && !row[1])) continue;
        const org = String(row[0] || '').trim();
        const category = String(row[1] || '').trim();
        if (org.toLowerCase().includes('tổng') || category.toLowerCase().includes('tổng')) continue;
        const nos = parseOptionalNum(row[2]);
        const hrs = parseOptionalNum(row[3]);
        const manHrs = nos !== undefined && hrs !== undefined ? Math.round(nos * hrs * 100) / 100 : parseOptionalNum(row[4]);
        const isDefaultMpLocked = ['tcjv', 'ntp-tv01'].indexOf(org.toLowerCase()) !== -1;
        manpower.push({
          id: `mp-${r}`,
          org,
          category,
          nos,
          hrs,
          manHrs,
          isLocked: parseIsLocked(row[5], isDefaultMpLocked)
        });
      }
    }

    // 3. Parse 03_Equipment
    const equipment: EquipmentRow[] = [];
    const wsEq = getSectionSheet(wb, 'equipment');
    if (wsEq) {
      const rows: unknown[][] = XLSX.utils.sheet_to_json(wsEq, { header: 1, defval: '' });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !String(row[0] || '').trim()) continue;
        equipment.push({
          id: `eq-${r}`,
          equipment: String(row[0] || '').trim(),
          qty: parseOptionalNum(row[1]),
          hrs: parseOptionalNum(row[2]),
          status: String(row[3] || 'Hoạt động').trim(),
          remark: String(row[4] || '').trim(),
          isLocked: parseIsLocked(row[5], false)
        });
      }
    }

    // 4. Parse 04_WorkQty
    const workItems: WorkQtyRow[] = [];
    const wsWq = getSectionSheet(wb, 'workqty');
    if (wsWq) {
      const rows: unknown[][] = XLSX.utils.sheet_to_json(wsWq, { header: 1, defval: '' });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || (!String(row[0] || '').trim() && !String(row[1] || '').trim())) continue;
        const actId = String(row[0] || '').trim();
        const plan = parseOptionalNum(row[3]);
        const actual = parseOptionalNum(row[4]);
        const percent =
          plan !== undefined && plan > 0 && actual !== undefined
            ? Math.round((actual / plan) * 100)
            : parseOptionalNum(row[5]);
        const isDefaultWqLocked = /-dw-1120$|-dw-1130$/i.test(actId);
        workItems.push({
          id: `wq-${r}`,
          activityId: actId,
          item: String(row[1] || '').trim(),
          unit: String(row[2] || '').trim(),
          plan,
          actual,
          percent,
          remark: String(row[6] || '').trim(),
          isLocked: parseIsLocked(row[7], isDefaultWqLocked)
        });
      }
    }

    // 5. Parse 05_Materials
    const materials: MaterialRow[] = [];
    const wsMat = getSectionSheet(wb, 'materials');
    if (wsMat) {
      const rows: unknown[][] = XLSX.utils.sheet_to_json(wsMat, { header: 1, defval: '' });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !String(row[0] || '').trim()) continue;
        materials.push({
          id: `mat-${r}`,
          material: String(row[0] || '').trim(),
          unit: String(row[1] || '').trim(),
          recv: parseOptionalNum(row[2]),
          used: parseOptionalNum(row[3]),
          mirMar: String(row[4] || '').trim(),
          isLocked: parseIsLocked(row[5], false)
        });
      }
    }

    // 6. Parse 06_Issues
    const issues: IssueRow[] = [];
    const wsIss = getSectionSheet(wb, 'issues');
    if (wsIss) {
      const rows: unknown[][] = XLSX.utils.sheet_to_json(wsIss, { header: 1, defval: '' });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || (!String(row[0] || '').trim() && !String(row[1] || '').trim())) continue;
        issues.push({
          id: `iss-${r}`,
          issueId: String(row[0] || '').trim(),
          description: String(row[1] || '').trim(),
          severity: String(row[2] || 'Trung bình').trim(),
          criticalPath: String(row[3] || '–').trim(),
          owner: String(row[4] || '').trim(),
          dueDate: String(row[5] || '').trim(),
          status: String(row[6] || 'Mở').trim(),
          isLocked: parseIsLocked(row[7], false)
        });
      }
    }

    const doc: ReportDoc = {
      docId,
      unit,
      dateStr,
      metadata,
      manpower,
      equipment,
      workItems,
      materials,
      issues,
      tasks: [],
      sections: TABLE_SECTIONS.filter((k) => findSectionSheet(wb.SheetNames, k) !== undefined)
    };
    return doc;
  }

  // Fallback: Legacy 1-sheet template (e.g. WeeklyReport or old DailyReport_template.xlsx on SharePoint)
  const cleanUnit = unit.trim().toLowerCase();
  const sheetName =
    wb.SheetNames.find((s) => s.trim().toLowerCase() === cleanUnit) ||
    wb.SheetNames[0];

  if (!sheetName) {
    throw new Error('File Excel không có sheet dữ liệu hợp lệ');
  }

  const ws = wb.Sheets[sheetName];
  const rawRows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  let headerRowIdx = -1;
  const maxHeaderSearch = Math.min(rawRows.length, 15);

  for (let r = 0; r < maxHeaderSearch; r++) {
    const row = rawRows[r];
    if (!row || row.length === 0) continue;

    const col0 = String(row[0] || '').trim();
    const col0Lower = col0.toLowerCase();
    const col1 = String(row[1] || '').trim();

    if (col0Lower === 'id' || col0Lower === 'stt' || String(row[1] || '').toLowerCase().includes('công việc')) {
      headerRowIdx = r;
      break;
    }

    if (col0Lower === 'doctile' || col0Lower === 'doctitle') {
      if (col1) metadata.docTitle = col1;
    } else if (col0Lower === 'subtile' || col0Lower === 'subtitle') {
      if (col1) metadata.subTitle = col1;
    } else if (col0Lower === 'doccode') {
      if (col1) metadata.docCode = col1;
    } else if (col0Lower === 'location') {
      if (col1) metadata.location = col1;
    } else if (col0Lower === 'reportno') {
      if (col1) metadata.reportNo = col1;
    } else if (col0Lower === 'shift') {
      if (col1) metadata.shift = col1;
    } else if (col0Lower === 'email') {
      if (col1) metadata.email = col1;
    } else if (col0Lower === 'updatetime') {
      if (col1) metadata.updateTime = col1;
    }
  }

  const tasks: ReportTaskRow[] = [];
  if (headerRowIdx !== -1) {
    const headerRow = rawRows[headerRowIdx];
    let idIdx = 0;
    let nameIdx = 1;
    let rateIdx = 2;
    let qtyIdx = 3;
    let unitIdx = 4;

    headerRow.forEach((colName: unknown, idx: number) => {
      const c = String(colName || '').trim().toLowerCase();
      if (c === 'id' || c === 'stt') idIdx = idx;
      else if (c.includes('công việc') || c.includes('tên việc') || c === 'name') nameIdx = idx;
      else if (c.includes('tỉ lệ') || c.includes('ti le') || c.includes('rate')) rateIdx = idx;
      else if (c.includes('khối lượng') || c.includes('khoi luong') || c.includes('qty')) qtyIdx = idx;
      else if (c.includes('đơn vị') || c.includes('don vi') || c === 'unit') unitIdx = idx;
    });

    for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row || row.length === 0) continue;

      const rawId = row[idIdx];
      if (rawId === undefined || rawId === null || String(rawId).trim() === '') continue;

      const id = String(rawId).trim();
      const taskName = String(row[nameIdx] || '').trim();
      if (!taskName) continue;

      const completionRate = parseOptionalNum(row[rateIdx]);
      const quantity = parseOptionalNum(row[qtyIdx]);
      const taskUnit = String(row[unitIdx] || '').trim();
      const level = id.split('.').length;

      tasks.push({
        id,
        taskName,
        completionRate,
        quantity,
        unit: taskUnit,
        level,
        isLeaf: true
      });
    }
  }

  const isWeekly = docId.toLowerCase().includes('week');
  const recalculatedTasks = isWeekly ? recalculateAllRates(tasks) : [];

  const doc: ReportDoc = {
    docId,
    unit,
    dateStr,
    metadata,
    manpower: defaults.manpower,
    equipment: defaults.equipment,
    workItems: defaults.workItems,
    materials: defaults.materials,
    issues: defaults.issues,
    tasks: recalculatedTasks
  };
  return doc;
}

/**
 * Encode ReportDoc into an Excel ArrayBuffer: 01_General + only the table sheets the template has
 * (Manpower, Equipment, WorkQty, Materials, Issues), numbered consecutively.
 */
export async function encodeReportDocToBuffer(
  doc: ReportDoc,
  _templateBuffer?: ArrayBuffer
): Promise<ArrayBuffer> {
  const XLSX = await getXlsx();
  const wb = XLSX.utils.book_new();

  const m = doc.metadata;

  // 1. Sheet 01_General
  const requiredNow = getRequiredFields(m);
  const req = (key: string): string => {
    const entry = REQUIRABLE_FIELDS[key.toLowerCase()];
    return entry ? (requiredNow.indexOf(entry.field) !== -1 ? 'TRUE' : 'FALSE') : '';
  };
  const withReq = (row: (string | number)[]): (string | number)[] => [...row, req(String(row[0]))];
  const generalRowsRaw: (string | number)[][] = [
    ['Key', 'Value', 'Mô tả'],
    ['DocTile', m.docTitle || 'BÁO CÁO NGÀY', 'Tiêu đề chính (VI)'],
    ['DocTileEn', m.docTitleEn || 'DAILY REPORT', 'Tiêu đề chính (EN)'],
    ['SubTile', m.subTitle || '', 'Tên dự án / gói thầu'],
    ['DocCode', m.docCode || '', 'Mã tài liệu'],
    ['TemplateVersion', m.templateVersion || '', 'Phiên bản biểu mẫu (Template Version)'],
    ['Location', m.location || doc.unit, 'Vị trí / Trạm'],
    ['ReportNo', m.reportNo || '', 'Số BCN'],
    ['ReportDate', doc.dateStr || '', 'Ngày báo cáo'],
    ['Shift', m.shift || '', 'Ca làm việc'],
    ['WeatherAM', m.weatherAM || '', 'Thời tiết Sáng (AM)'],
    ['WeatherPM', m.weatherPM || '', 'Thời tiết Chiều (PM)'],
    ['MaxTemp', m.maxTemp || '', 'Nhiệt độ max'],
    ['HrsLost', m.hrsLost ?? 0, 'Giờ mất do thời tiết'],
    ['ProgressEval', m.progressEval || 'Đạt', 'Đánh giá tiến độ'],
    ['Summary', m.summary || '', 'I. Tóm tắt công việc & giải trình'],
    ['ManpowerSummary', m.manpowerSummary || '', 'II. Tóm tắt nhân lực'],
    ['EquipmentSummary', m.equipmentSummary || '', 'III. Tóm tắt thiết bị'],
    ['WorkQtySummary', m.workQtySummary || '', 'IV. Tóm tắt khối lượng'],
    ['MaterialsSummary', m.materialsSummary || '', 'V. Tóm tắt vật tư'],
    ['WorksPlanSummary', m.worksPlanSummary || '', 'VI. Tóm tắt công việc & kế hoạch'],
    ['IssuesSummary', m.issuesSummary || '', 'VII. Tóm tắt vướng mắc'],
    ['WorksToday', m.worksToday || '', 'VI. Công việc thực hiện hôm nay'],
    ['PlanTomorrow', m.planTomorrow || '', 'VI. Kế hoạch ngày mai'],
    ['HseIncidents', m.hseIncidents ?? 0, 'VIII. Sự cố HSE'],
    ['HseInspections', m.hseInspections ?? 0, 'VIII. Kiểm tra ATLĐ'],
    ['InspectionRequests', m.inspectionRequests ?? 0, 'VIII. Nghiệm thu (IR)'],
    ['NewNCR', m.newNCR ?? 0, 'VIII. NCR mới'],
    ['PhotosCount', m.photosCount ?? 0, 'VIII. Số ảnh hiện trường'],
    ['IncidentDetails', m.incidentDetails || 'Không có | None', 'VIII. Mô tả sự cố'],
    ['EngineerInstructions', m.engineerInstructions || '', 'VIII. Chỉ thị/ý kiến TVGS'],
    ['PlanSummary', m.planSummary || '', 'VIII. Kế hoạch ngày mai (tóm tắt)'],
    ['PreparedBy', m.preparedBy || m.email || '', 'Người lập'],
    ['Email', m.email || '', 'Email người lập'],
    ['UpdateTime', m.updateTime || new Date().toISOString(), 'Thời gian nộp'],
    ['SiteManager', m.siteManager || '', 'Trưởng ga / CHT'],
    ['TcjvRep', m.tcjvRep || '', 'Đại diện TCJV'],
    ['EngineerAck', m.engineerAck || '', 'TVGS xác nhận'],
    ['ApprovalComment', m.approvalComment || '', 'Ý kiến duyệt']
  ];
  const generalRows = generalRowsRaw.map((row, i) =>
    i === 0 ? [...row, 'Required'] : withReq(row)
  );
  const wsGeneral = XLSX.utils.aoa_to_sheet(generalRows);
  wsGeneral['!cols'] = [{ wch: 22 }, { wch: 70 }, { wch: 35 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, wsGeneral, '01_General');

  // Chỉ ghi các bảng template có; số thứ tự sheet liên tục (02_, 03_, …)
  let sheetNo = 1;
  const appendSection = (key: ReportSectionKey, ws: XLSXTypes.WorkSheet): void => {
    if (!hasSection(doc, key)) return;
    sheetNo += 1;
    XLSX.utils.book_append_sheet(wb, ws, `${('0' + sheetNo).slice(-2)}_${SECTION_SHEET_NAME[key]}`);
  };

  // 2. Sheet Manpower
  const mpRows: (string | number)[][] = [['Org', 'Category', 'Nos', 'Hrs', 'ManHrs', 'IsLocked']];
  doc.manpower.forEach((r) => {
    const nos = r.nos ?? 0;
    const hrs = r.hrs ?? 0;
    const manHrs = Math.round(nos * hrs * 100) / 100;
    mpRows.push([r.org || '', r.category || '', nos, hrs, manHrs, r.isLocked ? 'TRUE' : 'FALSE']);
  });
  const wsMp = XLSX.utils.aoa_to_sheet(mpRows);
  wsMp['!cols'] = [{ wch: 16 }, { wch: 28 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }];
  appendSection('manpower', wsMp);

  // 3. Sheet 03_Equipment
  const eqRows: (string | number)[][] = [['Equipment', 'Qty', 'Hrs', 'Status', 'Remark', 'IsLocked']];
  doc.equipment.forEach((r) => {
    eqRows.push([
      r.equipment || '',
      r.qty ?? 0,
      r.hrs ?? 0,
      r.status || 'Hoạt động',
      r.remark || '',
      r.isLocked ? 'TRUE' : 'FALSE'
    ]);
  });
  const wsEq = XLSX.utils.aoa_to_sheet(eqRows);
  wsEq['!cols'] = [{ wch: 32 }, { wch: 10 }, { wch: 12 }, { wch: 16 }, { wch: 28 }, { wch: 12 }];
  appendSection('equipment', wsEq);

  // 4. Sheet 04_WorkQty
  const wqRows: (string | number)[][] = [['ActivityID', 'Item', 'Unit', 'Plan', 'Actual', 'Percent', 'Remark', 'IsLocked']];
  doc.workItems.forEach((r) => {
    const plan = r.plan ?? 0;
    const actual = r.actual ?? 0;
    const pct = plan > 0 ? Math.round((actual / plan) * 100) : r.percent ?? 0;
    wqRows.push([
      r.activityId || '',
      r.item || '',
      r.unit || '',
      plan,
      actual,
      pct,
      r.remark || '',
      r.isLocked ? 'TRUE' : 'FALSE'
    ]);
  });
  const wsWq = XLSX.utils.aoa_to_sheet(wqRows);
  wsWq['!cols'] = [{ wch: 18 }, { wch: 34 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 24 }, { wch: 12 }];
  appendSection('workqty', wsWq);

  // 5. Sheet 05_Materials
  const matRows: (string | number)[][] = [['Material', 'Unit', 'Recv', 'Used', 'MirMar', 'IsLocked']];
  doc.materials.forEach((r) => {
    matRows.push([
      r.material || '',
      r.unit || '',
      r.recv ?? 0,
      r.used ?? 0,
      r.mirMar || '',
      r.isLocked ? 'TRUE' : 'FALSE'
    ]);
  });
  const wsMat = XLSX.utils.aoa_to_sheet(matRows);
  wsMat['!cols'] = [{ wch: 30 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 28 }, { wch: 12 }];
  appendSection('materials', wsMat);

  // 6. Sheet 06_Issues
  const issRows: (string | number)[][] = [
    ['IssueID', 'Description', 'Severity', 'CriticalPath', 'Owner', 'DueDate', 'Status', 'IsLocked']
  ];
  doc.issues.forEach((r) => {
    issRows.push([
      r.issueId || '',
      r.description || '',
      r.severity || 'Trung bình',
      r.criticalPath || '–',
      r.owner || '',
      r.dueDate || '',
      r.status || 'Mở',
      r.isLocked ? 'TRUE' : 'FALSE'
    ]);
  });
  const wsIss = XLSX.utils.aoa_to_sheet(issRows);
  wsIss['!cols'] = [{ wch: 16 }, { wch: 50 }, { wch: 14 }, { wch: 12 }, { wch: 24 }, { wch: 12 }, { wch: 14 }, { wch: 12 }];
  appendSection('issues', wsIss);

  // Optional: If tasks exist (WeeklyReport), also include a 07_Tasks sheet
  if (doc.tasks && doc.tasks.length > 0) {
    sheetNo += 1;
    const taskRows: (string | number)[][] = [['ID', 'Tên công việc', 'Tỉ lệ hoàn thành', 'Khối lượng', 'Đơn vị']];
    doc.tasks.forEach((t) => {
      taskRows.push([t.id, t.taskName, t.completionRate ?? '', t.quantity ?? '', t.unit || '']);
    });
    const wsTasks = XLSX.utils.aoa_to_sheet(taskRows);
    wsTasks['!cols'] = [{ wch: 10 }, { wch: 35 }, { wch: 18 }, { wch: 14 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, wsTasks, `${('0' + sheetNo).slice(-2)}_Tasks`);
  }

  const array = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return array as ArrayBuffer;
}
