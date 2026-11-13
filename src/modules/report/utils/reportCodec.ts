import type * as XLSXTypes from 'xlsx';
import {
  ReportDoc,
  ReportMetadata,
  FormField,
  PhotoGroup,
  ReportForm,
  ReportPhotoSection,
  ReportSection,
  ReportTable,
  TableColumn,
  TableRow,
  ReportTaskRow,
  RequirableField,
  CellValue
} from '../types/report.types';
import { recalculateAllRates } from './reportTaskTree';
import { parseFieldSpec } from './reportForm';
import { syncMetadata } from './reportMeta';
import {
  LOCK_COLUMN_CODE,
  excelSerialToVN,
  isEmptyRow,
  parseColumnSpec,
  recalcTable
} from './reportTable';
import { REQUIRABLE_FIELDS, getRequiredFields } from './reportValidation';

/** Bỏ tiền tố số của tên sheet: "05_Issues" / "5.Issues" -> "Issues". */
function sheetBase(sheetName: string): string {
  return sheetName.trim().replace(/^\d+\s*[_.\-\s]\s*/, '');
}

function isGeneralSheet(sheetName: string): boolean {
  return sheetBase(sheetName).toLowerCase() === 'general';
}

const TOTAL_LABEL_RE = /^(tổng(\s+cộng)?|total)\s*(\|.*)?$/i;

/** Đọc một sheet bảng (dòng 1 bắt đầu bằng `Title`). */
function parseTableSheet(
  XLSX: typeof XLSXTypes,
  ws: XLSXTypes.WorkSheet,
  sheetName: string
): ReportTable | undefined {
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length === 0) return undefined;
  const base = sheetBase(sheetName);
  const key = base.toLowerCase();

  const first = String((rows[0] && rows[0][0]) ?? '').trim().toLowerCase();
  if (first !== 'title') return undefined;

  const columns: TableColumn[] = [];
  const colIdx: number[] = [];
  let lockIdx = -1;
  const title = String(rows[0][1] ?? '').trim();
  const titleEn = String(rows[0][2] ?? '').trim();
  const shortTitle = String(rows[0][3] ?? '').trim();
  const summary =
    rows[1] && String(rows[1][0] ?? '').trim().toLowerCase() === 'summary' ? String(rows[1][1] ?? '').trim() : '';
  const showSummary = !(rows[1] && String(rows[1][2] ?? '').trim().toLowerCase() === 'off');
  const single = String(rows[0][4] ?? '').trim().toLowerCase() === 'single';
  const codes = rows[2] || [];
  const labels = rows[3] || [];
  const specs = rows[4] || [];
  codes.forEach((c, i) => {
    const code = String(c ?? '').trim();
    if (!code) return;
    if (code.toLowerCase() === LOCK_COLUMN_CODE.toLowerCase()) {
      lockIdx = i;
      return;
    }
    columns.push(parseColumnSpec(code, String(labels[i] ?? '').trim(), String(specs[i] ?? '')));
    colIdx.push(i);
  });
  const dataStart = 5;
  if (columns.length === 0) return undefined;

  const tableRows: TableRow[] = [];
  for (let r = dataStart; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const values: Record<string, CellValue> = {};
    columns.forEach((c, ci) => {
      const raw = row[colIdx[ci]];
      if (c.type === 'number') values[c.code] = parseOptionalNum(raw);
      else if (c.type === 'calc') values[c.code] = 0;
      else if (c.type === 'date' && typeof raw === 'number') values[c.code] = excelSerialToVN(raw);
      else values[c.code] = String(raw ?? '').trim();
    });
    if (isEmptyRow(columns, values)) continue;
    const firstText = columns.length > 0 ? values[columns[0].code] : '';
    if (typeof firstText === 'string' && TOTAL_LABEL_RE.test(firstText)) continue;
    columns.forEach((c) => {
      if (c.type === 'select' && values[c.code] === '') {
        values[c.code] = c.defaultValue !== undefined ? c.defaultValue : c.options && c.options.length ? c.options[0] : '';
      }
    });
    const lockRaw = lockIdx >= 0 ? row[lockIdx] : undefined;
    tableRows.push({
      id: `${key}-${r}`,
      isLocked: parseIsLocked(lockRaw, false),
      values
    });
  }

  return recalcTable({ kind: 'table', key, sheetBase: base, title, titleEn, shortTitle, summary, showSummary, single, columns, rows: tableRows });
}

/** Đọc sheet form: Title | Summary | Layout | (Key Label Spec Value Required) | các trường. */
function parseFormSheet(
  XLSX: typeof XLSXTypes,
  ws: XLSXTypes.WorkSheet,
  sheetName: string
): ReportForm | undefined {
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const cell = (r: number, c: number): string => String((rows[r] && rows[r][c]) ?? '').trim();
  if (rows.length < 4) return undefined;
  if (cell(0, 0).toLowerCase() !== 'title') return undefined;
  if (cell(2, 0).toLowerCase() !== 'layout' || cell(3, 0).toLowerCase() !== 'key') return undefined;
  const base = sheetBase(sheetName);
  const cols = parseInt(cell(2, 1), 10);
  const fields: FormField[] = [];
  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    const key = cell(r, 0);
    if (!key) continue;
    const probe = parseFieldSpec(key, cell(r, 1), String(row[2] ?? ''), undefined, false);
    const raw = row[3];
    let value: CellValue;
    if (probe.type === 'number') value = parseOptionalNum(raw);
    else value = String(raw ?? '').trim();
    const required = row[4] !== undefined && String(row[4]).trim() !== '' && parseIsLocked(row[4], false);
    fields.push({ ...probe, value, required });
  }
  return {
    kind: 'form',
    key: base.toLowerCase(),
    sheetBase: base,
    title: cell(0, 1),
    titleEn: cell(0, 2),
    shortTitle: cell(0, 3),
    columns: isNaN(cols) || cols < 1 ? 2 : cols,
    sign: cell(2, 2).toLowerCase() === 'sign',
    summary: cell(1, 1),
    showSummary: cell(1, 2).toLowerCase() !== 'off',
    fields
  };
}

/** Đọc sheet ảnh: Title | Summary | Layout | (GroupID Department Summary Files) | các nhóm ảnh. */
function parsePhotoSheet(
  XLSX: typeof XLSXTypes,
  ws: XLSXTypes.WorkSheet,
  sheetName: string
): ReportPhotoSection | undefined {
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const cell = (r: number, c: number): string => String((rows[r] && rows[r][c]) ?? '').trim();
  if (rows.length < 4) return undefined;
  if (cell(0, 0).toLowerCase() !== 'title') return undefined;
  if (cell(2, 0).toLowerCase() !== 'layout' || cell(3, 0).toLowerCase() !== 'groupid') return undefined;
  const base = sheetBase(sheetName);
  const cols = parseInt(cell(2, 1), 10);
  const groups: PhotoGroup[] = [];
  for (let r = 4; r < rows.length; r++) {
    const id = cell(r, 0);
    if (!id) continue;
    groups.push({
      id,
      department: cell(r, 1),
      summary: cell(r, 2),
      files: cell(r, 3).split('|').map((f) => f.trim()).filter(Boolean)
    });
  }
  return {
    kind: 'photo',
    key: base.toLowerCase(),
    sheetBase: base,
    title: cell(0, 1),
    titleEn: cell(0, 2),
    shortTitle: cell(0, 3),
    summary: cell(1, 1),
    showSummary: cell(1, 2).toLowerCase() !== 'off',
    columns: isNaN(cols) || cols < 1 ? 2 : cols,
    groups
  };
}

function parseSectionSheet(
  XLSX: typeof XLSXTypes,
  ws: XLSXTypes.WorkSheet,
  sheetName: string
): ReportSection | undefined {
  return (
    parseFormSheet(XLSX, ws, sheetName) ||
    parsePhotoSheet(XLSX, ws, sheetName) ||
    parseTableSheet(XLSX, ws, sheetName)
  );
}

async function getXlsx(): Promise<typeof XLSXTypes> {
  return await import(/* webpackChunkName: 'xlsx' */ 'xlsx');
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
 * Parse an Excel buffer (multi-sheet template/report, or the 1-sheet WeeklyReport placeholder) into ReportDoc
 */
export async function parseReportDocFromBuffer(
  buffer: ArrayBuffer,
  docId: string,
  unit: string,
  dateStr: string
): Promise<ReportDoc> {
  const XLSX = await getXlsx();
  const wb = XLSX.read(buffer, { type: 'array' });


  const metadata: ReportMetadata = {
    docTitle: docId === 'WeeklyReport' ? 'BÁO CÁO TUẦN' : 'BÁO CÁO NGÀY',
    docTitleEn: docId === 'WeeklyReport' ? 'WEEKLY REPORT' : 'DAILY REPORT',
    subTitle: '',
    docCode: '',
    templateVersion: '',
    location: unit,
    reportNo: '',
    subcontractor: '',
    email: '',
    updateTime: ''
  };

  const hasMultiSheet = wb.SheetNames.some((n) => isGeneralSheet(n));

  if (hasMultiSheet) {
    // 1. Parse 01_General
    const generalName = wb.SheetNames.find((n) => isGeneralSheet(n));
    const wsGen = generalName ? wb.Sheets[generalName] : wb.Sheets[wb.SheetNames[0]];
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
        const reqFlag = row[3] !== undefined && String(row[3]).trim() !== '';
        if (reqEntry && reqFlag) {
          hasRequiredColumn = true;
          if (parseIsLocked(row[3], false)) requiredFields.push(reqEntry.field);
        }

        if (key === 'doctitle') metadata.docTitle = val || metadata.docTitle;
        else if (key === 'doctitleen') metadata.docTitleEn = val || metadata.docTitleEn;
        else if (key === 'subtitle') metadata.subTitle = val || metadata.subTitle;
        else if (key === 'doccode') {
          if (/\{(unit|reportNo)\}/i.test(val)) metadata.docCodePattern = val;
          else metadata.docCode = val;
        } else if (key === 'doccodepattern') metadata.docCodePattern = val || undefined;
        else if (key === 'templateversion') metadata.templateVersion = val;
        else if (key === 'location') metadata.location = val || metadata.location;
        else if (key === 'reportno') metadata.reportNo = val;
        else if (key === 'subcontractor') metadata.subcontractor = val;
        else if (key === 'subcontractoroptions') {
          metadata.subcontractorOptions = val.split('/').map((o) => o.trim()).filter(Boolean);
        }
        else if (key === 'email') metadata.email = val;
        else if (key === 'updatetime') metadata.updateTime = val;
            }
      if (hasRequiredColumn) metadata.requiredFields = requiredFields;
    }

    // 2. Mỗi sheet còn lại = một mục: form (A1 "Title" + A3 "Key") hoặc bảng (A1 "Title"); thứ tự mục = thứ tự sheet
    const sections: ReportSection[] = [];
    wb.SheetNames.forEach((name) => {
      if (isGeneralSheet(name)) return;
      const sec = parseSectionSheet(XLSX, wb.Sheets[name], name);
      if (sec) sections.push(sec);
    });

    const doc: ReportDoc = {
      docId,
      unit,
      dateStr,
      metadata: syncMetadata(metadata, unit),
      sections,
      tasks: []
    };
    return doc;
  }

  // Báo cáo ngày chỉ hỗ trợ định dạng nhiều sheet có 01_General; file khác là sai định dạng (không còn đọc định dạng cũ)
  if (!docId.toLowerCase().includes('week')) {
    throw new Error('File báo cáo/template không đúng định dạng (thiếu sheet 01_General).');
  }

  // WeeklyReport placeholder: 1-sheet task-tree template (không có mục form/bảng)
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
    sections: [],
    tasks: recalculatedTasks
  };
  return doc;
}

/**
 * Encode ReportDoc into an Excel ArrayBuffer: 01_General + one sheet per table section
 * (self-describing: title, summary, columns, spec), numbered consecutively.
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
    ['DocTitle', m.docTitle || 'BÁO CÁO NGÀY', 'Tiêu đề chính (VI)'],
    ['DocTitleEn', m.docTitleEn || 'DAILY REPORT', 'Tiêu đề chính (EN)'],
    ['SubTitle', m.subTitle || '', 'Tên dự án / gói thầu'],
    ['DocCode', m.docCode || '', 'Mã tài liệu'],
    ['DocCodePattern', m.docCodePattern || '', 'Mẫu mã tài liệu ({unit} {reportNo})'],
    ['TemplateVersion', m.templateVersion || '', 'Phiên bản biểu mẫu (Template Version)'],
    ['Location', m.location || doc.unit, 'Vị trí / Trạm'],
    ['ReportNo', m.reportNo || '', 'Số báo cáo (Report no.)'],
    ['Subcontractor', m.subcontractor || '', 'Đơn vị thi công (Subcontractor), nhiều đơn vị nối bằng dấu phẩy'],
    ['SubcontractorOptions', (m.subcontractorOptions || []).join('/'), 'Danh sách đơn vị thi công có thể chọn, ngăn bằng /'],
    ['ReportDate', doc.dateStr || '', 'Ngày báo cáo'],
    ['Email', m.email || '', 'Email người lập'],
    ['UpdateTime', m.updateTime || new Date().toISOString(), 'Thời gian nộp']
  ];
  const generalRows = generalRowsRaw.map((row, i) =>
    i === 0 ? [...row, 'Required'] : withReq(row)
  );
  const wsGeneral = XLSX.utils.aoa_to_sheet(generalRows);
  wsGeneral['!cols'] = [{ wch: 22 }, { wch: 70 }, { wch: 35 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, wsGeneral, '01_General');

  // Mỗi mục dạng bảng = một sheet <số>_<Tên>, ghi lại đúng cấu trúc template (tiêu đề, cột, spec, tóm tắt)
  let sheetNo = 1;
  doc.sections.forEach((sec) => {
    sheetNo += 1;
    const sheetName = `${('0' + sheetNo).slice(-2)}_${sec.sheetBase}`;
    if (sec.kind === 'form') {
      const aoa: (string | number | boolean)[][] = [
        ['Title', sec.title, sec.titleEn, sec.shortTitle],
        sec.showSummary ? ['Summary', sec.summary || ''] : ['Summary', sec.summary || '', 'off'],
        sec.sign ? ['Layout', sec.columns, 'sign'] : ['Layout', sec.columns],
        ['Key', 'Label', 'Spec', 'Value', 'Required']
      ];
      sec.fields.forEach((f) => {
        aoa.push([
          f.key,
          f.label,
          f.spec,
          f.value === undefined ? '' : f.value,
          f.required ? 'TRUE' : 'FALSE'
        ]);
      });
      const wsForm = XLSX.utils.aoa_to_sheet(aoa);
      wsForm['!cols'] = [{ wch: 22 }, { wch: 40 }, { wch: 34 }, { wch: 50 }, { wch: 10 }];
      XLSX.utils.book_append_sheet(wb, wsForm, sheetName);
      return;
    }
    if (sec.kind === 'photo') {
      const aoa: (string | number)[][] = [
        ['Title', sec.title, sec.titleEn, sec.shortTitle],
        sec.showSummary ? ['Summary', sec.summary || ''] : ['Summary', sec.summary || '', 'off'],
        ['Layout', sec.columns],
        ['GroupID', 'Department', 'Summary', 'Files']
      ];
      sec.groups.forEach((g) => {
        aoa.push([g.id, g.department, g.summary, g.files.join('|')]);
      });
      const wsPhoto = XLSX.utils.aoa_to_sheet(aoa);
      wsPhoto['!cols'] = [{ wch: 22 }, { wch: 28 }, { wch: 50 }, { wch: 60 }];
      XLSX.utils.book_append_sheet(wb, wsPhoto, sheetName);
      return;
    }
    const t = sec;
    const codes: string[] = t.columns.map((c) => c.code).concat(LOCK_COLUMN_CODE);
    const aoa: (string | number)[][] = [
      t.single ? ['Title', t.title, t.titleEn, t.shortTitle, 'single'] : ['Title', t.title, t.titleEn, t.shortTitle],
      t.showSummary ? ['Summary', t.summary || ''] : ['Summary', t.summary || '', 'off'],
      codes,
      t.columns.map((c) => c.label).concat(''),
      t.columns.map((c) => c.spec).concat('')
    ];
    t.rows.forEach((r) => {
      aoa.push(
        t.columns
          .map((c): string | number => {
            const v = r.values[c.code];
            return v === undefined ? '' : v;
          })
          .concat(r.isLocked ? 'TRUE' : 'FALSE')
      );
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = t.columns
      .map((c) => ({ wch: Math.max(10, Math.round((c.width || 14) * 1.2)) }))
      .concat({ wch: 12 });
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

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
