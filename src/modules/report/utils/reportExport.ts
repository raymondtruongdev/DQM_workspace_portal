import type * as XLSXTypes from 'xlsx';
import { DAY_STATUS_LABEL, formatDMY, weekdayLong } from './reportCalendar';
import { ReportListItem } from './reportOverview';

export type ExportCell = string | number;

const HEADER: ExportCell[] = [
  'Ngày',
  'Thứ',
  'Đơn vị',
  'Loại báo cáo',
  'Trạng thái',
  'Số báo cáo',
  'Người lập',
  'Lưu lúc',
  'File đã ký',
  'Upload bản ký lúc',
  'Người phụ trách',
  'Email'
];

function pad(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function formatDateTime(iso?: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Bảng xuất (dòng đầu là tiêu đề) từ danh sách báo cáo đang hiển thị. */
export function buildExportRows(items: ReportListItem[]): ExportCell[][] {
  const rows: ExportCell[][] = [HEADER];
  items.forEach((i) => {
    rows.push([
      formatDMY(i.dateStr),
      weekdayLong(i.dateStr),
      i.target.unit,
      i.target.docId,
      DAY_STATUS_LABEL[i.status],
      i.info && i.info.reportNo ? i.info.reportNo : '',
      i.info && i.info.reporter ? i.info.reporter : '',
      formatDateTime(i.info ? i.info.savedAt : undefined),
      i.info && i.info.signed ? (i.info.signed.name || 'Có') : '',
      i.info && i.info.signed ? formatDateTime(i.info.signed.timeLastModified) : '',
      i.target.owners.map((o) => o.name || o.email).join('; '),
      i.target.owners.map((o) => o.email).filter(Boolean).join('; ')
    ]);
  });
  return rows;
}

function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Xuất .xlsx (thư viện SheetJS chỉ tải khi bấm xuất). */
export async function exportXlsx(rows: ExportCell[][], fileName: string): Promise<void> {
  const XLSX: typeof XLSXTypes = await import(/* webpackChunkName: 'xlsx' */ 'xlsx');
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 20 }, { wch: 12 }, { wch: 30 }, { wch: 18 }, { wch: 34 }, { wch: 18 }, { wch: 40 }, { wch: 44 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Bao cao');
  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
  downloadBlob(
    new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    fileName
  );
}

/** Tên file xuất, ví dụ BaoCao_ST01_20261112.xlsx (đơn vị có khi đang lọc một đơn vị). */
export function exportFileName(ext: 'xlsx', unit?: string, today: Date = new Date()): string {
  const d = `${today.getFullYear()}${pad(today.getMonth() + 1)}${pad(today.getDate())}`;
  return `BaoCao${unit && unit !== 'ALL' ? `_${unit}` : ''}_${d}.${ext}`;
}
