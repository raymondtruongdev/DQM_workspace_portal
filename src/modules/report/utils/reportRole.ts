import { extractCleanEmail } from '../../../core/auth';
import { RoleEntry } from '../../../core/types';

/** Tên nội bộ (SharePoint) của các cột trong list Report_Role; tra từ tên hiển thị bằng getListFieldMap. */
export interface RoleListNames {
  displayName: string;
  email: string;
  unit: string;
  docId: string;
  docTitle: string;
  active: string;
}

const clean = (v: unknown): string =>
  v === undefined || v === null ? '' : String(v).replace(/[​-‍﻿ ]/g, ' ').trim();

/** "drp", "BCN", "daily report"… -> DailyReport; "wrp", "bct"… -> WeeklyReport; giá trị khác giữ nguyên. */
export function normalizeDocId(raw: string): string {
  const lower = raw.trim().toLowerCase().replace(/[\s_-]+/g, '');
  if (lower === 'dailyreport' || lower === 'drp' || lower === 'daily' || lower === 'bcn') return 'DailyReport';
  if (lower === 'weeklyreport' || lower === 'wrp' || lower === 'weekly' || lower === 'bct') return 'WeeklyReport';
  return raw.trim();
}

/** Cột Active: Yes/No (true/false) hoặc chữ TRUE/YES/1/Active…; để trống = đang hiệu lực. */
export function parseActive(raw: unknown): boolean {
  if (raw === undefined || raw === null || clean(raw) === '') return true;
  if (typeof raw === 'boolean') return raw;
  if (typeof raw === 'number') return raw !== 0;
  return ['true', '1', 'yes', 'y', 'active', 'x', 'hoạt động'].indexOf(clean(raw).toLowerCase()) !== -1;
}

/**
 * Các dòng của list Report_Role thành RoleEntry: mỗi dòng = một người × một đơn vị × một loại báo cáo.
 * Một ô Email có thể chứa nhiều email ngăn bằng `;` `,` hoặc xuống dòng (mỗi email thành một phân quyền).
 * Dòng thiếu Unit, DocID hoặc Email bị bỏ qua. `DisplayName` là tên người phụ trách (tự điền vào ô "Trưởng ga / CHT").
 */
export function parseRoleRows(rows: Array<Record<string, unknown>>, n: RoleListNames): RoleEntry[] {
  const out: RoleEntry[] = [];
  rows.forEach((row, i) => {
    const unit = clean(row[n.unit]);
    const docIdRaw = clean(row[n.docId]);
    const emailCell = clean(row[n.email]);
    if (!unit || !docIdRaw || !emailCell) return;
    const emails = emailCell
      .split(/[;,\n]+/)
      .map((part) => extractCleanEmail(part))
      .filter(Boolean);
    const base = row.Id !== undefined ? String(row.Id) : String(i + 1);
    const docTitle = clean(row[n.docTitle]) || undefined;
    const siteManager = clean(row[n.displayName]);
    const status = parseActive(row[n.active]);
    emails.forEach((email, k) => {
      out.push({
        id: emails.length > 1 ? `${base}_${k + 1}` : base,
        docId: normalizeDocId(docIdRaw),
        docTitle,
        unit,
        siteManager,
        email,
        status
      });
    });
  });
  return out;
}
