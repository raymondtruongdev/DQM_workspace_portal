/**
 * General path and date utility functions for SharePoint document libraries.
 */

export function sanitizeSharePointPath(segment: string): string {
  // Remove chars forbidden by SharePoint
  return segment.replace(/["*:<>?/\\|#%~&{}]/g, '-').trim();
}

export function padZero(n: number): string {
  return n < 10 ? `0${n}` : `${n}`;
}

/**
 * Format a Date object to YYYY-MM-DD
 */
export function formatDateToYYYYMMDD(d: Date): string {
  const yyyy = d.getFullYear();
  const mm = padZero(d.getMonth() + 1);
  const dd = padZero(d.getDate());
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Format a Date object to yyyyMMdd (e.g. 20261110)
 */
export function formatDateToYYYYMMDDWithoutDash(d: Date): string {
  const yyyy = d.getFullYear();
  const mm = padZero(d.getMonth() + 1);
  const dd = padZero(d.getDate());
  return `${yyyy}${mm}${dd}`;
}

/**
 * Parse a YYYY-MM-DD or yyyyMMdd string into a Date object
 */
export function parseDateString(str: string): Date {
  const clean = str.trim();
  if (clean.includes('-')) {
    const parts = clean.split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  if (clean.length === 8 && /^\d{8}$/.test(clean)) {
    const yyyy = parseInt(clean.substring(0, 4), 10);
    const mm = parseInt(clean.substring(4, 6), 10);
    const dd = parseInt(clean.substring(6, 8), 10);
    return new Date(yyyy, mm - 1, dd);
  }
  return new Date(clean);
}

