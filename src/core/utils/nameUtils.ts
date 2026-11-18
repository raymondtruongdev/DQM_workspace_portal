/**
 * Chữ viết tắt của tên để hiện trong avatar: chữ cái đầu của mỗi từ, tối đa 5 chữ.
 * "Trương Quốc Toàn" → "TQT", "Trương Đình" → "TĐ"; tên một từ thì lấy 2 chữ đầu ("admin" → "AD").
 */
export function getInitials(name: string): string {
  const words = (name || '').trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return '?';
  const letters = words.length === 1 ? Array.from(words[0]).slice(0, 2) : words.slice(0, 5).map((w) => Array.from(w)[0]);
  return letters.join('').toUpperCase();
}
