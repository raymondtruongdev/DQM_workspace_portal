/**
 * Trả về mảng mới sau khi đổi chỗ phần tử ở `index` với phần tử liền trên (`delta = -1`) hoặc liền dưới (`delta = 1`).
 * Ngoài phạm vi thì trả lại đúng mảng cũ (không tạo mảng mới).
 */
export function moveItem<T>(items: T[], index: number, delta: -1 | 1): T[] {
  const target = index + delta;
  if (index < 0 || index >= items.length || target < 0 || target >= items.length) return items;
  const next = items.slice();
  const tmp = next[index];
  next[index] = next[target];
  next[target] = tmp;
  return next;
}
