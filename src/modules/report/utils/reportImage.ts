/** Ảnh nhúng vào file HTML độc lập: đổi sang data URI, ảnh lớn được thu nhỏ để file không quá nặng. */

const MIME_BY_EXT: Record<string, string> = {
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  gif: 'image/gif',
  webp: 'image/webp',
  bmp: 'image/bmp'
};

export function mimeOfFile(name: string): string {
  const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
  return MIME_BY_EXT[ext] || 'application/octet-stream';
}

export function bufferToDataUrl(buffer: ArrayBuffer, mime: string): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes, i, i + chunk) as number[]);
  }
  return `data:${mime};base64,${window.btoa(binary)}`;
}

/** Cạnh dài tối đa (px) của ảnh nhúng: bản in A4 chỉ rộng khoảng 8 cm mỗi ảnh nên đủ nét. */
export const EMBED_MAX_SIDE = 1280;

/**
 * Data URI của ảnh để nhúng vào HTML. Ảnh có cạnh dài hơn `maxSide` được thu nhỏ (JPEG, chất lượng `quality`);
 * ảnh nhỏ hoặc không thu nhỏ được thì giữ nguyên nội dung gốc.
 */
export async function embeddableImage(
  buffer: ArrayBuffer,
  name: string,
  maxSide: number = EMBED_MAX_SIDE,
  quality = 0.82
): Promise<string> {
  const mime = mimeOfFile(name);
  try {
    if (typeof createImageBitmap !== 'function') return bufferToDataUrl(buffer, mime);
    const bitmap = await createImageBitmap(new Blob([buffer], { type: mime }));
    const scale = Math.min(1, maxSide / Math.max(bitmap.width, bitmap.height));
    if (scale >= 1) {
      bitmap.close();
      return bufferToDataUrl(buffer, mime);
    }
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(bitmap.width * scale);
    canvas.height = Math.round(bitmap.height * scale);
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      bitmap.close();
      return bufferToDataUrl(buffer, mime);
    }
    ctx.fillStyle = 'white';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    bitmap.close();
    return canvas.toDataURL('image/jpeg', quality);
  } catch {
    return bufferToDataUrl(buffer, mime);
  }
}
