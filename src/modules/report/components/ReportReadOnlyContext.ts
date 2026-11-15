import * as React from 'react';

/**
 * true = báo cáo đang khoá: người dùng chỉ xem và sao chép chữ. Ô nhập chữ/số/ngày dùng `readOnly` (chọn và copy được),
 * còn ô chọn, nút thêm/xoá dòng và tải ảnh bị vô hiệu hoá bằng `disabled` như thường.
 */
export const ReportReadOnlyContext = React.createContext<boolean>(false);
