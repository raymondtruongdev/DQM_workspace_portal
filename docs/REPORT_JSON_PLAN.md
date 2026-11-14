# Plan: lưu Daily Report (DRP) dạng JSON + HTML độc lập

> Trạng thái: **đã triển khai** (xem mục 13). Tài liệu này giữ hợp đồng JSON và lý do thiết kế; quy tắc vận hành nằm ở `AGENTS.md` mục 5.3. File mẫu: JSON [sample/DailyReport.sample.data.json](sample/DailyReport.sample.data.json) và HTML độc lập [sample/DailyReport.sample.html](sample/DailyReport.sample.html) (ảnh nhúng sẵn). Thư mục báo cáo demo đúng cấu trúc SharePoint: `data_test/sharepoint_sample/data/ST01/DRP_20261114_ST01/`.
> Hợp đồng JSON (mục 3) đã được tóm tắt vào mục Report của `AGENTS.md`.
> **Lưu ý (đã chốt sau khi triển khai):** app chỉ hỗ trợ JSON. Đường đọc dự phòng `.xlsx`, `reportCodec` và script chuyển hàng loạt xlsx → json (nhắc ở các mục 3-12 bên dưới) **đã bỏ, không làm**; template Excel cũ đã xoá khỏi repo. Các đoạn nhắc `.xlsx` bên dưới chỉ là bối cảnh lịch sử của kế hoạch.

## 1. Mục tiêu

Hiện mỗi báo cáo ngày được lưu thành một file `.xlsx`, nên dữ liệu bị buộc vào vị trí sheet/dòng/cột và cú pháp spec riêng của Excel. Mục tiêu:

- Khi bấm **Lưu** tạo ra **2 file**:
  - `*.json`: dữ liệu báo cáo (backup, đọc lại để sửa).
  - `*.html`: bản xem/in độc lập, **nhúng thẳng JSON**, mở offline và in PDF mà không cần vào web.
- Dữ liệu tổ chức theo dòng-cột, thêm section mới dễ, không phụ thuộc tên cột cố định hay vị trí cột.
- Người mở JSON đọc được số liệu thật mà không cần web UI.

## 2. Nguyên tắc thiết kế

| Nằm trong JSON | Nằm ở template / SharePoint list (UI lo) |
|---|---|
| Tiêu đề, nhãn cột, giá trị thật (kể cả ô đã tính và dòng Tổng) | Kiểu ô, độ rộng, căn lề, bố cục, ô trải hết hàng |
| `id` ổn định của section, cột, dòng | Danh sách dropdown, danh mục, công thức, mặc định, luỹ kế, ô chỉ đọc |

- Web ghép cột trong JSON với template **theo `id` cột** để lấy hành vi sửa.
- Khi mở báo cáo, web **tính lại công thức và `totals`**, không tin giá trị đã lưu.
- File HTML được sinh lúc Lưu bằng chính code UI (đã có template), nên ngày tháng và số đã được định dạng sẵn; JSON không cần biết kiểu ô.

## 3. Hợp đồng JSON

### 3.1 Cấp cao nhất
| Trường | Ý nghĩa |
|---|---|
| `schema` | Phiên bản định dạng (hiện là `1`) |
| `unit`, `docId`, `date` | Đơn vị (`ST01`), loại báo cáo (`DailyReport`), ngày `yyyy-MM-dd` |
| `reportNo` | Số báo cáo dạng số nguyên (in thành `00001`) |
| `docCode` | Mã tài liệu đã thay `{unit}`, `{reportNo}`. Mẫu `docCodePattern` thuộc template nên không lưu |
| `title` | `vi`, `en`, `sub` (tên dự án, một chuỗi, HTML tự xuống dòng) |
| `template` | Tên template đã dùng (vd `template_DRP_ST01_20261113`), để thông báo cập nhật phiên bản |
| `created`, `updated` | `{ by, at }`. `created` ghi một lần ở lần Lưu đầu, `updated` ghi mỗi lần Lưu. `by` do trình duyệt ghi nên chỉ để hiển thị |
| `subcontractors` | Mảng chuỗi các đơn vị thi công **được chọn** trong báo cáo này. Danh sách để chọn lấy từ SharePoint list, không lưu trong JSON |
| `sections` | Mảng các mục; **thứ tự trong mảng = thứ tự hiển thị**, HTML tự đánh số I, II, III |

### 3.2 Section
| Trường | Ý nghĩa |
|---|---|
| `id` | Định danh ổn định, không đổi (`weather`, `workQty`, …) |
| `title` | `{ vi, en, short }`. `short` chỉ cho thanh nhảy nhanh trên web |
| `summary` | **Luôn có**, trống thì `""` |
| `columns` + `rows` | Phần bảng (dữ liệu lặp theo dòng). Tuỳ chọn |
| `totals` | `{ <columnId>: số }`; có thì in dòng Tổng. Tuỳ chọn |
| `fields` | Phần ô văn bản đơn lẻ của cả section `[{ id, label, value }]` (ghi chú HSE, ký duyệt, chỉ đạo). Tuỳ chọn |
| `type` | Chỉ có giá trị `"photo"` cho mục ảnh (dòng ảnh có hình dạng khác). Các mục khác không ghi |

### 3.3 Cột, dòng, field, ảnh
- **Cột:** chỉ `{ id, label: { vi, en } }`.
- **Dòng bảng:** `{ id, locked?, cells }`. `locked: true` chỉ ghi khi đúng (dòng mẫu không được xoá). `cells` khoá theo `column.id`, chứa **mọi ô thật kể cả ô đã tính** (`progress`, `cumTotal`).
- **`id` dòng:** `r1`, `r2`…; dòng mới = số lớn nhất hiện có trong section + 1; duy nhất trong từng section; **không bao giờ đổi** sau khi tạo; dòng `locked` của template giữ `id` theo template.
- **Dòng ảnh:** `{ id, department, summary, files[] }`. `id` = `g1`, `g2`… do hệ thống cấp (số lớn nhất + 1, thay cho `g<Date.now()>` hiện tại), đồng thời là tên thư mục `photos/<id>/`. Tên phòng ban người dùng nhập nằm ở `department`, đổi tên không làm đổi `id`.
- **Giá trị:** ngày `yyyy-MM-dd`, giờ ISO 8601, số là số JSON (`1.5`), ô dropdown chỉ lưu giá trị đã chọn (chuỗi).

### 3.4 Không lưu trong JSON
`docCodePattern`, `location` (= `unit`), `required`, `order` của section, tiền tố số của sheet, `type`/`align`/`width`/`layout`/`single`/`perRow`/`ui`, `options`, `catalog`, `from`, `readonly`, `default`, `reset`, `carry`, `formula`, cờ `total`/`locked` của cột, danh sách `subcontractorOptions`, bộ đếm mã tự tăng (đã có trong sổ `ReportRegistry`).

## 4. Cấu trúc file lưu

```
data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/
├── <DRP|WRP>_<yyyyMMdd>_<Unit>.json    nguồn dữ liệu (backup)
├── <DRP|WRP>_<yyyyMMdd>_<Unit>.html         bản xem/in độc lập, nhúng JSON
├── photos/<GroupID>/…              ảnh (giữ nguyên)
└── <DocCode>_signed.pdf            bản ký (giữ nguyên)
```

## 5. File HTML độc lập

- Nội dung: markup bản in đã render + CSS in + `<script type="application/json" id="report-data">` chứa nguyên JSON.
- Render bằng chính các component `ReportPrint*` hiện có (qua `renderToStaticMarkup`) nên giống bản PDF đang dùng. CSS in hiện nằm trong chuỗi của `handlePrint` (`ReportPrintPreviewModal.tsx`) cần tách thành module dùng chung.
- Nút "In" gọi `window.print()`; tên file PDF mặc định lấy từ `docCode`.
- Logo nhúng base64. Ảnh hiện trường: **chưa chốt** (nhúng base64 đã nén hay giữ link `photos/…`), xem mục 11.

## 6. Luồng Lưu / Mở

- **Lưu** (`ReportDataService.saveReport`): `checkMaintenance()` → ghi `.json` (giữ ETag) → ghi `.html` → ghi sổ `ReportRegistry`. HTML lỗi thì cảnh báo nhưng vẫn giữ JSON (cách xử lý như ghi sổ).
- **Mở** (`loadReport`): ưu tiên `.json`; chưa có thì đọc `.xlsx` cũ (đường dự phòng); chưa có gì thì lấy báo cáo ngày trước, rồi template.
- Khi mở: ghép cột với template theo `id`, tính lại công thức và `totals`.
- **Sao chép từ ngày trước:** áp dụng luỹ kế (`carry`) và reset theo template; bỏ ảnh; `reportNo` do sổ cấp.

## 7. Cập nhật template (giữ cơ chế hiện nay)

- So ngày trong `template` của JSON với bản mới nhất trong `_template/`; cũ hơn thì hiện khung vàng "Cần cập nhật template mới" + nút **Apply**.
- Apply ghép theo `id`: thêm cột và section mới (ô trống hoặc mặc định), giữ dữ liệu cũ, bỏ cột đã bị xoá khỏi template.
- Mỗi lần sửa template phải đổi ngày trong tên file; cột phải giữ `id` ổn định giữa các bản.
- Cột có trong JSON mà template đã xoá (chưa Apply): hiển thị như ô thường, chưa có công thức/dropdown.

## 8. Dropdown từ SharePoint list

- Dùng lại list `CatalogItems`. Thêm danh mục cố định với `Unit = ALL`: `weather`, `event_type`, `event_status`, `severity`, `issue_status`, `equip_status`, `supplier`, `subcontractor`.
- Template khai báo cột nào dùng danh mục nào theo `id` cột. "Khác…" ghi thẳng vào list.
- Quyền như hiện nay: người báo cáo chỉ thêm, admin toàn quyền.

## 9. Phần code bị ảnh hưởng

| Phần | Thay đổi |
|---|---|
| `types/report.types.ts` | thêm kiểu JSON, ánh xạ với `ReportDoc` |
| `utils/reportCodec.ts` | giữ làm đường đọc dự phòng xlsx |
| mới `utils/reportJson.ts` | `docToJson`, `jsonToDoc`, tính lại công thức/`totals`, kiểm tra schema |
| `services/ReportFileService.ts` | ghi/đọc file text (json, html) |
| `services/ReportDataService.ts` | `saveReport`, `loadReport`, `getSavedDocCode` |
| `components/ReportPrint*`, `ReportPrintPreviewModal.tsx` | tách CSS in, thêm `buildStandaloneHtml` |
| `ReportPhotoSectionEditor.tsx` | `id` nhóm ảnh `g<n>` thay `g<Date.now()>` |
| `ReportCalendarView`, `ReportOverviewView`, `ReportFileService.getReportDayInfos` | phát hiện "đã nhập" bằng `.json` thay vì `.xlsx` |
| `ReportSignedDocument`, `utils/reportFileNaming.ts` | đọc `docCode` từ JSON |
| `AGENTS.md` | cập nhật mục Report: kiến trúc JSON, bỏ quy tắc Excel |

## 10. Các giai đoạn

1. **Schema + chuyển đổi:** `reportJson.ts`, test round-trip `ReportDoc ↔ JSON` với file mẫu.
2. **Lưu/Mở JSON** (kèm đường dự phòng xlsx) và sửa các chỗ phụ thuộc `.xlsx`.
3. **HTML độc lập:** tách CSS, sinh HTML, nút In, nhúng logo/ảnh.
4. **Dropdown từ SharePoint list:** thêm danh mục, nối cột với danh mục theo `id`.
5. **Apply template** theo `id` cột, ghép dữ liệu cũ.
6. **Script chuyển hàng loạt** xlsx → json + html (admin chạy một lần).
7. **Kiểm thử + tài liệu:** `tsc`, `eslint`, `lint:styles`, `build`; thử trên SharePoint thật; cập nhật AGENTS.md.

### Kiểm thử chính
- Round-trip JSON không mất dữ liệu; Apply thêm/bớt cột giữ đúng dữ liệu cũ.
- Mở file HTML trực tiếp trên máy (không mạng), bản in giống PDF hiện tại.
- Báo cáo copy từ ngày trước: luỹ kế đúng, reset đúng, ảnh không sao chép.
- Lịch 14 ngày và Tổng quan vẫn đúng trạng thái khi chuyển sang `.json`.
- Mở báo cáo xlsx cũ rồi Lưu: sinh đủ json + html.

## 11. Cần chốt trước khi code

1. **Ảnh trong HTML:** nhúng base64 (HTML độc lập thật, nặng) hay giữ link `photos/…` (nhẹ, cần cả thư mục)?
2. **Xlsx cũ:** giữ đường đọc dự phòng bao lâu; có cần script chuyển hàng loạt không?
3. **Dòng "Tóm tắt" rỗng:** HTML có ẩn dòng này khi `summary` trống không? (Bản in hiện không in dòng này ở Thời tiết, HSE, Ký duyệt, Chỉ đạo, Ảnh.)

## 12. Rủi ro

- Nhiều chỗ ngầm định `.xlsx` (lịch, Tổng quan, đặt tên, upload bản ký, thông báo cập nhật template); cần rà kỹ ở giai đoạn 2.
- Render tĩnh cần các component in không dùng hook/DOM của trình duyệt (`HeaderLogo` dùng state/`onError`).
- HTML nhúng ảnh base64 có thể nặng; cần nén và giới hạn.
- `created.by` do trình duyệt ghi, không đáng tin tuyệt đối; truy vết thật dựa vào Created/Modified By của SharePoint.
- Web mở báo cáo phụ thuộc template để lấy hành vi sửa; template đổi hoặc xoá `id` cột thì cột đó chỉ hiển thị như ô thường cho tới khi Apply.
- Hai người Lưu cùng lúc: ETag chặn ghi đè file (như hiện nay).

## 13. Trạng thái triển khai

**Đã làm (chạy thử bằng `dev/`, chưa thử trên SharePoint thật):**
- App chỉ đọc/ghi `.json` (không còn codec Excel). `utils/reportJson.ts`: `jsonToDoc` / `docToJson` / `docFromTemplate` / `applyTemplate` / `carryForwardDoc`, kiểm tra `schema`.
- **Template JSON** (`_template/template_DRP_<Unit>_<yyyyMMdd>.json`) giữ mọi thứ về cách nhập (kiểu ô, danh mục, công thức, mặc định, luỹ kế, bố cục); mẫu 4 đơn vị trong `data_test/sharepoint_sample/_template/`. Apply template ghép theo `id`, giữ dữ liệu cũ.
- Mục có thể gồm bảng + `fields` (HSE) hoặc chỉ `fields` (ký duyệt, chỉ đạo); mã dòng `r<n>`; mã nhóm ảnh `g<n>`.
- Mọi dropdown (kể cả danh sách trước đây cố định) lấy từ danh mục dùng chung. Nguồn danh mục: file `ActivityP6.xlsx` + list `CatalogItems` (file mẫu đã thêm các danh mục `weather`, `event_type`, `event_status`, `supplier`, `equipment_status`, `issue_severity`, `issue_status`, `subcontractor`). Chạy thử cục bộ dùng list giả lập nạp từ `CatalogItems_list.xlsx` (`dev/`). Đơn vị thi công chọn từ danh mục `subcontractor` (không còn `subcontractorOptions` trong JSON).
- **File HTML độc lập** sinh lúc Lưu: render tĩnh cùng `ReportPrintDocument` của cửa sổ in, nhúng logo + ảnh (nén cạnh dài 1280px, base64) + toàn bộ JSON. Quyết định các điểm mở: ảnh nhúng base64 có nén; không giữ đường đọc Excel cũ; dòng `Tóm tắt | Summary` vẫn in như bản in hiện tại (HTML/in chưa ẩn khi trống).
- `ActivityP6.xlsx` gọn thành MỘT sheet `ActivityID | Unit | Description` (lọc theo đơn vị trên web, mô tả hiện kèm mã trong dropdown); nguồn tách theo ai được ghi: file chỉ-đọc cho Activity (admin), list `CatalogItems` cho phần người dùng được thêm.
- Khung chạy thử `dev/` (SharePoint giả lập trong bộ nhớ), 9 test chuyển đổi (`npm run test:report`).

**Chưa làm / cần theo dõi:**
- Thử trên SharePoint thật (list `CatalogItems`, ghi `.html`, quyền, ETag 412, dung lượng HTML khi nhiều ảnh lớn).
- Không có script chuyển báo cáo Excel cũ sang JSON (không dùng báo cáo cũ).
- Báo cáo tuần vẫn là placeholder.
