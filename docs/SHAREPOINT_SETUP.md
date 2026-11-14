# Hướng dẫn cài đặt SharePoint cho module Report (bản lưu JSON)

Dành cho admin triển khai lần đầu (site dev: `https://ytf84.sharepoint.com/sites/metro-checkin-dev`). Làm theo thứ tự các mục; mục 8 là danh sách kiểm tra sau khi cài.
> **Tự động hoá:** mục 2 (tạo thư viện, upload file) và mục 3 (tạo list) có thể chạy bằng script `docs/sharepoint_setup/setup-sharepoint.mjs` (xem `docs/sharepoint_setup/README.md`); phân quyền vẫn làm tay.

Tất cả file mẫu nằm trong `data_test/sharepoint_sample/` (riêng các file `*_list.xlsx` tạo list nằm ở `docs/sharepoint_setup/`) (thư mục này mô phỏng đúng cấu trúc thư viện SharePoint).

> Lưu ý: app mới chưa từng chạy trên SharePoint thật (mới chạy thử bằng khung giả lập `dev/`). Lần cài đầu hãy làm kỹ mục 8; gặp lỗi xem mục 9 hoặc gửi nguyên văn thông báo để sửa code.

## 1. Tổng quan thành phần

| Thành phần | Loại | Dùng để | Ai ghi |
|---|---|---|---|
| `DailyReportLibrary` | Thư viện tài liệu | Template, danh mục Activity, logo, báo cáo hằng ngày | Xem bảng ở mục 2 |
| `UserProfiles` | List | Nhận diện người đăng nhập, vai trò admin/user | Chỉ admin |
| `ReportRole` | List | Ai được làm báo cáo của đơn vị nào (một người có thể nhiều dòng) | Chỉ admin |
| `ReportRegistry` | List | Sổ báo cáo: số báo cáo tự tăng, trạng thái từng ngày, tab Tổng quan, lịch 14 ngày | Người báo cáo thêm/sửa, admin toàn quyền |
| `CatalogItems` | List | Danh mục cho dropdown (Element, nhân lực, thiết bị, vật tư, thời tiết, trạng thái, đơn vị thi công…) | Người báo cáo chỉ thêm, admin toàn quyền |
| `SiteAssets/portal_config.json` | File | Chế độ bảo trì (Shutdown Server) | Admin (tự tạo khi bật bảo trì lần đầu) |

Tên thư viện/list phải đúng chữ hoa/thường như trên (tên thư viện đặt ở `REPORT_CONFIG.LIBRARY_NAME` trong `src/modules/report/constants/reportConfig.ts`).

## 2. Thư viện tài liệu `DailyReportLibrary`

Tạo thư viện tài liệu tên `DailyReportLibrary` rồi tạo cấu trúc sau (thư mục `data/<Unit>/` và các thư mục con **do web part tự tạo** khi lưu, trừ khi bạn tạo sẵn `data/<Unit>` để phân quyền theo đơn vị, xem 2.4):

```
DailyReportLibrary/
├── ActivityP6.xlsx           danh mục Activity (một sheet)
├── _template/
│   ├── template_DRP_ST01_20261114.json
│   ├── template_DRP_ST02_20261114.json
│   ├── template_DRP_ST04_20261114.json
│   └── template_DRP_ST05_20261114.json
├── resource/
│   ├── thaco_logo.jpg
│   └── crec_logo.jpg
└── data/                     (tự tạo)
    └── ST01/DRP_20261114_ST01/
        ├── DRP_20261114_ST01.json
        ├── DRP_20261114_ST01.html
        ├── photos/g1/…
        └── HCM2_…_signed.pdf
```

### 2.1 File cần upload (lấy từ `data_test/sharepoint_sample/`)

| File / thư mục | Upload vào | Quyền |
|---|---|---|
| `ActivityP6.xlsx` | gốc thư viện | Người dùng: chỉ đọc. **Chỉ admin được ghi** (bảo mật thật cho mã Activity) |
| `_template/template_DRP_<Unit>_<yyyyMMdd>.json` | `_template/` | Người dùng: chỉ đọc. Admin: ghi |
| `resource/thaco_logo.jpg`, `crec_logo.jpg` | `resource/` | Người dùng: chỉ đọc. Admin: ghi |

**Không upload:** `data/` (báo cáo demo, chỉ để tham khảo cấu trúc; muốn thử thì upload nguyên thư mục `data/ST01/DRP_20261114_ST01`), các file `*_list.xlsx` (dùng để tạo list ở mục 3; không còn file `ReportRole.xlsx`, phân quyền nằm ở list `ReportRole`), `_template/*.xlsx` (template Excel cũ, app không đọc), `*.docx`, `~$*`.

### 2.2 `ActivityP6.xlsx` (một sheet `ActivityP6`)

Dòng 1: `ActivityID | Unit | Description`; mỗi dòng sau một Activity.
- `Unit` = mã đơn vị (vd `ST01`), hoặc `ALL` = dùng chung mọi đơn vị. Thêm đơn vị mới chỉ cần thêm dòng.
- `Description` không bắt buộc, hiện kèm mã trong dropdown ("CV-ST01-00001 – Thi công tường vây"); giá trị lưu vẫn là mã.
- Cột tìm theo tên tiêu đề nên thứ tự cột tuỳ ý. Mã trùng trong cùng một đơn vị bị báo cảnh báo. Muốn bỏ mã thì xoá dòng (báo cáo cũ giữ nguyên giá trị đã lưu).
- Người dùng không thêm được Activity từ giao diện; chỉ admin sửa file. Web nhớ nội dung file 5 phút trong bộ nhớ trang (sửa xong, tải lại trang hoặc chờ 5 phút).

### 2.3 Template JSON `_template/template_DRP_<Unit>_<yyyyMMdd>.json`

- Mỗi đơn vị một file; **ngày `yyyyMMdd` trong tên file là phiên bản**, web luôn dùng bản có ngày lớn nhất.
- Sửa nội dung template (thêm/bớt mục, cột, trường, đổi danh mục, công thức…) thì **upload file mới với ngày mới hơn** (giữ file cũ hay xoá đều được). Báo cáo đang dùng bản cũ sẽ hiện khung vàng "Cần cập nhật template mới"; người dùng bấm "Apply template mới" thì dữ liệu cũ được giữ và cột/mục mới được thêm vào.
- Giữ nguyên `id` của mục/cột/trường giữa các bản (đổi `id` nghĩa là cột mới, dữ liệu cột cũ bị bỏ khi Apply).
- Cấu trúc file: `docs/REPORT_JSON_PLAN.md` mục 3 và `src/modules/report/types/reportJson.types.ts`; mẫu 4 đơn vị: `data_test/sharepoint_sample/_template/`.
- Đơn vị chưa có báo cáo nào bắt buộc phải có template của đơn vị đó, nếu không web báo "Chưa có template cho STxx".

### 2.4 Phân quyền thư viện (bảo mật thật nằm ở SharePoint)

1. Ngắt kế thừa quyền ở thư viện `DailyReportLibrary`: nhóm người dùng báo cáo = **Read** ở gốc, `_template`, `resource`; admin (Site Owners) = Full Control.
2. Thư mục `data/`: người báo cáo cần quyền **Contribute** (tạo thư mục, ghi/ghi đè file, tải ảnh, upload bản ký). Muốn mỗi nhóm chỉ ghi vào đơn vị của mình: admin tạo sẵn `data/<Unit>` (vd `data/ST01`), ngắt kế thừa và cấp Contribute cho nhóm của đơn vị đó trên chính thư mục này; khi đó nhóm báo cáo không cần quyền ghi ở `data/`. Web part tự tạo các thư mục con `DRP_<yyyyMMdd>_<Unit>/` bên trong.
3. Nên chỉ admin được **xoá** file bản ký (`*_signed.pdf`): dùng mức quyền tuỳ chỉnh "Contribute không xoá" cho người báo cáo.
4. Bật versioning cho thư viện (mỗi lần ghi đè `.json` có bản lịch sử).

## 3. Bốn list trên site

Đủ bốn list: `UserProfiles`, `ReportRole`, `ReportRegistry`, `CatalogItems`. Mỗi file `docs/sharepoint_setup/*_list.xlsx` có sẵn sheet `Huong_dan` mô tả từng bước và kiểu cột; phần dưới tóm tắt những điểm bắt buộc.

### 3.1 `UserProfiles` (từ `docs/sharepoint_setup/UserProfiles_list.xlsx`)
- Tạo list: New > List > From Excel > chọn bảng `UserProfiles`. Thay dữ liệu mẫu bằng danh sách người dùng thật.
- Cột quan trọng nhất: `Email` (UPN M365), `Role` (`ADMIN` hoặc `USER`).
- Cột `Active` (Yes/No): `No` = tạm khoá người dùng mà không xoá dòng; trống = Yes. Luôn giữ ít nhất một admin `Yes`. Import từ Excel không có kiểu Yes/No: chọn Single line of text (chữ Yes/No) hoặc Number (1/0), web đọc được cả hai; cột Yes/No thật chỉ có khi tạo bằng script `docs/sharepoint_setup`. Cột `Active` của `ReportRole` cũng vậy.
- Quyền: người thường chỉ **đọc**; chỉ admin thêm/sửa/xoá (vai trò ADMIN trong list chính là quyền admin của web part).

### 3.2 `ReportRole` (từ `docs/sharepoint_setup/ReportRole_list.xlsx`): phân quyền báo cáo theo người
- Tạo list từ Excel (bảng `ReportRole`), đặt tên list đúng **`ReportRole`**. Cột: `Email`, `DisplayName`, `Unit`, `DocID`, `DocTitle` (Single line of text; `DisplayName` có thể là cột Title) và `Active` (**Yes/No**).
- Mỗi dòng = một người × một đơn vị × một loại báo cáo (`DocID` = `DailyReport`; `WeeklyReport` chỉ là placeholder). Một ô `Email` có thể chứa nhiều email ngăn bằng `;` hoặc dấu phẩy. `DisplayName` là tên người phụ trách, tự điền vào ô "Trưởng ga / CHT" của mục ký duyệt.
- Tên cột cùng ý nghĩa dùng cùng tên ở các list khác: `Email`, `DisplayName` (như `UserProfiles`), `Unit`, `DocID` (như `ReportRegistry`).
- Người dùng chỉ thấy ô báo cáo có dòng `Active = Yes` với email của mình. Admin (`Role` = `ADMIN` trong `UserProfiles`) thấy tất cả mà không cần dòng ở đây. **Người dùng phải có trong cả `UserProfiles` và `ReportRole`.**
- Quyền: người báo cáo chỉ **đọc**; chỉ admin thêm/sửa/xoá. Bật versioning.

### 3.3 `ReportRegistry` (từ `docs/sharepoint_setup/ReportRegistry_list.xlsx`)
- Tạo list từ Excel (bảng `ReportRegistry`); cột `ReportDate`, `SavedAt`, `SignedAt`, `Status`, `Counters`, `SignedUploadBy`, `SignedRemovedAt`, `SignedRemovedBy` (ba cột cuối tuỳ chọn: người upload bản ký hiện tại, lần xoá bản ký gần nhất) đều chọn **Single line of text** (không chọn Date/Choice); `ReportNo` kiểu Number, **0 chữ số thập phân**.
- Xoá dòng mẫu `ST00|DailyReport|2000-01-01` sau khi tạo.
- List settings: cột `Title` bật **Indexed** và **Enforce unique values**; tạo Index cho `Title`, `Unit`, `ReportDate`, `ReportNo` (tránh ngưỡng 5000 dòng); bật versioning.
- Quyền: người báo cáo được **thêm và sửa**, **không xoá** ("Contribute không xoá"); admin toàn quyền. Không xoá dòng cũ (số báo cáo kế tiếp lấy từ số lớn nhất trong list).

### 3.4 `CatalogItems` (từ `docs/sharepoint_setup/CatalogItems_list.xlsx`)
- Tạo list từ Excel (bảng `CatalogItems`), mọi cột **Single line of text**. File đã gồm 63 dòng: mã nhân lực/thiết bị dùng chung (`Unit = ALL`), Element và Vật tư theo từng đơn vị, và các danh sách chọn dùng chung (`ALL`): `weather`, `event_type`, `event_status`, `supplier`, `equipment_status`, `issue_severity`, `issue_status`, `subcontractor`.
- List settings: cột `Title` bật **Indexed** + **Enforce unique values**; Index thêm cho `Unit`; bật versioning.
- Quyền: người báo cáo chỉ **thêm** dòng (không sửa/xoá); admin toàn quyền (sửa chính tả, gán mã, xoá mục sai).
- Ý nghĩa cột: `Title` = khoá duy nhất do web tự sinh (`<Unit>|<Catalog>|code|<giá trị chuẩn hoá>`); `Unit` = mã đơn vị hoặc `ALL`; `Catalog` = tên danh mục chữ thường; `Code` = mã/giá trị; `Link` = giá trị ràng buộc (vd mã nhân lực ↔ "Chỉ huy trưởng").
- Thêm giá trị cho danh sách chọn dùng chung: thêm dòng `Unit = ALL`, `Catalog` đúng tên, `Code` = giá trị, `Title` = `ALL|<Catalog>|code|<giá trị chuẩn hoá>` (không dấu, chữ thường, khoảng trắng đơn).

## 4. `SiteAssets/portal_config.json` (chế độ bảo trì)
Không cần tạo tay: admin vào Admin Config > Shutdown Server rồi bật/tắt, web part ghi file này. Cần để quyền **ghi vào SiteAssets** chỉ cho admin. Không có file nghĩa là không bảo trì.

## 5. Đóng gói và cài app

1. **Tắt `npm run start`** nếu đang chạy (build xoá `dist/` đang phục vụ).
2. Tuỳ chọn: tăng `solution.version` trong `config/package-solution.json` (vd `1.0.1.0`) để các site nhận bản mới.
3. Chạy `npm run build` → ra `sharepoint/solution/workspace_portal.sppkg`.
4. Mở App Catalog (vd `https://ytf84.sharepoint.com/sites/appcatalog`) > **Apps for SharePoint** > upload `.sppkg` (ghi đè bản cũ) > **Deploy**. Tích **Enable this app and add it to all sites** (hoặc thêm app vào từng site) rồi Deploy.
5. Trên site: mở trang cần dùng > Edit > `+` > **Workspace Portal** > **Republish**.
6. Hard reload trình duyệt (Ctrl/Cmd + Shift + R) nếu vẫn thấy bản cũ.

Web part tự lấy site hiện hành từ `pageContext.web.absoluteUrl` (hoặc cấu hình `siteUrl` ở Property Pane nếu dữ liệu nằm ở site khác).

## 6. Quy trình thêm đơn vị mới (vd ST06)
1. List `ReportRole`: thêm các dòng `Email | DisplayName | ST06 | DailyReport | Báo cáo ngày | Yes`.
2. `_template/`: upload `template_DRP_ST06_<yyyyMMdd>.json` (sao chép template đơn vị khác, đổi dòng mẫu nếu cần).
3. `ActivityP6.xlsx`: thêm các dòng Activity với `Unit = ST06`.
4. `CatalogItems`: thêm Element/Vật tư riêng của ST06 (nếu có); mã nhân lực/thiết bị `ALL` đã dùng chung.
5. Tạo sẵn thư mục `data/ST06` và cấp quyền Contribute cho nhóm của đơn vị (nếu phân quyền theo đơn vị; xem mục 2.5).

## 7. Sao lưu
Mỗi báo cáo là một thư mục tự chứa: `.json` là dữ liệu gốc, `.html` là bản xem/in độc lập (mở offline), `photos/` là ảnh, `*_signed.pdf` là bản ký. Sao chép thư mục `data/` là đủ để sao lưu; sổ `ReportRegistry` chỉ là chỉ mục có thể dựng lại.
SharePoint Online thường buộc file `.html` trong thư viện **tải về** thay vì mở trực tiếp: người dùng tải về rồi mở bằng trình duyệt.

## 8. Kiểm tra sau khi cài (làm lần lượt)

| # | Thao tác | Kết quả đúng |
|---|---|---|
| 1 | Mở web part bằng tài khoản **admin** | Trang chủ hiện tên, vai trò Admin; có Admin Config |
| 2 | Vào Quản lý Báo cáo | Tab "Tổng quan" hiện danh sách đơn vị (không có cảnh báo vàng về `ReportRegistry`) |
| 3 | Mở một đơn vị, ngày hôm nay (chưa có báo cáo) | Form tạo từ template, số báo cáo `00001`, có đủ 11 mục |
| 4 | Mở dropdown Activity ở mục Khối lượng | Có mã + mô tả của đúng đơn vị đó; không có cảnh báo danh mục |
| 5 | Chọn "Khác…" ở một ô (vd Thời tiết), nhập giá trị mới | Giá trị được chọn; list `CatalogItems` có thêm một dòng |
| 6 | Chọn một mã nhân lực | Ô "Loại" tự điền và khoá |
| 7 | Điền đơn vị thi công, bấm **Lưu báo cáo** | Thông báo "Đã lưu báo cáo số 00001"; thư mục `data/<Unit>/DRP_<yyyyMMdd>_<Unit>/` có `.json` và `.html`; `ReportRegistry` có 1 dòng `unsigned` |
| 8 | Tải `.html` về và mở bằng trình duyệt (tắt mạng) | Hiện đủ bản in A4, nút In/Lưu PDF hoạt động |
| 9 | Tải một ảnh ở mục Ảnh hiện trường, Lưu lại | Ảnh nằm ở `photos/g1/`; `.html` mới nhúng ảnh |
| 10 | Upload bản PDF đã ký | File `<MãTàiLiệu>_signed.pdf` trong thư mục; sổ chuyển `done` |
| 11 | Mở ngày hôm sau | Kế thừa ngày trước: luỹ kế cộng dồn, ô "hôm nay" về 0, số báo cáo `00002` |
| 12 | Đăng nhập bằng tài khoản **thường** | Chỉ thấy ô báo cáo của mình; không có Admin Config; không xoá được dòng list |

## 9. Lỗi thường gặp

| Thông báo / hiện tượng | Nguyên nhân thường gặp | Cách xử lý |
|---|---|---|
| "Chưa có template cho STxx trong thư mục _template/…" | Chưa upload template hoặc sai tên file | Upload đúng `template_DRP_<Unit>_<yyyyMMdd>.json` vào `_template/` |
| "template_….json: Template không hợp lệ…" | File JSON sai cú pháp / thiếu `sections` | Mở file bằng trình soạn thảo kiểm tra JSON |
| "Danh mục: Chưa có file danh mục ActivityP6.xlsx…" | Chưa upload hoặc sai tên/vị trí | Đặt `ActivityP6.xlsx` ở **gốc** thư viện |
| "Sheet … cần có cột ActivityID và Unit ở dòng 1" | Sai tiêu đề cột | Sửa dòng 1 thành `ActivityID | Unit | Description` |
| "Chưa có Activity nào của đơn vị …" | Không có dòng `Unit` khớp | Thêm dòng cho đơn vị (hoặc `ALL`) |
| "Chưa đọc được danh mục dùng chung (kiểm tra list CatalogItems…)" | List chưa tạo/sai tên, thiếu quyền đọc | Tạo list `CatalogItems`, cấp quyền đọc |
| "Chưa thêm được vào danh mục dùng chung" | Thiếu quyền thêm hoặc sai tên cột | Cấp quyền Add; giữ nguyên tên hiển thị 5 cột |
| "Đã lưu báo cáo nhưng chưa ghi được vào sổ ReportRegistry…" | List chưa tạo, sai kiểu cột (Date/Choice), thiếu quyền | Làm lại mục 3.2; đọc phần "Chi tiết" trong thông báo |
| "Đã lưu dữ liệu báo cáo nhưng chưa tạo được file HTML xem độc lập" | Thiếu quyền ghi hoặc tenant chặn `.html` | Kiểm tra quyền; dữ liệu `.json` vẫn an toàn |
| Lưu lỗi "412 / Precondition failed" | Người khác vừa lưu cùng báo cáo | Bấm **Làm mới** rồi sửa lại |
| Không thấy ô báo cáo nào ("Chưa được phân quyền…") | Email chưa có dòng `Active = Yes` trong list `ReportRole` (hoặc chưa có trong `UserProfiles`) | Thêm/sửa dòng phân quyền |
| "Chưa đọc được phân quyền báo cáo" (kèm lỗi 404/403) | List `ReportRole` chưa tạo, sai tên hoặc người dùng không có quyền đọc | Tạo list đúng tên `ReportRole`, cấp quyền đọc |
| Hệ thống đang bảo trì | Admin bật Shutdown Server | Admin tắt ở Admin Config |
| Dropdown vẫn thiếu giá trị mới sửa ở `ActivityP6.xlsx` | Web nhớ file 5 phút | Tải lại trang |
| "Tài khoản của bạn chưa có trong danh sách phân quyền UserProfiles" dù đã có dòng | Tên cột nội bộ khác tên hiển thị (list tạo từ Excel hoặc đổi tên cột); hoặc trang đang chạy bản đã deploy cũ | Bản mới tự tra tên cột; mở F12 → Console xem dòng `[PortalShell]`/`[UserProfileService]`; đảm bảo đang chạy bản local hoặc đã deploy lại |
| "Chưa có template cho STxx trong thư mục _template/" | Chưa upload `template_DRP_<Unit>_<yyyyMMdd>.json` | Upload template JSON vào thư mục `_template` ở gốc thư viện (không upload file `.xlsx`) |
| Workbench không nạp bản local (Console: CORS / loopback `localhost:4321`) | Chrome chặn trang SharePoint truy cập máy local | Cho phép **Local network access** cho site trong Chrome, dùng link `debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2F...` rồi **Load debug scripts** |
