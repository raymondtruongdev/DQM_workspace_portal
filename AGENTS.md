# THADICO DQM - Workspace Portal (SPFx)

## AI Agent & Developer Reference Guide

> **Project Location:** `/Volumes/DATA/THADICO_local/REPO_DQM/DQM_Check_In_App`
> **Target Environment:** SharePoint Online (Dev: `https://ytf84.sharepoint.com/sites/metro-checkin-dev`)
> **Architecture Model:** SharePoint Framework (SPFx 1.24.1) Client-Side Web Part
> **Licensing Tier:** Free (Standard Microsoft 365 Seeded License - Zero Power Apps Premium cost)

---

## 1. Project Background & Architectural Decision

### Why SPFx Instead of Power Apps Code App?

* **Tránh rào cản bản quyền:** Power Apps Code App luôn bị Microsoft gán nhãn *"Premium"*, ép người dùng mua license riêng.
* **Giải pháp SPFx:** Chạy trực tiếp trong SharePoint Online bằng tài khoản Microsoft 365 có sẵn của người dùng:
  - **Miễn phí 100%:** Sử dụng license M365 tiêu chuẩn, không giới hạn số lượng người dùng.
  - **Hiệu năng gốc (Native):** Kết nối trực tiếp REST API qua `SPHttpClient`, không có iframe hay container trễ.
  - **SSO tự động:** Tận dụng phiên đăng nhập Microsoft Entra ID hiện tại.

---

## 2. Technology Stack & Tooling

* **Framework:** SharePoint Framework (SPFx) `1.24.1`
* **UI Library:** React `18.3.1` (Hooks, Functional Components)
* **Build System:** `@rushstack/heft` `1.3.2` + Webpack `5.105.4` + TypeScript `~5.8.0`
* **Styling:** Custom Modern SaaS Design System (CSS thuần, token trong `src/core/styles/tokens.css`, accent `#1d4ed8`, một tông xanh duy nhất; xem mục 8)
* **Icons:** Fluent UI System Icons (`@fluentui/react-icons`, bản Regular) qua component `core/components/Icon.tsx` (tree-shaken, bundle ~205KB)
* **HTTP Client:** `@microsoft/sp-http` (`SPHttpClient` v1)

---

## 3. SharePoint Lists & Data Architecture

WebPart tự động nhận diện site hiện hành qua `pageContext.web.absoluteUrl` (hoặc cấu hình `siteUrl` qua Property Pane). Dữ liệu gồm danh sách `UserProfiles` (phân quyền) và thư viện tài liệu `CapNhatCongViec` (module Report):

### 1. `UserProfiles` (Role & Department Authorization List)
* `Email` / `Title` / `field_1` (Text): Email tài khoản Microsoft 365 (UPN).
* `Username` / `field_2` (Text): Mã định danh hoặc Tên người dùng.
* `DepartmentCode` / `field_3` (Text): Mã phòng ban (VD: ADMIN, DIRC...).
* `DepartmentName` / `field_4` (Text): Tên phòng ban hiển thị (VD: Admin site, Bộ phận Giao diện...).
* `Role` / `field_5` (Text): Vai trò người dùng (`admin`, `user`).
* `Active` (Yes/No, tuỳ chọn): `No` = tạm khoá, coi như không có trong list (`UserProfileService` lọc bỏ dòng này nên người đó không vào được portal và không hiện trong bộ chuyển vai trò); trống hoặc không có cột = đang hoạt động (`parseProfileActive`). Khoá nhầm dòng admin cuối cùng thì phải sửa tay trên SharePoint. Vẫn chỉ là kiểm tra ở giao diện, không thay cho phân quyền SharePoint.

---

## 4. Directory & File Structure

```text
DQM_Check_In_App/
├── config/                      # Cấu hình SPFx build, workbench serve & package solution
├── sharepoint/solution/         # Gói cài đặt thành phẩm (workspace_portal.sppkg)
├── scripts/check-styles.js      # Kiểm tra quy ước style (`npm run lint:styles`)
├── src/
│   ├── entry/                   # Điểm vào SPFx: PortalWebPart.ts, manifest, loc/
│   ├── app/                     # Khung portal: PortalShell.tsx (auth + chọn module), moduleRegistry.tsx, IPortalProps.ts
│   ├── core/                    # Dùng chung cho mọi module (KHÔNG import từ modules/)
│   │   ├── auth/                # Session, UserProfileService, extractCleanEmail, checkIsAdmin
│   │   ├── services/            # SharePointClient (generic CRUD qua SPHttpClient)
│   │   ├── components/          # React dùng chung: Icon, PageHeader, FilterChips, SearchInput, State (Loading/Empty), layout/ (TopAppBar, ToastAlert)
│   │   ├── types/ constants/ utils/
│   │   └── styles/              # Design System (xem mục 8): tokens.css, base.css, components/*.css, layout/*.css, index.ts
│   ├── modules/
│   │   ├── home/                # HomeView: trang chủ liệt kê module từ registry
│   │   ├── admin/               # Admin Config (chỉ admin): UI Kit (Style Guide), Shutdown Server (chế độ bảo trì)
│   │   └── report/              # Module: Quản lý Report (list Report_Role, Daily report lưu JSON + HTML độc lập, Weekly placeholder, SharePoint data/<Unit>/<DocID>_<yyyyMMdd>/)
│   └── assets/
├── dev/                         # Khung chạy thử cục bộ của module Report (SharePoint giả lập trong bộ nhớ) + test; KHÔNG vào gói web part
├── docs/                        # Tài liệu thiết kế (REPORT_JSON_PLAN.md) + mẫu docs/sample/ (báo cáo JSON, HTML độc lập)
├── data_test/sharepoint_sample/ # Bản mẫu thư viện SharePoint: _template/*.json, data/<Unit>/…, resource/, ActivityP6.xlsx, (file tạo list *_list.xlsx đã chuyển sang `docs/sharepoint_setup/`: UserProfiles, Report_Role, ReportRegistry, CatalogItems)
├── AGENTS.md
└── package.json

---

## 5. Key Components & Behavioral Rules

### 1. Portal Architecture & Authentication Flow (Landing Page + Zero-Click SSO)

* **Portal Landing Page:** On entry, users land on the modern SaaS Welcome Portal (`modules/home/HomeView.tsx`). It greets the user with their verified Microsoft 365 identity, role badge, department, and an overview of system features.
* **Module Tiles:** The Landing page lists modules from `moduleRegistry.tsx` (via `<TileGrid />`); clicking a tile opens that module. Modules flagged `adminOnly` (Admin Config) only show for admins.
* **Admin Config (`modules/admin`, chỉ admin):** màn hình menu giống trang chủ (`<TileGrid />`) gồm **UI Kit (Style Guide)** (tài liệu sống của mọi thành phần dùng chung) và **Shutdown Server** = chế độ bảo trì: admin bật/tắt, cấu hình lưu ở `SiteAssets/portal_config.json` của site (`core/services/PortalConfigService.ts`; cần quyền ghi vào SiteAssets). Khi bật, `PortalShell` chỉ cho admin vào; người dùng thường thấy màn hình "Hệ thống đang bảo trì" kèm thông báo admin nhập; admin thấy cảnh báo trên trang chủ. Cấu hình được đọc lại mỗi lần quay về trang chủ / đổi module; không đọc được file thì coi như không bảo trì.
* **Return Home Navigation:** Inside a module, users can click the **"Trang Chủ"** button to return to the Landing page at any time.
* **Zero-Click SSO & Profile Resolution:** Identity is resolved via Microsoft Entra ID (`pageContext.user.email` / `loginName`) against the `UserProfiles` list.
  - If a matching email is found, role and department are seamlessly assigned.
  - Danh tính luôn được đọc lại từ list `UserProfiles` mỗi lần mở web part; KHÔNG lưu/tin `sessionStorage` (người dùng có thể sửa).
  - Bộ chuyển vai trò Developer / QA ("xem với vai trò khác") trên trang chủ CHỈ hiện cho admin thật (quyền lấy từ `UserProfiles` theo tài khoản M365 đăng nhập, không phải hồ sơ đang xem). Vai trò đang xem chỉ giữ trong bộ nhớ (tải lại trang là về tài khoản thật). Quyền bỏ qua chế độ bảo trì dựa trên admin thật. Lưu ý: kiểm tra quyền này chỉ ở phía giao diện, bảo mật thật phải dựa vào phân quyền SharePoint (list `UserProfiles` chỉ đọc cho người thường, `SiteAssets/portal_config.json` chỉ admin được ghi).

### 2. Data Service (`SharePointClient.ts`)

* Uses `SPHttpClient.configurations.v1`.
* All calls specify `'Accept': 'application/json;odata=nometadata'` to minimize payload overhead.
* Updates use `POST` with headers `'X-HTTP-Method': 'MERGE'` and `'If-Match': '*'`.
* Deletions use `POST` with headers `'X-HTTP-Method': 'DELETE'` and `'If-Match': '*'`.

### 3. Report Management Module (`modules/report`)

* **Storage Model:** Uses the SharePoint Document Library (`CapNhatCongViec`).
  - Phân quyền báo cáo theo người: SharePoint List `Report_Role` (bản mẫu import `docs/sharepoint_setup/Report_Role_list.xlsx`; cột `Email`, `DisplayName` (cột Title), `Unit`, `DocID`, `DocTitle`, `Active` Yes/No). **Quy ước đặt tên cột:** cột cùng ý nghĩa ở các list/file dùng cùng tên (`Email`, `DisplayName` như `UserProfiles`; `Unit`, `DocID` như `ReportRegistry`/`CatalogItems`/`ActivityP6.xlsx`); cột khác ý nghĩa thì khác tên (vd `Active` của `Report_Role` khác `Status` của `ReportRegistry`); tên cột không có dấu cách.
  - Danh mục dùng chung gồm HAI nguồn (xem mục "Danh mục dùng chung" bên dưới), tách theo AI ĐƯỢC GHI: (1) `ActivityP6.xlsx` ở gốc thư viện (gốc thư viện; bản mẫu `data_test/sharepoint_sample/ActivityP6.xlsx`): MỘT sheet các Activity của mọi đơn vị, người dùng chỉ đọc, CHỈ ADMIN được ghi (bảo mật thật nhờ quyền chỉ-đọc của file); (2) SharePoint List `CatalogItems` (bản mẫu import `docs/sharepoint_setup/CatalogItems_list.xlsx`): Element, Resource nhân lực/thiết bị, Vật tư và các danh sách chọn dùng chung (thời tiết, trạng thái, đơn vị thi công…), người dùng được thêm bằng "Khác…", một list cho mọi nhà ga.
  - Template báo cáo (JSON): `_template/template_DRP_<Unit>_<yyyyMMdd>.json` (mẫu: `data_test/sharepoint_sample/_template/`); ngày trong tên file là phiên bản, bản mới nhất thắng. **App chỉ hỗ trợ JSON, không còn template/báo cáo Excel.**
  - Mỗi báo cáo ngày là một thư mục `data/<Unit>/<DocID>_<yyyyMMdd>/` gồm: `<DocID>_<yyyyMMdd>.data.json` (nguồn dữ liệu, backup), `<DocID>_<yyyyMMdd>.html` (bản xem/in độc lập, nhúng sẵn dữ liệu + ảnh + logo), `photos/<GroupID>/` (ảnh) và `<DocCode>_signed.pdf` (bản ký). Ví dụ `data/ST01/DailyReport_20261110/` (bản mẫu đầy đủ: `data_test/sharepoint_sample/data/ST01/DailyReport_20261110/`). Lưu ý: SharePoint Online thường buộc file `.html` trong thư viện tải về thay vì mở trực tiếp; người dùng tải rồi mở bằng trình duyệt.
  - Logo bản in: `resource/thaco_logo.jpg` (trái) và `resource/crec_logo.jpg` (phải), đặt ở thư mục `resource/` ở gốc thư viện (cùng cấp `_template`, `data`; bản mẫu: `data_test/sharepoint_sample/resource/`); không có file thì bản in hiện chữ THACO / CREC. Người dùng chỉ cần quyền đọc, chỉ admin được ghi. Khuyến nghị ảnh ngang tỉ lệ khoảng 2:1, nền trắng, khoảng 400×200 px (hiển thị tối đa cao 16mm), dưới 100KB.
  - Site Photos Subfolder: `data/<Unit>/<DocID>_<yyyyMMdd>/photos/<GroupID>/` (mỗi bộ phận một thư mục con; tải/xoá trực tiếp từ form qua `ReportFileService`).
* **Role Authorization (`ReportRoleService`, `utils/reportRole.ts`, `core/services/RoleResolutionService.ts`):** đọc lại list `Report_Role` mỗi lần mở module (`getItems` + tra tên nội bộ cột qua `getListFieldMap`; `DisplayName` đọc từ cột Title nếu cần). Email của người đăng nhập khớp dòng `Active = Yes` có `DocID` hợp lệ (`DailyReport`, `WeeklyReport`; chấp nhận biệt danh drp/bcn/wrp/bct). Một ô `Email` có thể chứa nhiều email ngăn bằng `;` `,` hoặc xuống dòng. `DisplayName` của dòng tự điền vào ô "Trưởng ga / CHT". Người dùng phải có trong CẢ `UserProfiles` (nhận diện đăng nhập) và `Report_Role` (được làm báo cáo nào). Admin (`UserProfiles.Role = admin`) thấy mọi phân quyền đang hiệu lực mà không cần dòng riêng. Không đọc được list (chưa tạo, thiếu quyền đọc…) thì hiện trạng thái lỗi "Chưa đọc được phân quyền báo cáo" kèm nút Thử lại (không im lặng hiện trang trống). Người báo cáo chỉ cần quyền ĐỌC list này, chỉ admin được ghi (web đọc toàn bộ list rồi lọc ở giao diện, nên mọi người đọc được email của nhau: bảo mật thật nhờ quyền ghi của list).
* **Lịch 14 ngày & bản đã ký (`ReportCalendarView.tsx`, `ReportSignedDocument.tsx`, `utils/reportCalendar.ts`):**
  - Luồng báo cáo ngày: danh sách tile → **lịch 14 ngày gần nhất** → form nhập của ngày đã chọn (Back từ form về lịch, lịch đọc lại trạng thái). WeeklyReport (placeholder) vào thẳng form.
  - **3 trạng thái mỗi ngày** (đọc từ sổ `ReportRegistry` như tab Tổng quan nên hai màn hình luôn khớp; chưa đọc được sổ thì tạm quét thư mục `data/<Unit>/<DocID>_<yyyyMMdd>/` qua `ReportFileService.getReportDayInfos`; trong sổ: `unsigned` = Chưa upload bản ký, `done` = Hoàn tất, không có dòng = Chưa nhập): **Chưa nhập** = chưa có file báo cáo `.xlsx` (chưa bấm Lưu); **Chưa upload bản ký** = có file `.xlsx` nhưng thư mục chưa có file `.pdf`; **Hoàn tất** = có file `.xlsx` và ít nhất một file `.pdf` trong thư mục (không tính thư mục con `photos/`).
  - **Thẻ "Ngày đang chọn" (`report-day-panel`):** ô ngày (T<tháng>/<ngày>) + tên ngày + nhãn trạng thái + nút hành động; hàng 3 bước (Lưu lần cuối, Upload bản ký, File PDF đã ký; bước xong có dấu tick xanh, tên file hiển thị đầy đủ, tự xuống dòng); chân thẻ có đường dẫn thư mục kèm nút sao chép và gợi ý theo trạng thái.
  - **Bản đã ký:** sau khi Lưu, người dùng in PDF, ký tên rồi upload bản đã ký ở mục "BẢN ĐÃ KÝ | SIGNED DOCUMENT" cuối form (hoặc nút "Upload file đã ký" ở lịch). Chỉ nhận `.pdf`, lưu cùng thư mục với báo cáo, tự đổi tên khi upload thành `<DocCode>_signed.pdf` (cùng tên file PDF in ra + `_signed`, ví dụ `HCM2_1-23_CS-THAC_DCEM-CPZZ-ST01-CWS-DRP-00001_signed.pdf`; upload từ lịch đọc Doc. code từ file báo cáo đã lưu; thiếu Doc. code thì dùng `<DRP|WRP>_<yyyyMMdd>_<Unit>_signed.pdf`). Upload lại = ghi đè, mọi file `*_signed.pdf` khác trong thư mục bị xoá để không còn hai bản. Sổ `ReportRegistry` không lưu tên file ký nên lịch/Tổng quan đọc thư mục khi cần tên hoặc khi bấm tải (`openSignedDocument`). Chỉ upload được khi báo cáo đã lưu và không còn thay đổi chưa lưu. Sửa báo cáo sau khi ký không đổi trạng thái (vẫn Hoàn tất). Upload gọi `checkMaintenance()` trước khi ghi.
  - Lịch không hiển thị người lưu (thông tin này không có trong danh sách file); chỉ hiển thị thời điểm lưu báo cáo và thời điểm upload bản ký. Cả hai là thời điểm **sửa lần cuối** (`TimeLastModified`) của file `.xlsx` và file PDF đã ký trên SharePoint, theo giờ máy người xem; lưu lại báo cáo hoặc thay file ký thì thời điểm đổi theo.
  - **Tab "Tổng quan" (chỉ admin, `ReportOverviewView.tsx`, `utils/reportOverview.ts`)** trên trang danh sách báo cáo (đứng trước tab "Báo cáo của tôi" và là tab mặc định khi admin mở module): hai thẻ tóm tắt: **hôm nay** (số đơn vị Chưa nhập / Chưa upload bản ký / Hoàn tất) và **14 ngày gần nhất** (số báo cáo đơn vị × ngày theo từng trạng thái, kèm phần trăm), **ma trận đơn vị × 14 ngày** (bấm ô mở form của đơn vị ở ngày đó, bấm mã đơn vị mở lịch), **thanh công cụ riêng của ma trận** (làm nổi bật theo trạng thái + tìm theo mã ga/người phụ trách, chỉ tác động lên ma trận), danh sách **Tất cả report** (mọi đơn vị × ngày với cả 3 trạng thái, ngày mới nhất trước, mỗi dòng gồm khối ngày (hôm nay tô xanh), mã đơn vị + nhãn trạng thái có chấm màu, người phụ trách dạng thẻ nhỏ (chữ cái đầu + phần tên sau " - ", mặc định hiện 1 người, "+N" để mở rộng), giờ "Lưu" và "Ký" (chỉ kèm ngày khi khác ngày báo cáo); khi khung danh sách rộng dưới 640px (đo bằng ResizeObserver) các nút xuống hàng dưới. Nút theo trạng thái: Chưa nhập = Nhập báo cáo; Chưa upload bản ký = Upload bản ký; Hoàn tất = Xem báo cáo + nút tải bản ký (chỉ icon); Chưa upload bản ký có thêm nút xem (icon mắt)) có bộ lọc riêng theo trạng thái (Tất cả / Chưa nhập / Chưa upload bản ký / Hoàn tất), theo ngày (Tất cả 14 ngày / Theo ngày với ô chọn ngày, mặc định là hôm nay; cùng một hàng) và theo đơn vị (Tất cả hoặc từng đơn vị) (kèm người phụ trách từ list `Report_Role` và nút **Xuất Excel** (xuất đúng danh sách đang lọc ra `.xlsx`: ngày, thứ, đơn vị, loại báo cáo, trạng thái, số báo cáo, người lập, lưu lúc, file đã ký, upload bản ký lúc, người phụ trách, email; SheetJS chỉ tải khi bấm; `utils/reportExport.ts`)). Chỉ theo dõi báo cáo ngày. Dữ liệu đọc MỘT lần từ sổ `ReportRegistry` (không quét thư mục; ngày không có dòng = Chưa nhập) và nhớ 3 phút trong bộ nhớ trang; nút Làm mới bỏ qua bộ nhớ; list chưa tạo/không đọc được thì hiện cảnh báo. Chưa có ngày nghỉ/lễ (mọi ngày đều tính).
* **Navigation & Tile Buttons (`ReportLandingView.tsx`):**
  - Landing page shows tile buttons based on assigned roles (e.g. `ST01 DailyReport`, `ST01 WeeklyReport`, `ST02 DailyReport`), grouped with unit tabs and search filter.
* **Kiến trúc dữ liệu của Daily Report (`utils/reportJson.ts`, `types/reportJson.types.ts`, `reportTable.ts`, `reportForm.ts`, `reportSections.ts`):** hợp đồng đầy đủ và lý do thiết kế: `docs/REPORT_JSON_PLAN.md`; mẫu: `docs/sample/DailyReport.sample.data.json`.
  - **Nguyên tắc:** file báo cáo `.data.json` chỉ chứa định danh, nhãn và GIÁ TRỊ THẬT (kể cả ô công thức và dòng Tổng) để mở một mình đọc được; mọi thứ quyết định cách NHẬP/TRÌNH BÀY (kiểu ô, danh mục dropdown, công thức, mặc định, luỹ kế, căn lề, độ rộng, bố cục) nằm ở TEMPLATE JSON, ghép với báo cáo theo `id` mục / `id` cột / `id` trường (`jsonToDoc`). Khi mở, web luôn tính lại ô công thức và dòng Tổng, không tin giá trị đã lưu. Cột/mục/trường có trong báo cáo mà template không có (hoặc chưa tải được template) vẫn hiển thị như ô văn bản thường.
  - **File báo cáo:** `schema`, `unit`, `docId`, `date` (`yyyy-MM-dd`), `reportNo` (số nguyên, in 5 chữ số), `docCode` (đã thay `{unit}`/`{reportNo}`; mẫu `docCodePattern` thuộc template nên không lưu), `title {vi,en,sub}` (`sub` một chuỗi, bản in tự xuống dòng), `template` (tên template đã dùng, để báo cập nhật), `created`/`updated` `{by, at}` (`created` chỉ ghi ở lần Lưu đầu tiên; `by` do trình duyệt ghi nên chỉ để hiển thị, truy vết thật dùng Created/Modified By của SharePoint), `subcontractors` (mảng chuỗi các đơn vị thi công ĐÃ CHỌN; danh sách để chọn là danh mục `subcontractor`), `sections` (thứ tự mảng = thứ tự hiển thị, số La Mã tự đánh liên tục).
  - **Mục (`sections[]`):** `id`, `title {vi,en,short}`, `summary` (luôn có, trống là `""`), tuỳ chọn `columns` + `rows` (bảng; `totals` = giá trị dòng Tổng) và/hoặc `fields` (ô văn bản đơn lẻ `{id,label,value}`: ghi chú HSE, ký duyệt, chỉ đạo); mục ảnh có `type: "photo"` và `rows` = `{id, department, summary, files[]}` (`id` nhóm `g1`, `g2`… = tên thư mục `photos/<id>/`). Cột chỉ có `id` + `label`. Dòng `{id, locked?, cells}`: `id` dòng `r<n>` (= số lớn nhất + 1 trong mục, không bao giờ đổi, `nextRowId`), `locked: true` chỉ ghi ở dòng mẫu không được xoá. Ngày lưu `yyyy-MM-dd`, giờ ISO 8601, số là số JSON.
  - **Template JSON** (`ReportTemplateJson`, `services/ReportTemplateService.ts`): `title`, `docCodePattern`, `sections[]` cùng `id` với báo cáo, thêm cách nhập. Cột: `type` (text/number/date/select/calc), `catalog` (+`catalogBy: "link"`, `from`, `noNone`), `options` (danh sách cố định), `formula`, `readonly`, `key`, `total`, `align`, `width`, `default`, `reset`, `carry`, `tones`. Trường `fields`: `type`, `full`, `span`, `rows`, `lines` (bản in kẻ n dòng, không in nhãn), `placeholder`, `default` (có thể dùng {user} {siteManager} {unit}, chỉ điền khi trống). Mục: `single` (bảng một dòng, ẩn thêm/xoá dòng), `showSummary: false` (ẩn ô Tóm tắt trên form), `layout {columns, sign}` (lưới `fields`; `sign` = in khung chữ ký, tự co theo số người ký), `perRow` (ảnh mỗi hàng), `rows` mặc định (dòng mẫu, `locked`). **Thêm/bớt mục, cột, trường chỉ cần sửa template (đổi ngày trong tên file), không sửa code.** Không có template và đơn vị chưa có báo cáo nào → báo lỗi "Chưa có template" (không có dữ liệu mẫu trong code).
  - **Trường tự động (`utils/reportMeta.ts`):** vị trí = mã đơn vị, không sửa được. `reportNo` tự tăng độc lập theo từng đơn vị theo sổ `ReportRegistry` (xem dưới), không sửa được. `docCode` = `docCodePattern` của template thay `{unit}` và `{reportNo}` (ví dụ `HCM2_1-23_CS-THAC_DCEM-CPZZ-{unit}-CWS-DRP-{reportNo}` → `...-ST01-CWS-DRP-00001`); đã lưu thì giữ nguyên dù template đổi mẫu. Đơn vị thi công chọn nhiều từ danh mục `subcontractor` (chip có nút bỏ, mục "Khác…" thêm đơn vị mới vào danh mục), bắt buộc chọn ít nhất một.
  - **Sổ quản lý báo cáo (SharePoint List `ReportRegistry`, `services/ReportRegistryService.ts`):** mỗi báo cáo (đơn vị × loại × ngày) một dòng, là nguồn của số báo cáo và tab Tổng quan. Cột (tên hiển thị đúng chữ hoa/thường; code tự tra tên nội bộ thật của từng cột qua `/fields` nên không lỗi khi SharePoint đặt tên nội bộ khác như `Unit0`): `Title` (Text, khoá duy nhất `ST01|DailyReport|2026-11-12`, bật "Enforce unique values"), `Unit` (Text), `DocID` (Text), `ReportDate` (Text `yyyy-MM-dd`), `ReportNo` (Number, 0 chữ số thập phân), `Reporter` (Text, email người lập), `Status` (Text: `unsigned` | `done`), `SavedAt` (Text ISO), `SignedAt` (Text ISO), `Counters` (Text, JSON bộ đếm mã tự tăng, vd `{"events.eventid":5,"issues.issueid":4}`; cột không bắt buộc: thiếu thì mã tự tăng chỉ dựa vào báo cáo đang mở). Phải tạo chỉ mục (Indexed) cho `Title`, `Unit`, `ReportDate`, `ReportNo` để truy vấn không vướng ngưỡng 5000 dòng.
    - **Mở báo cáo mới:** số hiển thị = số lớn nhất của đơn vị + 1 (đơn vị chưa có dòng = `00001`); nếu ngày đó đã có dòng trong sổ (vd file bị xoá) thì dùng lại số cũ. Không đọc được sổ thì quay về số của báo cáo ngày trước + 1.
    - **Bấm Lưu:** báo cáo mới đọc lại sổ để lấy số mới nhất rồi mới ghi file, sau đó tạo dòng mới (`Status = unsigned`); báo cáo đã có thì chỉ cập nhật `Reporter` + `SavedAt`. **Upload bản ký:** `Status = done`, `SignedAt`. Ghi sổ lỗi thì báo cảnh báo nhưng không chặn việc lưu file. Mọi thao tác ghi đều sau `checkMaintenance()`.
    - **Phân quyền SharePoint (bảo mật thật):** người dùng báo cáo cần quyền thêm/sửa dòng; nên bỏ quyền xoá của họ (mức quyền tuỳ chỉnh "Contribute không xoá"), bật version history, chỉ admin được xoá/sửa tay. Lưu ý hai lần Lưu cùng lúc của cùng một đơn vị ở hai ngày khác nhau có thể trùng số (hiếm; sổ không chặn).
    - Lịch 14 ngày của một đơn vị và tab Tổng quan cùng đọc sổ này nên luôn khớp nhau, kể cả khi thư mục `data/` bị chuyển đi. Riêng form báo cáo (xem, sửa, tải bản ký) vẫn cần file thật trong `data/`.
    - **Mục dạng bảng / form / ảnh:** bảng (`columns`+`rows`) có nút thêm/xoá dòng (trừ `single`), dòng Tổng khi có cột `total`; `fields` hiển thị dưới bảng (HSE: 5 ô số liệu + 2 ghi chú) hoặc đứng một mình (ký duyệt, chỉ đạo). Ô ngày lưu `yyyy-MM-dd` (ô nhập `type="date"`, bản in `dd/MM/yyyy`). Mã dòng mới `nextRowId`.
    - **Mục ảnh:** người dùng thêm bộ phận bằng nút "Thêm bộ phận" (tên bộ phận + tải ảnh + tóm tắt); mỗi bộ phận một thư mục `photos/<id>/`; lưu vào `rows` của mục khi bấm Lưu. Báo cáo tạo từ ngày trước không sao chép ảnh.
    - **Danh mục dùng chung (template: `catalog`, `catalogBy`, `from`; `utils/reportCatalog.ts`, `ReportDataService.getCatalog`):** mọi dropdown (kể cả danh sách trước đây cố định) lấy từ danh mục theo tên bảng (chữ thường): `activityp6_id`, `element_id`, `reshuman_id`, `resequip_id`, `material_id` và các danh mục dùng chung mọi nhà ga (`Unit = ALL`): `weather`, `event_type`, `event_status`, `supplier`, `equipment_status`, `issue_severity`, `issue_status`, `subcontractor` (các cột này đặt `noNone` = không có mục "---"). Đổi/thêm giá trị chỉ sửa ở nguồn danh mục, không sửa template; báo cáo cũ giữ đúng giá trị đã lưu. **Nguồn:** (1) file `ActivityP6.xlsx` (Activity, chỉ admin ghi) và (2) list `CatalogItems` (phần còn lại, người dùng được thêm bằng "Khác…", ghi vào list dùng chung sau `checkMaintenance()`), gộp bởi `ReportDataService.getCatalog`; không đọc được nguồn nào thì cảnh báo và form vẫn dùng được với giá trị đã lưu. Hướng dẫn cài đặt trên SharePoint: `docs/SHAREPOINT_SETUP.md`.
      - **File `ActivityP6.xlsx` (admin-only):** MỘT sheet (sheet tên `ActivityP6`, hoặc sheet đầu tiên), dòng 1 là tiêu đề, mỗi dòng sau là một Activity với các cột `ActivityID` (mã), `Unit` (đơn vị, ví dụ `ST01`; `ALL` = dùng chung mọi đơn vị) và `Description` (mô tả, KHÔNG bắt buộc; hiện kèm mã trong dropdown dạng "CV-ST01-00001 – Thi công tường vây", giá trị lưu vẫn là mã). Cột tìm theo tên tiêu đề (thứ tự cột tuỳ ý). Web lọc theo đơn vị đang mở để tạo dropdown (`parseCatalogFromWorkbook`): thêm đơn vị mới chỉ thêm dòng, không tạo sheet. Cảnh báo (toast) khi thiếu cột `ActivityID`/`Unit`, trùng mã trong một đơn vị, hoặc đơn vị chưa có Activity nào. Muốn bỏ một mã thì xoá dòng (nên Insert Table Row để thêm dòng). Người dùng KHÔNG thêm được Activity từ giao diện ("Khác…" bị tắt). File được đọc một lần và nhớ 5 phút trong bộ nhớ trang.
      - **List `CatalogItems` (người dùng được thêm, một list cho MỌI nhà ga và MỌI loại danh mục):** cột `Title` (khoá duy nhất, bật Enforce unique values), `Unit`, `Catalog` (tên bảng chữ thường: `element_id`, `reshuman_id`, `resequip_id`, `material_id`, bảng mới chỉ cần dùng tên mới), `Code` (mã; để trống nếu mục chỉ có tên, vd vật tư người dùng thêm theo tên), `Link` (giá trị ràng buộc: Element → mã Activity; Resource → Loại/Thiết bị; Vật tư → tên). Thêm nhà ga hoặc loại danh mục mới chỉ thêm dòng, không tạo list mới. **Mục dùng chung mọi nhà ga: `Unit = ALL`** (hằng `CATALOG_SHARED_UNIT`; vd mã nhân lực `RES-HUM-…` và thiết bị `RES-EQP-…` không gắn với nhà ga): nhập MỘT dòng, mọi nhà ga đều đọc được (truy vấn `Unit = <nhà ga> hoặc ALL`), sửa một chỗ là tất cả nhà ga cùng đổi; nhà ga mới có sẵn các mục này. Chống trùng kiểm tra trên cả `ALL` lẫn nhà ga đang mở. Giá trị người dùng thêm bằng "Khác…" luôn lưu riêng cho nhà ga của họ (`Unit` = nhà ga); admin muốn dùng chung thì đổi `Unit` thành `ALL` và `Title` thành `ALL|<catalog>|…` cho khớp. Khoá `Title` của dòng `ALL` khác khoá của dòng riêng nhà ga nên SharePoint không chặn trùng giữa hai loại, chỉ giao diện chặn. **Quy ước khi cập nhật dữ liệu mẫu `CatalogItems_list.xlsx` (và khi nhập list):** luôn sắp xếp các dòng theo `Unit`: nhóm `ALL` lên đầu, rồi lần lượt ST01, ST02, ST04, ST05… (nhà ga mới chèn theo thứ tự tên); trong mỗi nhóm giữ thứ tự bảng. Khoá `Title` phải khớp `Unit|catalog|code|<mã chuẩn hoá>` (hoặc `…|link|…` khi chỉ có tên) và không trùng nhau. Index `Title`, `Unit`. Phân quyền SharePoint: người dùng báo cáo chỉ THÊM (không xoá/sửa), admin toàn quyền (sửa chính tả, gán mã, xoá mục sai); bật version history. Code đọc/ghi: `ReportCatalogService` (tự tra tên nội bộ cột như `ReportRegistry`). Không đọc được list thì cảnh báo và vẫn dùng phần từ file.
      - **Chống trùng (hai lớp):** (1) giao diện: giá trị người dùng nhập bằng "Khác…" được chuẩn hoá (`normKey`: bỏ khoảng trắng thừa, không phân biệt hoa thường và dấu tiếng Việt: "thep  d32" = "Thép D32") rồi so với mã (hoặc với giá trị ràng buộc nếu cột chọn theo `.Cột`); đã có thì KHÔNG tạo mới mà chọn mục sẵn có và báo; (2) SharePoint: `Title` = `<Unit>|<catalog>|code|<mã chuẩn hoá>` (hoặc `…|link|<giá trị chuẩn hoá>` khi mục chỉ có tên) duy nhất; hai người thêm cùng lúc thì dòng sau bị từ chối và giao diện tự chọn mục đã có. Hạn chế: mục có mã và mục chỉ-có-tên của cùng một vật tư có khoá khác nhau nên chỉ lớp (1) chặn; admin gộp tay khi cần.
      - **Mục mới có liên kết (hai bước):** chọn "Khác…" ở cột mã có ô phụ thuộc (`from=`) thì sau khi nhập mã, hộp thoại thứ hai hỏi giá trị liên kết (Element → Activity chọn từ `ActivityP6_ID`; Mã nhân lực → Loại; Mã thiết bị → Thiết bị), KHÔNG bắt buộc (bỏ trống rồi bấm Thêm được; Huỷ = không thêm gì). Mã và liên kết ghi vào `CatalogItems` trong MỘT lần thêm (không cần quyền sửa dòng), nên mọi người sau đó thấy liên kết tự điền và khoá như các mã cũ. Bỏ trống liên kết thì ô phụ thuộc hiện `---`/trống chọn tay (chỉ trong báo cáo đó), admin điền liên kết sau trong list. Mã đã có (chống trùng) thì không hỏi. Vật tư (chọn theo tên) không hỏi liên kết.
      - **Quyền thêm:** bảng `activityp6_id` KHÔNG BAO GIỜ thêm được từ giao diện (hằng `ADMIN_ONLY_TABLES` trong `reportCatalog.ts`; "Khác…" bị tắt); bảng đến từ file cũng không; bảng còn lại được thêm (ghi vào `CatalogItems` sau `checkMaintenance()`, mọi người và mọi báo cáo thấy ngay). Ghi list lỗi thì giá trị chỉ lưu trong spec của báo cáo đó (`select:@bảng/giá-trị`) kèm cảnh báo.
      - **Mã `---`:** luôn có sẵn trong dropdown làm giá trị "chưa chọn"; không ghi vào file. Mã không có ràng buộc thì để trống ô cột 2.
      - **Bỏ mã:** admin xoá dòng khỏi file hoặc list (không còn `inactive`). Báo cáo cũ vẫn hiện đúng giá trị đã lưu (nằm trong file báo cáo, không tra lại danh mục) và lưu lại không bị đổi.
      - **`from` (id cột nguồn):** chọn mã ở cột nguồn thì ô này tự điền giá trị ràng buộc của mã và bị khoá (Element → Activity; ResHuman → Loại; ResEquip → Thiết bị; Material → Vật tư). Mã chưa có ràng buộc: ô select về `---`, ô text để trống, vẫn nhập/chọn tay được; chọn `---` ở cột mã thì giữ nguyên ô đang có. Chỉ điền khi người dùng đổi mã, mở báo cáo cũ không tự điền lại.
      - **Chọn theo giá trị ràng buộc (`catalogBy: "link"`):** mục Vật tư: cột `material` (`catalog: "Material_ID"`, `catalogBy: "link"`) liệt kê tên vật tư và cột `matId` (`catalog: "Material_ID"`, `readonly`, `from: "material"`) tự điền mã theo vật tư đã chọn, không chọn tay; mục Nhân lực làm tương tự với `category` / `resId` (`ResHuman_ID`) và mục Thiết bị với `equipment` / `resId` (`ResEquip_ID`). Giá trị mới nhập bằng "Khác…" thì mã về `---` (mục mới ghi vào danh mục chỉ có tên, admin cấp mã sau).
      - Cột dùng danh mục (trừ bảng admin-only) có mục "Khác…" ở cuối: giá trị mới ghi vào danh mục dùng chung (xem trên); ô ràng buộc (`from`) của dòng đó về `---` (select) hoặc để trống (text).
      - `readonly` khoá cột ở mọi dòng, `key` chỉ khoá ở dòng `locked` (cột select danh mục `key` cũng chỉ khoá ở dòng khoá). Dropdown hiển thị đúng mã. Danh mục `'sharepoint'` được đọc một lần và nhớ 5 phút trong bộ nhớ trang; không đọc được thì form vẫn dùng được (chỉ còn `---` và giá trị cũ) kèm thông báo.
      - **Mã tự tăng (Event ID, Issue ID):** không nằm trong danh mục; cột `text;readonly;default=EVT-{unit}-{n5}` / `default=ISS-{unit}-{n5}`: khi thêm dòng mới, `{n5}` = số lớn nhất đang dùng + 1 đệm 5 chữ số (`EVT-ST01-00005`), người dùng không sửa được. "Số lớn nhất đang dùng" = lớn hơn cả các dòng trong báo cáo lẫn **bộ đếm của đơn vị trong sổ `ReportRegistry`** (cột `Counters`, khoá `<mục>.<cột>` chữ thường như `events.eventid`): mỗi lần Lưu ghi bộ đếm cộng dồn vào dòng báo cáo, khi mở báo cáo đọc bộ đếm từ dòng có `ReportNo` cao nhất của đơn vị (`ReportRegistryService.getCounters`), nên số liên tục giữa các ngày và không dùng lại số của dòng đã xoá. Khi mở báo cáo, các ô mã tự tăng còn trống hoặc còn nguyên mẫu `{n5}` (dòng tạo bởi bản cũ) được điền lại (`fillAutoIds`). Giới hạn như `ReportNo`: hai người Lưu cùng lúc ở hai ngày khác nhau của cùng đơn vị có thể trùng số (hiếm).
      - Hạn chế đã biết: dropdown là `<select>` thường (chưa có ô gõ lọc).
    - **Dropdown mở rộng:** mọi ô `select` có mục cuối "Khác…" (`core/components/CreatableSelect`) mở hộp thoại nhập giá trị mới; cột dùng danh mục thì ghi vào danh mục dùng chung (chống trùng như trên), cột dùng `options` cố định thì thêm vào danh sách của cột trong báo cáo đang mở. Ký tự `/` và `;` trong giá trị mới đổi thành `-`.
    - **Cột mã (ID)** lấy từ danh mục là `select` có `default: "---"`. Cột `select` có `key` hoặc `readonly` thì dropdown bị vô hiệu hoá (dòng thêm tay giữ giá trị mặc định).
    - **Màu trạng thái tự động:** ô `select` tự tô màu theo ý nghĩa giá trị, không cần khai báo trong template (`AUTO_TONES` trong `reportTable.ts`): Xong/Active/Hoạt động/Đạt = xanh lá, Đóng/Huỷ = xám, Đang xử lý = xanh dương, Mở/Pending/Chờ/Bảo trì = vàng, Hỏng/Chậm/Quá hạn = đỏ; mức độ Cao = đỏ, Trung bình = vàng, Thấp = xanh lá. `tones` trong template vẫn được ưu tiên.
    - **Bảng rộng** cuộn ngang trong `.table-responsive` (thanh cuộn luôn hiện); mỗi cột có bề rộng tối thiểu theo `width` (`columnMinPx`), không bị bóp chữ.
    - **Tiêu đề mục** có thể dùng `{tomorrow}` = ngày hôm sau của báo cáo (`dd/MM/yyyy`), ví dụ `KẾ HOẠCH NGÀY MAI – {tomorrow}`.
    - **Luỹ kế (mục Nhân lực, Khối lượng):** cột `prevCum` (Luỹ kế hôm qua, `carry: "daily+prevCum"`), `daily` (`reset`), tổng = `calc` `daily+prevCum`, tiến độ = `calc` `round((daily+prevCum)/planned*100)`.
  - **Báo cáo tuần (placeholder):** `tasks` (cây công việc) lưu thẳng trong `.data.json`; chưa có template JSON tuần nên chỉ là khung trống.
* **⚠️ WeeklyReport hiện chỉ là placeholder:** tile `WeeklyReport`, file template tuần tạm và code module tuần (`ReportTaskTable`, sheet `07_Tasks`, tính tỉ lệ cha/con) chỉ dùng để hiển thị UI, chưa phải tính năng hoàn chỉnh. Chủ dự án sẽ làm phần báo cáo tuần sau. Không test, sửa logic hay mở rộng phần riêng của weekly (task tree, template/data tuần) và không coi đó là việc tồn đọng cho tới khi được yêu cầu; chỉ giữ UI dùng chung (nút, icon, token) nhất quán.
* **Locked Predefined Template Rows (`locked`):**
  - Dòng có `locked: true` (dòng mẫu của template; mặc định các dòng của Khối lượng) không xoá được (biểu tượng khoá thay nút xoá) và các cột `key` (Activity/cấu kiện, Luỹ kế hôm qua, Tổng phải làm…) chỉ đọc ở dòng đó. Người dùng vẫn sửa số lượng/giờ và thêm/xoá dòng tự thêm.
* **Smart Data Loading & Previous-Day Prefill (`ReportDataService.loadReport`):**
  1. Nếu có `data/<Unit>/<DocID>_<yyyyMMdd>/<DocID>_<yyyyMMdd>.data.json` thì đọc đúng file đó (`isExisting = true`), ghép với template mới nhất theo `id`.
  2. Chưa có: tìm báo cáo ngày liền trước (`dateStr - 1`) hoặc quét `data/<Unit>/` lấy ngày trước gần nhất (`getPreviousReportJson`), dùng làm nền (`copiedFromDate`, `isExisting = false`): cột `carry`/`reset` được áp dụng, ảnh không sao chép, số báo cáo theo sổ `ReportRegistry` (hoặc ngày trước + 1).
  3. Đơn vị chưa có báo cáo nào: tạo từ template mới nhất trong `_template/` (không có template thì báo lỗi).
  - `parseReportJson`/`parseTemplateJson` báo lỗi dễ hiểu khi file hỏng hoặc `schema` mới hơn bản hỗ trợ.
* **A4 Print Preview, PDF và HTML độc lập (`ReportPrintDocument.tsx`, `ReportPrintPreviewModal.tsx`, `utils/reportPrintCss.ts`, `utils/reportHtml.ts`):**
  - `ReportPrintDocument` render một trang A4 song ngữ (header THACO/CREC, bảng thông tin Mã tài liệu / Vị trí / Số báo cáo / Đơn vị thi công / Ngày, các mục theo thứ tự trong file, chân trang), không dùng hook/DOM của trình duyệt nên dùng được cho cả cửa sổ xem trước/in của web (iframe, `PRINT_CSS`) lẫn render tĩnh. Mẫu giấy chuẩn `Mau_BCN_ST06_20261025.pdf`; ngắt trang: tiêu đề mục + hàng tiêu đề + 2 dòng đầu luôn đi cùng nhau, 2 dòng cuối không bị bỏ lẻ; ảnh không bị ngắt giữa hai trang; `handlePrint` đợi mọi ảnh tải xong (tối đa 20s) rồi mới gọi `window.print()`; trên điện thoại khung xem trước giữ bề rộng A4 tối thiểu và cuộn ngang.
  - **File HTML độc lập:** mỗi lần Lưu, `buildStandaloneHtml` (`createRoot` + `flushSync` vào div rời rồi lấy `innerHTML`, import tĩnh, không tải chunk riêng; KHÔNG dùng `react-dom/server` vì trong gói SPFx lỗi `getCurrentStack`) render cùng `ReportPrintDocument` thành `<DocID>_<yyyyMMdd>.html` = nội dung bản in + CSS in + logo (`resource/`) và ảnh hiện trường nhúng base64 (ảnh cạnh dài > 1280px được thu nhỏ bằng `utils/reportImage.ts`) + toàn bộ `.data.json` trong `<script type="application/json" id="report-data">`; mở offline, nút "In / Lưu thành PDF (A4)" (tên PDF mặc định = Mã tài liệu). Ghi HTML lỗi thì cảnh báo, không chặn việc lưu `.data.json`. Màu hex của CSS in/khung giấy là ngoại lệ có chủ đích (`HEX_ALLOWED_TSX` trong `scripts/check-styles.js`).
* **Thông báo cập nhật template (`ReportEntryView.tsx`, class `report-tpl-notice`):** khi báo cáo (đã lưu hoặc kế thừa ngày trước) có `template` cũ hơn bản mới nhất trong `_template/`, form hiện khung vàng "Cần cập nhật template mới" với hai ô phiên bản (HIỆN TẠI gạch ngang, MỚI NHẤT xanh lá) và nút "Apply template mới để tiếp tục"; nút "Xem dữ liệu cũ (không đổi template)" chỉ có khi báo cáo đã lưu. Trong lúc này ẩn các nút Xem bản in / Làm mới / Lưu báo cáo và nhãn "Chưa lưu". Apply gọi `checkMaintenance()` trước (bảo trì thì không áp dụng) rồi `applyTemplate` GHÉP THEO `id`: giữ dữ liệu cũ, thêm mục/cột/trường mới (ô trống hoặc mặc định), bỏ cột/mục đã xoá khỏi template, thêm dòng mẫu `locked` mới, cập nhật nhãn và cách nhập. **Phiên bản so theo NGÀY trong tên file (`template_DRP_<Unit>_<yyyyMMdd>.json`)**: mỗi lần sửa nội dung template phải đổi ngày trong tên file sang ngày mới hơn, nếu không báo cáo đã lưu hoặc kế thừa từ ngày trước vẫn giữ bản cũ và không hiện thông báo cập nhật. Giữ `id` cột/mục/trường ổn định giữa các bản template (đổi `id` = cột mới, dữ liệu cột cũ bị bỏ).
* **Save & Overwrite (`ReportDataService.saveReport`):**
  - `checkMaintenance()` → đảm bảo `data/<Unit>` và `data/<Unit>/<DocID>_<yyyyMMdd>/` tồn tại → ghi `.data.json` (kiểm tra ETag) → ghi `.html` → ghi sổ `ReportRegistry`. `updated` = người đang lưu; `created` ghi ở lần đầu và giữ nguyên.
* **SheetJS Lazy-Loading:** chỉ còn dùng cho `ActivityP6.xlsx` và xuất Excel ở tab Tổng quan; nạp bằng `import(/* webpackChunkName: 'xlsx' */ 'xlsx')` để module khác không tốn dung lượng.
* **Chạy thử cục bộ không cần SharePoint (`dev/`, ngoài `src/` nên không vào gói web part):** `npm run dev:report` (= `node dev/build.js`) biên dịch `src/` + `dev/` (tsc → `temp/dev/js`), nạp `data_test/sharepoint_sample` (template JSON, `ActivityP6.xlsx`, logo, báo cáo mẫu; list `CatalogItems`, `Report_Role` từ các file `*_list.xlsx`; thêm `?as=<email>` vào URL để mở như người dùng thường, `?norole=1` để thử lỗi thiếu list) làm thư viện giả lập trong bộ nhớ (`dev/LocalSharePointClient.ts` thay `core/services/SharePointClient`, `@microsoft/sp-http` thay bằng stub) và gộp bằng webpack ra `temp/dev/index.html` (mở bằng trình duyệt, đủ landing/lịch/form/in/lưu/upload bản ký; dữ liệu mất khi tải lại trang). Chạy `npm run test:report` (= build + `node --test dev/tests`) để kiểm thử chuyển đổi JSON ↔ mô hình ↔ template (dùng bản biên dịch trong `temp/`, KHÔNG đụng `dist/`). Không chạy `npm run build`/`heft test --clean` khi `npm run start` đang chạy.

---

## 6. Developer Runbook & Common Commands

### Running Locally (Development Mode)

```bash
cd /Volumes/DATA/THADICO_local/REPO_DQM/DQM_Check_In_App
npm run start
```

* Local assets serve at `https://localhost:4321`.
* If testing on a new browser/machine, trust the SSL certificate once by opening:
  `https://localhost:4321/temp/build/manifests.js` $\rightarrow$ Click **Advanced** $\rightarrow$ **Proceed to localhost (unsafe)**.
* Test Workbench URLs:
  - **Khuyến nghị (Live-Reload mượt mà, không bị popup cảnh báo):**
    `https://ytf84.sharepoint.com/sites/metro-checkin-dev/_layouts/15/workbench.aspx?reset=true`
  - **Dự phòng (Explicit Manifest - yêu cầu bấm "Load debug scripts"):**
    `https://ytf84.sharepoint.com/sites/metro-checkin-dev/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true`

### Building for Production Deployment

```bash
cd /Volumes/DATA/THADICO_local/REPO_DQM/DQM_Check_In_App
npm run build
```

* Runs Heft test compilation and packages client-side assets directly into:
  📁 `sharepoint/solution/workspace_portal.sppkg`

### Deploying to SharePoint

1. Open the SharePoint App Catalog (e.g., `https://ytf84.sharepoint.com/sites/appcatalog` or tenant app catalog).
2. Navigate to **Apps for SharePoint**.
3. Upload `sharepoint/solution/workspace_portal.sppkg`.
4. Check **"Enable this app and add it to all sites"** (or add to specific sites) and click **Deploy**.
5. On any page in your SharePoint site, click **Edit** $\rightarrow$ Click **`+`** $\rightarrow$ Search and select **Workspace Portal** $\rightarrow$ **Republish**.

---

## 7. Guidelines for AI Agents Working on This Codebase

1. **Do NOT import `@microsoft/power-apps`:** This project is deliberately detached from Power Apps to eliminate licensing costs.
2. **UI & Style System:** mọi thay đổi giao diện phải tuân thủ **mục 8** (token, nút, icon, ô nhập, card, header, quy tắc đặt tên class, khung module). Chạy `npm run lint:styles` trước khi báo hoàn thành.
3. **Git Commit & Push Policy:** DO NOT automatically run `git commit` or `git push`. Only commit and push when the user explicitly requests or approves it.
4. **Responsive & Cross-Device UI Rules (BẮT BUỘC khi thêm/sửa giao diện):**
   - **Kiểm tra mọi cỡ màn hình:** desktop rộng (≥1440px), laptop (1024–1440px), **iPad mini / tablet dọc (768px)**, tablet ngang (1024px), điện thoại (≤480px). Mỗi thay đổi UI phải được rà lại ở các cỡ này, không chỉ ở desktop.
   - **Breakpoints chuẩn:** `≤1024px` (tablet), `≤768px` (iPad mini / tablet nhỏ), `≤640px` và `≤480px` (điện thoại). Đặt media query ở cuối file CSS chứa thành phần đó (`core/styles/components/*.css`, `core/styles/layout/*.css` hoặc `modules/<name>/<name>.css`). Token không dùng được trong `@media` nên ghi số trực tiếp theo đúng 4 mốc này. Lưu ý các quy tắc `@media` cũ có thể ghi đè; khi thêm quy tắc mới phải kiểm tra xung đột (đặc biệt tại đúng 768px).
   - **Không để nội dung bị cắt hoặc mất:** lưới dùng `repeat(N, minmax(0, 1fr))`, phần tử con flex/grid có `min-width: 0`; không dùng `white-space: nowrap` trên khối chứa văn bản dài nếu không có `text-overflow: ellipsis` hoặc cho phép xuống dòng. Không dùng `overflow: hidden` ở thẻ cha làm cắt cột bên phải.
   - **Không bóp chữ:** tiêu đề, badge, nút không được bị bẻ từng chữ. Nếu thiếu chỗ thì cho nhóm điều khiển xuống hàng (`flex-wrap: wrap`) thay vì ép `nowrap` toàn bộ hàng.
   - **Bố cục logic theo cỡ màn hình:** desktop/tablet: tiêu đề bên trái, điều khiển (sắp xếp, làm mới, view tabs) bên phải cùng hàng; điện thoại: xếp dọc, điều khiển rộng hết chiều ngang, dùng 1 cột thẻ. Tablet: 2 cột thẻ. Giới hạn bề rộng nội dung (`max-width: 1440px`, căn giữa) trên màn hình rộng.
   - **Chuyên nghiệp & nhất quán:** dùng token trong `core/styles/tokens.css`, khoảng cách theo `--space-*`, vùng chạm tối thiểu ~36–44px trên cảm ứng, căn hàng đều giữa các nhóm điều khiển, bảng dữ liệu cuộn ngang trong `.table-responsive` thay vì tràn trang.
   - **Xác minh trước khi báo hoàn thành:** build sạch (`npm run build`, 0 cảnh báo lint) và, nếu có thể, kiểm tra trực quan ở 1440px, 1024px, 768px, 390px. Nếu không kiểm tra được trực quan, phải nói rõ với người dùng là chưa xem lại trên thiết bị.
5. **Chế độ bảo trì (Shutdown Server) chặn MỌI module, kể cả module làm sau này (BẮT BUỘC):** khi admin bật Admin Config → Shutdown Server, người dùng không phải admin không được dùng bất kỳ module nào (không lưu, tải lên, xoá, tạo mới, Back/Home). `PortalShell` đã ẩn module đang mở và hiện màn hình bảo trì (đọc lại cấu hình mỗi 15 giây và khi quay lại tab); không được đi vòng qua cổng này. Mỗi module mới phải gọi `await checkMaintenance()` (có trong `ModuleProps`) ngay trước mọi thao tác ghi lên SharePoint và dừng nếu trả về `true`. Admin không bị chặn.
6. **Standard Recipe for Adding a New Module (e.g. `tasks`):**
   - Create `src/modules/<name>/` with `<Name>Module.tsx` (receives `ModuleProps`; import `./<name>.css` ở đây), `views/`, `components/`, `services/<Name>Service.ts`, `types/<name>.types.ts`, `constants/`, `utils/`, and `index.ts` exporting a `PortalModule`.
   - Only files belonging to a module carry that module's name (e.g. `Report*`). Shared code in `core/` and `app/` must not contain module names or company names.
   - Register the module in `src/app/moduleRegistry.tsx`; `HomeView` lists it automatically and `PortalShell` opens it on click.
   - SharePoint access: build module services on `core/services/SharePointClient`; list names live in the module's service file.
   - Mọi thao tác ghi lên SharePoint của module (lưu, tải/xoá file…) phải gọi `await checkMaintenance()` (có trong `ModuleProps`) ngay trước khi ghi và dừng nếu trả về `true` (portal đang bảo trì, người dùng không phải admin).
   - Dependencies are one-way: `modules/*` → `core/*`; `core/` never imports `modules/`; modules never import each other.
   - Storage keys use the prefix `portal_<module>_...`. Style: xem mục 8 (CSS riêng của module chỉ ở `modules/<name>/<name>.css`, class có tiền tố `<name>-`).

7. **Nơi lưu quy tắc dự án:** mọi quy tắc, quy ước, quyết định kiến trúc và lưu ý lâu dài của dự án CHỈ được ghi vào file `AGENTS.md` này (nằm trong repo để dùng được trên mọi máy). Không lưu vào bộ nhớ/cấu hình cục bộ của công cụ AI (ví dụ thư mục `~/.claude`).

8. **Bảo mật khi thiết kế module mới (BẮT BUỘC):**
   - **Kiểm tra quyền ở giao diện chỉ là tiện ích, KHÔNG phải bảo mật.** Web part chạy hoàn toàn trên trình duyệt; mọi thứ như `isAdmin`, vai trò, ẩn/hiện nút đều có thể bị người dùng biết kỹ thuật thay đổi. Bảo mật thật phải dựa vào **phân quyền SharePoint** (list, thư viện, thư mục): mỗi module phải ghi rõ trong AGENTS.md những list/thư viện/thư mục nó đọc-ghi và quyền SharePoint cần cấu hình (ai chỉ đọc, ai được ghi). Thiết kế sao cho dù giao diện bị vượt qua, người dùng vẫn chỉ làm được những gì quyền SharePoint của họ cho phép.
   - **Danh tính và vai trò luôn đọc lại từ nguồn thật** (list `UserProfiles` theo tài khoản M365 đăng nhập) mỗi lần mở web part. KHÔNG lưu quyền/vai trò vào `sessionStorage`/`localStorage` rồi tin lại; không tin dữ liệu do trình duyệt cung cấp.
   - **Chức năng chỉ dành cho admin** (đặt `adminOnly` cho module hoặc kiểm tra trong view) phải dựa trên **admin thật** của `PortalShell` (tài khoản đăng nhập), không dựa trên hồ sơ đang "xem với vai trò khác". Không thêm cơ chế chuyển vai trò/đóng vai nào cho người không phải admin thật.
   - **Dữ liệu nhạy cảm không đưa vào nơi mọi người đọc được** (list `UserProfiles` đang lộ email và vai trò cho ai có quyền đọc list). Cấu hình ghi bằng quyền cao (vd `SiteAssets/portal_config.json`) phải giới hạn quyền ghi cho admin trên SharePoint.
   - Mọi thao tác ghi tuân thủ điều 5 (chế độ bảo trì: `checkMaintenance()` trước khi ghi).
   - Không ghi vào log/console thông tin cá nhân hay token; không gửi dữ liệu người dùng ra dịch vụ bên ngoài SharePoint/M365.

---

## 8. UI & Style System (BẮT BUỘC cho mọi module hiện tại và sau này)

Mục tiêu: mọi module trông như MỘT sản phẩm. Module chỉ ghép các thành phần chung của `core`; CSS riêng của module chỉ lo bố cục/nghiệp vụ của chính nó.

### 8.1 Kiến trúc file

```text
src/core/styles/
├── index.ts            # điểm nạp DUY NHẤT (thứ tự: tokens → base → components → layout)
├── tokens.css          # CHỈ biến CSS (màu, chữ, spacing, radius, shadow, --control-h…) – nơi duy nhất được có màu hex
├── base.css            # nền chung: font, box-sizing, reset nút, focus, tiện ích (.text-center…), .icon/.spin
├── components/         # button, form (.input/.textarea/.field), badge, card (.card/.section-card), table,
│                       #   chips (.chip/.search), page-header, feedback (toast, spinner, .state), dialog
└── layout/             # app-bar, landing (khung portal)
src/core/components/    # React tương ứng: Icon, PageHeader, FilterChips, SearchInput, LoadingState/EmptyState…
src/modules/<name>/<name>.css   # CHỈ style riêng của module (nạp trong <Name>Module.tsx, SAU core)
src/modules/admin/              # Admin Config (chỉ admin thấy trên trang chủ): menu UI Kit (tài liệu sống của mọi thành phần dùng chung) + Shutdown Server
```

### 8.2 Nguyên tắc cốt lõi
1. **Token trước:** màu/khoảng cách/radius/shadow/cỡ chữ chỉ dùng `var(--…)`. Hex chỉ nằm trong `tokens.css` (ngoại lệ duy nhất: CSS in A4 trong `ReportPrintPreviewModal.tsx` – iframe riêng, màu cố định theo mẫu giấy: tiêu đề mục `#014a94`, nền hàng tiêu đề `#dbe6f1`, `DAILY REPORT` đỏ `#c00000`).
2. **Không `!important`.** Core scope bằng `:where(.app-layout) .btn` (specificity 0) nên CSS module luôn ghi đè được mà không cần ép.
3. **Phạm vi:** `PortalShell` bọc MỌI module trong `.app-layout > .main-content` (token + base + lề + max-width 1440px). Module KHÔNG tự bọc lại, KHÔNG tự đặt `max-width`/`margin auto`/`box-sizing`. Token nằm trên `.app-layout`, không rò ra trang SharePoint.
4. **Đặt tên:** class của core là tên chung (`.btn`, `.card`, `.input`…); class của module LUÔN có tiền tố module (`report-…`, `tasks-…`; bản in A4 dùng `bcn-…`). Core không bao giờ chứa tên module. Hai module không dùng chung class của nhau.
5. **Quy tắc “dùng lần thứ hai”:** một mẫu UI xuất hiện ở ≥2 chỗ (hoặc chắc chắn module sau cũng cần) → đưa lên `core` (CSS trong `core/styles/components/` + React trong `core/components/`), không chép sang module khác.
6. **Một tông xanh:** accent `--color-accent`; không thêm màu nhấn khác. Trạng thái dùng token success/warning/danger.
7. **Icon:** chỉ `<Icon name="…" size={16|20|24} />` (Fluent Regular; thêm icon mới vào bảng `ICONS` trong `Icon.tsx`). Không `<svg>` tay, không emoji/ký tự (×, ✓, 🔒) làm icon. Ngoại lệ: hình thương hiệu của ô module (`modules/<name>/index.ts`).
8. **Vùng chạm:** nút/ô nhập dùng `--control-h` (40px; tự 44px khi ≤768px hoặc màn cảm ứng) và `--control-h-sm` (32→36px). Không đặt `height` cứng cho nút/ô nhập.

### 8.3 Danh mục thành phần dùng chung (dùng ĐÚNG, đừng tự chế)

**Trước khi viết bất kỳ UI nào, theo thứ tự:** (1) đọc bảng dưới; (2) mở **UI Kit** (`modules/admin/views/StyleGuideView.tsx`; admin vào Admin Config → "UI Kit (Style Guide)") để xem mọi thành phần và biến thể; (3) đọc JSDoc `@example` của component trong `core/components/*.tsx` và đầu file `core/styles/components/*.css`; (4) chỉ khi KHÔNG có thành phần phù hợp mới thêm vào `core` (kèm cập nhật bảng này và UI Kit), rồi mới dùng.

| Cần | Dùng | Ghi chú |
|---|---|---|
| Nút | `.btn` + `btn--primary` (hành động chính: Lưu, Tải template) · `btn--secondary` (mọi nút còn lại: Trang chủ, Quay lại, Làm mới, In) · `btn--ghost` / `btn--ghost-danger` (nút nhỏ trong bảng) · `btn--danger` (chỉ phá huỷ) · `btn--light` (trên banner xanh đậm) | cỡ `btn--sm`, `btn--icon` (chỉ icon, vuông), `btn--block` (rộng hết hàng); KHÔNG tạo nút riêng kiểu `report-btn-*` |
| Ô nhập | `.input`, `.textarea`; biến thể `input--cell` (trong bảng), `input--readonly`, `input--locked`, `input--danger`, `input--warning`, `input--success`, `input--info`, `input--muted` | ô nhập đã `border-box`, cao `--control-h` |
| Nhãn + ô nhập | `.field` > `.field__label` (hoặc `field__label--bilingual` với `field__label-main`/`field__label-sub`) | |
| Đầu trang view | `<PageHeader leading title badge meta actions divider />` | desktop: tiêu đề trái, actions phải; ≤768px xếp dọc |
| Lọc / tìm | `<FilterChips />`, `<SearchInput />` | |
| Lưới ô chức năng (trang chủ, menu module) | `<TileGrid items={[{ id, title, icon, onClick }]} />` | nút xanh đặc: icon + tên + mũi tên; `icon` là `<Icon … />` hoặc `<svg>` thương hiệu |
| Đang tải / rỗng / không quyền | `<LoadingState />`, `<EmptyState icon title message hint action tone narrow />` | `tone="danger"` → `state--danger`; `narrow` → `state--narrow` (khung hẹp giữa trang) |
| Khối nội dung | `.card`; khối có banner xanh đậm: `.section-card` > `__banner` (`__title`, `__title-sub`) + `__body` | |
| Nhãn trạng thái | `.badge` + `badge--info` · `badge--success` · `badge--warning` · `badge--danger` | |
| Bảng nhập liệu | `<div class="table-responsive"><table class="portal-table …">` | cuộn ngang, không tràn trang |
| Dropdown có thể thêm giá trị | `<CreatableSelect value options onChange={(value, isNew)=>…} placeholder wrapperClassName allowOther optionLabels />` (khung `.select` + icon `select__icon`, mũi tên gốc đã bỏ) | mục cuối "Khác…" mở hộp thoại; nơi gọi lưu giá trị mới (`isNew`) vào danh sách; `allowOther={false}` = danh sách cố định (không có "Khác…"); `optionLabels` = nhãn hiển thị khác giá trị |
| Hộp thoại nhập một giá trị | `<PromptDialog title label onSubmit onCancel options allowEmpty />` (`.dialog-backdrop` > `.dialog` > `dialog__title`, `dialog__body`, `dialog__actions`) | Enter xác nhận, Esc huỷ; `options` = chọn trong danh sách thay vì gõ chữ; `allowEmpty` = cho xác nhận khi để trống (thông tin không bắt buộc) |
| Thông báo | `<ToastAlert />` | |

### 8.4 Khung module (skeleton)
```tsx
// modules/tasks/TasksModule.tsx
import './tasks.css';                       // CSS riêng, nạp sau core
export const TasksModule: React.FC<ModuleProps> = ({ onGoHome }) => (
  <div className="tasks-module-root">       {/* chỉ flex column + gap; lề/max-width do PortalShell lo */}
    <PageHeader leading={<button className="btn btn--secondary" onClick={onGoHome}><Icon name="home" /><span>Trang Chủ</span></button>}
                title="Quản lý Công việc" actions={…} />
    …
  </div>
);
```

### 8.5 Checklist khi làm module mới
1. Tạo `modules/<name>/<name>.css` (tiền tố class `<name>-`), import trong `<Name>Module.tsx`.
2. Dùng `PageHeader`, `FilterChips`, `SearchInput`, `LoadingState/EmptyState`, `.btn`, `.input`, `.card/.section-card`, `<Icon>` – không viết lại.
3. Cần mẫu UI mới dùng chung? Thêm vào `core` (CSS + React) và ghi vào bảng 8.3, rồi mới dùng ở module.
4. Responsive đủ 4 mốc: `≤1024`, `≤768`, `≤640`, `≤480` (xem mục 7, điều 4).
5. Chạy `npm run lint:styles` – kiểm: không `!important`; không hex ngoài token; class module có tiền tố; **class module trùng ý thành phần có sẵn** (tên chứa `btn/card/badge/chip/header/input/search/loading/empty/toast…` → báo lỗi và gợi ý dùng thành phần của core; ngoại lệ có chủ đích phải ghi vào `REINVENT_ALLOW` kèm lý do); core không chứa tên module/không import `modules/`; không `<svg>` tay; **mọi component, biến thể (`block--x`) và block của core phải được ghi trong bảng 8.3** (thêm thành phần mà quên cập nhật tài liệu → lint lỗi). Chạy kèm `npx tsc --noEmit -p tsconfig.json` và `npx eslint src --ext .ts,.tsx`.
6. Nếu thêm thành phần/biến thể mới ở core: thêm một ví dụ vào `modules/admin/views/StyleGuideView.tsx`.
7. Xem trên workbench ở 1440 / 1024 / 768 / 390px. KHÔNG chạy `npm run build` khi `npm run start` đang chạy (build xoá `dist/` đang phục vụ nên workbench rơi về bản cũ).

### 8.6 Nợ kỹ thuật đã biết (đừng nhân rộng)
* `report.css` còn một số class chưa có tiền tố thuộc phần Weekly placeholder (`kpi-box`, `id-tag`, `rate-input`… – danh sách trong `scripts/check-styles.js` → `LEGACY_CLASSES`). Dọn (đổi sang `report-…`) khi làm báo cáo tuần; KHÔNG thêm class mới không tiền tố.

---

## 9. Việc có thể làm tiếp (backlog, CHƯA làm)

Các ý định đã bàn nhưng chưa triển khai; chỉ làm khi chủ dự án yêu cầu. Ghi ở đây để phiên làm việc sau nắm được bối cảnh. (Upload bản ký trong danh sách Tổng quan: đã chốt giữ như hiện tại, nút "Upload bản ký" chỉ mở báo cáo của ngày, không làm upload thẳng.)

**Bàn giao giữa các phiên làm việc (cập nhật 14/11/2026)**
* **Trạng thái:** Daily Report đã chuyển sang `.data.json` + `.html` độc lập + template JSON (mục 5.3, `docs/REPORT_JSON_PLAN.md`); danh mục = `ActivityP6.xlsx` (một sheet) + list `CatalogItems`; phân quyền báo cáo = list `Report_Role`; nhận diện người dùng = list `UserProfiles`. Code, `docs/SHAREPOINT_SETUP.md` và khung chạy thử `dev/` đã xong, `tsc`/`eslint`/`lint:styles` sạch, `npm run test:report` đạt 18 test. **CHƯA chạy trên SharePoint thật** và chưa `npm run build` (chủ dự án luôn chạy `heft start`: không chạy `npm run build`/`heft test --clean` khi nó đang chạy, chỉ dùng `tsc --noEmit`, `eslint`, `dev/`).
* **Cài đặt tự động:** `docs/sharepoint_setup/` (`setup-sharepoint.mjs` + `setup.config.json` + `README.md`) tạo 4 list, thư viện `CapNhatCongViec`, thư mục, upload file mẫu, nạp dữ liệu qua SharePoint REST (đăng nhập device code bằng app Entra riêng hoặc `SP_TOKEN`); idempotent, KHÔNG đổi phân quyền. Mới chạy `--dry-run`, CHƯA chạy thật trên SharePoint. Đổi tên list/cột trong code thì sửa `setup.config.json` cho khớp.
* **Việc tiếp theo:** (1) chủ dự án làm theo `docs/SHAREPOINT_SETUP.md` (tạo 4 list từ `*_list.xlsx`, upload `ActivityP6.xlsx` + template JSON + logo, build, deploy) rồi chạy bảng kiểm tra mục 8; (2) sửa lỗi theo thông báo/ảnh chụp chủ dự án gửi; (3) sau đó mới tới backlog bên dưới. Chỗ dễ lỗi khi lên SharePoint thật: tên nội bộ cột list (`getListFieldMap`), ghi file `.html` (SPO có thể buộc tải về hoặc chặn), quyền ghi thư mục `data/<Unit>`, ETag 412 khi hai người lưu cùng lúc, dung lượng HTML khi nhiều ảnh lớn, nạp chunk `react-dom/server`/`xlsx` trong gói SPFx.
* **Đã chốt, không hỏi lại:** app chỉ hỗ trợ JSON (không đọc Excel cũ); ảnh nhúng base64 thu nhỏ cạnh dài 1280px; dòng "Tóm tắt | Summary" trống vẫn in như bản in cũ; Apply template ghép theo `id`; `subcontractors` chỉ là mảng đã chọn (danh sách để chọn là danh mục `subcontractor`); mọi dropdown (kể cả danh sách cố định) lấy từ danh mục, mục dùng chung `Unit = ALL`; tên cột cùng ý nghĩa thì cùng tên giữa các list; cột `Id` của `UserProfiles` đã bỏ; Activity tách riêng (file chỉ-đọc, chỉ admin ghi) vì lý do bảo mật.
* **Dọn dẹp còn lại:** xoá các template Excel cũ `data_test/sharepoint_sample/_template/*.xlsx` (app không đọc); đảm bảo commit không chứa file khoá `~$*`; `docs/REPORT_MODULE.md` đã bị xoá từ trước (xác nhận có chủ ý).
* **Chưa làm:** xem giao diện ở 1024/768px; báo cáo tuần vẫn là placeholder; `departmentCode` của `UserProfiles` chưa dùng.
* **Cách làm việc với chủ dự án:** trả lời tiếng Việt, đề xuất trước khi làm thay đổi lớn; KHÔNG tự `git commit`/`push` (chủ dự án tự commit/amend); mọi ghi chú dự án chỉ ghi vào `AGENTS.md` hoặc `docs/`, không lưu vào `~/.claude`/skill.

**Báo cáo ngày (lịch 14 ngày và tab Tổng quan)**
* **Ngày nghỉ / lễ:** hiện mọi ngày (kể cả Chủ nhật) đều tính, nên ngày nghỉ hiện "Chưa nhập". Cách làm gọn: một danh sách ngày nghỉ (chung hoặc theo đơn vị) lưu trong thư viện; ô ngày nghỉ hiện xám "Nghỉ" và không tính thiếu.
* **Nộp trễ:** đánh dấu ngày "Hoàn tất · nộp trễ" / "quá hạn ký" và đếm số ngày trễ, cần chốt hạn nộp báo cáo và hạn upload bản ký (chung hay theo đơn vị; nếu theo đơn vị thì thêm cột vào list `Report_Role`). Cần đọc thêm thời điểm **tạo** file (`TimeCreated`) vì thời điểm hiện dùng (`TimeLastModified`) đổi mỗi lần lưu lại / thay file.
* **Người lưu / người upload:** sổ `ReportRegistry` đã có `Reporter`; Tổng quan hiện chỉ xuất ra Excel, chưa hiển thị trên giao diện.
* **Chú thích nhanh cho ô ma trận Tổng quan:** hiện dùng `title` của trình duyệt (hiện chậm ~1 giây); ô nằm trong khung cuộn ngang nên chú thích CSS bị cắt, cần vẽ chú thích nổi ngoài khung.
* **Người phụ trách chính:** hiện hiển thị mọi người trong list `Report_Role` của đơn vị; nếu chỉ cần một người chính, thêm cột đánh dấu trong list đó.
* **Nhắc nhở tự động:** web part không tự gửi email; dùng luồng Power Automate chạy theo lịch (đọc cùng thư viện) để nhắc các đơn vị còn thiếu. Hiện chỉ có xuất danh sách ra Excel.
* **Hành động admin chưa có:** upload hộ / thay file ký cho đơn vị khác, xoá file ký để buộc ký lại (cần xác nhận và phân quyền SharePoint tương ứng).
* **Chú thích sửa báo cáo sau khi ký:** hiện sửa sau khi ký không đổi trạng thái (vẫn Hoàn tất); có thể cân nhắc quay về "Chưa upload bản ký" để ký lại.

**Danh mục dùng chung và lưu báo cáo JSON (mô tả ở mục 5.3; đã làm code và đã chạy thử bằng `dev/`, CHƯA thử trên SharePoint thật)**
* Cần làm khi triển khai: (chi tiết từng bước: `docs/SHAREPOINT_SETUP.md`) (1) tạo các list `UserProfiles`, `Report_Role`, `ReportRegistry`, `CatalogItems` từ các file `*_list.xlsx` (sheet `Huong_dan` của từng file); với `CatalogItems` từ `CatalogItems_list.xlsx` (làm theo sheet `Huong_dan`: Enforce unique values cho `Title`, index, phân quyền chỉ thêm; file mẫu đã có thêm các danh mục dùng chung `weather`, `event_type`, `event_status`, `supplier`, `equipment_status`, `issue_severity`, `issue_status`, `subcontractor` ở `Unit = ALL`); (2) upload `ActivityP6.xlsx` lên gốc thư viện `CapNhatCongViec` cùng cấp list `Report_Role` (người dùng chỉ đọc, chỉ admin ghi); (3) upload các template JSON `data_test/sharepoint_sample/_template/template_DRP_<Unit>_20261114.json` (ST01 có dữ liệu demo, ST02/ST04/ST05 cùng cấu trúc nhưng bảng báo cáo trống, dòng mới bắt đầu từ `…-00001`) vào `_template/`; (4) thư mục `resource/` có `thaco_logo.jpg`, `crec_logo.jpg`. Các file template Excel cũ (`_template/*.xlsx`) không còn được đọc: xoá đi. Báo cáo cũ định dạng Excel không mở được (app chỉ đọc `.data.json`).
* Cần kiểm thử: "Khác…" ở Element/Resource/Vật tư ghi dòng vào `CatalogItems` và người dùng khác thấy ngay; nhập trùng (khác hoa thường/dấu/khoảng trắng) thì chọn mục sẵn có; hai người thêm cùng lúc; tài khoản thường không xoá/sửa được dòng; Activity không có "Khác…"; chọn cấu kiện → Activity tự điền; cấu kiện `Đường dẫn` chưa gán Activity → Activity `---` chọn tay; chọn Mã nhân lực/thiết bị/vật tư → Loại/Thiết bị/Vật tư tự điền; thêm dòng Sự kiện/Vướng mắc → mã tự tăng; mã bị xoá khỏi danh mục vẫn hiện ở báo cáo cũ; không có file/sheet đơn vị → thông báo; cột mã hiện đủ ở 1440 / 1024 / 768 / 390px (`columnMinPx` cột dùng danh mục = 170px; đã xem ở 1280 và 390px bằng `dev/`); bản in A4 của các cột mới (Thời tiết 3 buổi + ô Chỉ số đánh giá (text), Ghi chú sự kiện). Với JSON: mở file `.html` đã tải từ SharePoint bằng trình duyệt offline; hai người lưu cùng lúc (ETag 412); `created`/`updated`; Apply template khi bảo trì; ảnh lớn (hàng chục ảnh 3–5MB) làm HTML nặng bao nhiêu; xem ở 1024/768px.
* Chưa làm: ô gõ lọc cho dropdown mã (hiện là `<select>`; danh mục dài hàng trăm mã nên cân nhắc). Mọi giá trị mẫu dùng trong dòng của template phải có trong danh mục (nếu không, chọn sang giá trị khác thì giá trị mẫu biến mất khỏi dropdown); mục Word ghi `---` được nạp vào `CatalogItems` dưới dạng chỉ-có-tên (Code trống). Dữ liệu mẫu trong `ActivityP6.xlsx` (một sheet, 12 dòng: ST01, ST02, ST04, ST05 mỗi đơn vị 3 mã Activity kèm mô tả mẫu) và `CatalogItems_list.xlsx` (mã nhân lực và thiết bị ở `Unit = ALL` một bản; Element và Vật tư riêng từng đơn vị ST01, ST02, ST04, ST05 với mã vật tư và Activity đổi theo đơn vị) là dữ liệu thử theo file Word (Element_ID dùng tên cấu kiện như hiển thị trong Word).
**Bảo mật / triển khai (kiểm tra trên SharePoint, code không thay thế được)**
* `UserProfiles`: người thường chỉ đọc, chỉ admin được sửa; `SiteAssets/portal_config.json`: chỉ admin được ghi; thư viện `CapNhatCongViec`: phân quyền theo thư mục đơn vị `data/<Unit>`, chỉ admin được xoá file ký.
* Tuỳ chọn mạnh hơn: xác định admin theo nhóm SharePoint (ví dụ Site Owners) thay vì cột `Role` trong `UserProfiles`.
* Cột `departmentCode` của `UserProfiles` được đọc nhưng không dùng (có thể bỏ khỏi code). Cột `Id` đã bỏ (SharePoint có sẵn cột ID hệ thống); các tên `field_N` trong `mapSharePointUserProfile` chỉ còn là dự phòng cho list cũ.

**Kiểm thử còn thiếu (chưa xác nhận trên SharePoint thật)**
* Chưa xem trực quan ở 1440 / 1024 / 768 / 390px: thẻ "Ngày đang chọn" của lịch (`report-day-panel`) và thông báo cập nhật template (`report-tpl-notice`). Chưa thử upload bản ký với tên mới `<DocCode>_signed.pdf` (từ form và từ lịch), tải bản ký ở lịch/Tổng quan (đọc tên thật từ thư mục), Apply template khi bảo trì.
* Sổ `ReportRegistry`: tạo list theo mục 5.3, lưu báo cáo mới (số tăng 00001 → 00002, dòng mới), lưu lại (chỉ cập nhật dòng), upload bản ký (Status = done), xoá thư mục `data/` rồi mở báo cáo mới (số không về 00001), hai đơn vị lưu cùng lúc, tài khoản thường không xoá được dòng.
* Tải tab Tổng quan với ~20 đơn vị, upload bản ký (đổi tên + xoá file tên cũ), xuất Excel, chế độ bảo trì với hai tài khoản, `npm run build`, xem giao diện ở 1440 / 1024 / 768 / 390px.
