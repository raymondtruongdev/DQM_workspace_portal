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
│   │   └── report/              # Module: Quản lý Report (Report_Role.xlsx, Daily/Weekly report, tree-table, SharePoint data/<Unit>/<DocID>_<yyyyMMdd>/)
│   └── assets/
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
  - Permissions: `Report_Role.xlsx` (sheet `Role`: `ID`, `Unit`, `DocID`, `DocTitle`, `Site manager`, `Email`, `Status`).
  - Base Templates: `_template/<DocID>_template.xlsx` (`DailyReport_template.xlsx`, `WeeklyReport_template.xlsx`).
  - Daily/Weekly Report Folders & Files: `data/<Unit>/<DocID>_<yyyyMMdd>/<DocID>_<yyyyMMdd>.xlsx` (e.g. `data/ST01/DailyReport_20261110/DailyReport_20261110.xlsx`).
  - Site Photos Subfolder: `data/<Unit>/<DocID>_<yyyyMMdd>/photos/<GroupID>/` (mỗi bộ phận một thư mục con; tải/xoá trực tiếp từ form qua `ReportFileService`).
* **Role Authorization (`core/services/RoleResolutionService.ts`):** User email matches `Report_Role.xlsx` with `Status = TRUE` for valid DocIDs (`DailyReport`, `WeeklyReport`). Admin (`UserProfiles.Role = admin`) has access to all active assignments.
* **Lịch 14 ngày & bản đã ký (`ReportCalendarView.tsx`, `ReportSignedDocument.tsx`, `utils/reportCalendar.ts`):**
  - Luồng báo cáo ngày: danh sách tile → **lịch 14 ngày gần nhất** → form nhập của ngày đã chọn (Back từ form về lịch, lịch đọc lại trạng thái). WeeklyReport (placeholder) vào thẳng form.
  - **3 trạng thái mỗi ngày** (đọc từ thư mục `data/<Unit>/<DocID>_<yyyyMMdd>/`, `ReportFileService.getReportDayInfos`): **Chưa nhập** = chưa có file báo cáo `.xlsx` (chưa bấm Lưu); **Chưa upload bản ký** = có file `.xlsx` nhưng thư mục chưa có file `.pdf`; **Hoàn tất** = có file `.xlsx` và ít nhất một file `.pdf` trong thư mục (không tính thư mục con `photos/`).
  - **Bản đã ký:** sau khi Lưu, người dùng in PDF, ký tên rồi upload bản đã ký ở mục "BẢN ĐÃ KÝ | SIGNED DOCUMENT" cuối form (hoặc nút "Upload file đã ký" ở lịch). Chỉ nhận `.pdf`, lưu cùng thư mục với báo cáo, tên cố định `<DocID>_<yyyyMMdd>_signed.pdf` (upload lại = ghi đè). Chỉ upload được khi báo cáo đã lưu và không còn thay đổi chưa lưu. Sửa báo cáo sau khi ký không đổi trạng thái (vẫn Hoàn tất). Upload gọi `checkMaintenance()` trước khi ghi.
  - Lịch không hiển thị người lưu (thông tin này không có trong danh sách file); chỉ hiển thị thời điểm lưu.
  - **Tab "Tổng quan" (chỉ admin, `ReportOverviewView.tsx`, `utils/reportOverview.ts`)** trên trang danh sách báo cáo (đứng trước tab "Báo cáo của tôi" và là tab mặc định khi admin mở module): hai thẻ tóm tắt: **hôm nay** (số đơn vị Chưa nhập / Chưa upload bản ký / Hoàn tất) và **14 ngày gần nhất** (số báo cáo đơn vị × ngày theo từng trạng thái, kèm phần trăm), **ma trận đơn vị × 14 ngày** (bấm ô mở form của đơn vị ở ngày đó, bấm mã đơn vị mở lịch), **thanh công cụ riêng của ma trận** (làm nổi bật theo trạng thái + tìm theo mã ga/người phụ trách, chỉ tác động lên ma trận), danh sách **Tất cả report** (mọi đơn vị × ngày với cả 3 trạng thái, ngày mới nhất trước, hiện giờ lưu dưới nhãn trạng thái và nút tải bản ký khi đã hoàn tất; khi khung danh sách rộng dưới 900px (đo bề rộng thật của khung bằng ResizeObserver, không theo bề rộng cửa sổ) mỗi dòng xếp thành thẻ nhỏ) có bộ lọc riêng theo trạng thái (Tất cả / Chưa nhập / Chưa upload bản ký / Hoàn tất), theo ngày (Tất cả 14 ngày / Theo ngày với ô chọn ngày, mặc định là hôm nay; cùng một hàng) và theo đơn vị (Tất cả hoặc từng đơn vị) (kèm người phụ trách từ `Report_Role.xlsx` và nút "Sao chép danh sách"). Chỉ theo dõi báo cáo ngày. Dữ liệu đọc tiến dần theo đơn vị (tối đa 3 đơn vị × 5 ngày song song) và nhớ 3 phút trong bộ nhớ trang; nút Làm mới bỏ qua bộ nhớ. Chưa có ngày nghỉ/lễ (mọi ngày đều tính).
* **Navigation & Tile Buttons (`ReportLandingView.tsx`):**
  - Landing page shows tile buttons based on assigned roles (e.g. `ST01 DailyReport`, `ST01 WeeklyReport`, `ST02 DailyReport`), grouped with unit tabs and search filter.
* **Kiến trúc sheet của Daily Report (`reportCodec.ts`, `reportTable.ts`, `reportForm.ts`, `reportSections.ts`):**
  - `01_General`: chỉ thông tin chung dạng Key-Value: `DocTile`, `DocTileEn`, `SubTile`, `DocCode`, `TemplateVersion`, `Location`, `ReportNo`, `ReportDate`, `Shift`, `Email`, `UpdateTime`; cột `Required` đánh dấu `Location`/`ReportNo`/`Shift` bắt buộc. Mọi nội dung báo cáo nằm ở các sheet khác.
  - **Mọi mục của báo cáo là một sheet `0N_<Tên>` tự mô tả** (thời tiết, các bảng, công việc, HSE, ký duyệt, và bất kỳ sheet mới nào). Code chỉ đọc: có sheet → có mục, không có → không có; **thứ tự mục = thứ tự sheet**, số La Mã tự đánh liên tục (không còn mục cố định trong code). **Thêm hoặc bớt mục chỉ cần thêm/bớt sheet trong template, không sửa code.** Có 2 loại sheet:
    - **Sheet bảng** (`reportTable.ts`): dòng 1 `Title | <VI> | <EN> | <nhãn ngắn>` · dòng 2 `Summary | <tóm tắt>` · dòng 3 mã cột (kèm `IsLocked` ở cuối) · dòng 4 nhãn `Tiếng Việt | English` · dòng 5 spec cột · dòng 6+ dữ liệu. Spec: `text` · `number` · `select:A/B/C` · `calc:<công thức>` (+ − × ÷, ngoặc, `round()`, mã cột; chia cho 0 = 0; ví dụ `calc:Nos*Hrs`); cờ nối bằng `;`: `key` (ô chỉ đọc khi dòng `IsLocked`), `total` (dòng Tổng), `left|center|right`, `w=<%>`, `default=<giá trị; {unit} {n3}>`, `tone=<giá trị>:danger|warning,…`.
    - **Sheet form** (`reportForm.ts`): dòng 1 `Title | <VI> | <EN> | <nhãn ngắn>` · dòng 2 `Summary | <tóm tắt> | [off]` (`off` = ẩn ô Tóm tắt) · dòng 3 `Layout | <số cột lưới> | [sign]` (`sign` = in dạng khung chữ ký, tối đa `Layout` cột và tự co theo số trường ký: 2 người ký → 2 cột, 3 → 3 cột) · dòng 4 `Key | Label | Spec | Value | Required` · dòng 5+ mỗi dòng một trường. Tóm tắt luôn ở dòng 2 của mọi sheet (bảng lẫn form), ngay dưới tiêu đề, để thêm dòng/trường mới ở cuối không phải di chuyển nó. Spec: `text` · `textarea` · `number` · `select:A/B/C` · `photos` (khu tải ảnh hiện trường); cờ: `full` (trải hết hàng, in dạng "Nhãn: giá trị"), `span=<n>`, `rows=<n>`, `ph=<gợi ý>`, `default=<giá trị; {user} {siteManager} {unit} tự điền khi trống>`. Cột `Required = TRUE` → bắt buộc nhập trước khi lưu/in. Nhóm các trường ngắn liên tiếp in thành hàng tiêu đề + hàng giá trị.
    - **Sheet ảnh** (`Photo`): dòng 1 `Title | <VI> | <EN> | <nhãn ngắn>` · dòng 2 `Summary | <tóm tắt chung> | [off]` · dòng 3 `Layout | <số ảnh mỗi hàng>` · dòng 4 `GroupID | Department | Summary | Files` · dòng 5+ mỗi dòng một bộ phận (`Files` = tên file ngăn bởi `|`, nằm trong `photos/<GroupID>/`). Template chỉ cần 4 dòng đầu, các nhóm do người dùng thêm trên form bằng nút "Thêm bộ phận" (tên bộ phận + tải ảnh + tóm tắt) và được lưu vào sheet khi bấm Lưu. Báo cáo tạo từ báo cáo ngày trước không sao chép ảnh.
    - Nhận diện: sheet có `Title` ở A1, `Layout` ở A3 và `Key` ở A4 = form; `GroupID` ở A4 = ảnh; có `Title` ở A1 = bảng; sheet khác bị bỏ qua. Chỉ hỗ trợ định dạng này (không còn tương thích template/báo cáo định dạng cũ); không có template trong `_template/` thì báo lỗi "Chưa có template" (không có dữ liệu mẫu trong code).
  - `07_Tasks` (optional, used for `WeeklyReport` tree-table with automatic parent rate calculation).
* **⚠️ WeeklyReport hiện chỉ là placeholder:** tile `WeeklyReport`, file template tuần tạm và code module tuần (`ReportTaskTable`, sheet `07_Tasks`, tính tỉ lệ cha/con) chỉ dùng để hiển thị UI, chưa phải tính năng hoàn chỉnh. Chủ dự án sẽ làm phần báo cáo tuần sau. Không test, sửa logic hay mở rộng phần riêng của weekly (task tree, template/data tuần) và không coi đó là việc tồn đọng cho tới khi được yêu cầu; chỉ giữ UI dùng chung (nút, icon, token) nhất quán.
* **Locked Predefined Template Rows (`IsLocked`):**
  - Rows with `IsLocked = TRUE` in the template (by default `TCJV` and `NTP-TV01` in Manpower, and `<Unit>-DW-1120` / `<Unit>-DW-1130` in WorkQty) cannot be deleted (`🔒` badge replaces `×` button) and their columns flagged `key` in the sheet spec (`Org`, `Category`, `ActivityID`, `Item`, `Unit`) are `readOnly`. Users can still edit quantities/hours and add/delete custom rows (`IsLocked = FALSE`).
* **Smart Data Loading & Previous-Day Prefill (`ReportDataService.loadReport`):**
  1. If `data/<Unit>/<DocID>_<yyyyMMdd>/<DocID>_<yyyyMMdd>.xlsx` exists on SharePoint, loads that exact report (`isExisting = true`).
  2. Otherwise, looks for the adjacent previous day's report (`dateStr - 1 day`) or scans `data/<Unit>/` subfolders for the closest prior `<DocID>_<yyyyMMdd>` report (`getPreviousReportFile`). If found, pre-fills the form from that previous report (`copiedFromDate` badge shown, `isExisting = false`).
  3. If no prior report exists for that unit, falls back to `_template/<DocID>_template.xlsx`.
* **A4 Print Preview & PDF Export (`ReportPrintPreviewModal.tsx`):**
  - Renders the bilingual report (Sections I $\rightarrow$ VIII + 4-column Signature block) in A4 portrait layout (`Mau_BCN_ST06_20261025.pdf` standard) with continuous page flow (no blank page gaps when dynamic rows are added) and full table borders.
  - Ngắt trang: tiêu đề mục + hàng tiêu đề + 2 dòng đầu luôn đi cùng nhau, 2 dòng cuối không bị bỏ lẻ (`break-after: avoid` trên `thead tr`, `tbody tr:first-child`, `tbody tr:nth-last-child(2)`); phần giữa chảy liên tục. Trên điện thoại khung xem trước giữ bề rộng A4 tối thiểu và cuộn ngang.
  - Ảnh hiện trường: nếu có ảnh trong `photos/`, bản in thêm một dòng cuối bảng mục VIII, dưới "Kế hoạch ngày mai (tóm tắt)" (2 ảnh/hàng, mỗi ảnh có tên file; ảnh không bị ngắt giữa hai trang). `handlePrint` đợi mọi ảnh tải xong (tối đa 20s) rồi mới gọi `window.print()`; URL ảnh là tuyệt đối để iframe in tải được bằng phiên SharePoint.
* **Save & Overwrite:**
  - Automatically ensures `data/<Unit>` and `data/<Unit>/<DocID>_<yyyyMMdd>/` exist before saving, with ETag concurrency protection.
* **SheetJS Lazy-Loading:** Loaded via dynamic `import(/* webpackChunkName: 'xlsx' */ 'xlsx')` so other modules do not incur bundle overhead.

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
│                       #   chips (.chip/.search), page-header, feedback (toast, spinner, .state)
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
| Ô nhập | `.input`, `.textarea`; biến thể `input--cell` (trong bảng), `input--readonly`, `input--locked`, `input--danger`, `input--warning` | ô nhập đã `border-box`, cao `--control-h` |
| Nhãn + ô nhập | `.field` > `.field__label` (hoặc `field__label--bilingual` với `field__label-main`/`field__label-sub`) | |
| Đầu trang view | `<PageHeader leading title badge meta actions divider />` | desktop: tiêu đề trái, actions phải; ≤768px xếp dọc |
| Lọc / tìm | `<FilterChips />`, `<SearchInput />` | |
| Lưới ô chức năng (trang chủ, menu module) | `<TileGrid items={[{ id, title, icon, onClick }]} />` | nút xanh đặc: icon + tên + mũi tên; `icon` là `<Icon … />` hoặc `<svg>` thương hiệu |
| Đang tải / rỗng / không quyền | `<LoadingState />`, `<EmptyState icon title message hint action tone narrow />` | `tone="danger"` → `state--danger`; `narrow` → `state--narrow` (khung hẹp giữa trang) |
| Khối nội dung | `.card`; khối có banner xanh đậm: `.section-card` > `__banner` (`__title`, `__title-sub`) + `__body` | |
| Nhãn trạng thái | `.badge` + `badge--info` · `badge--success` · `badge--warning` · `badge--danger` | |
| Bảng nhập liệu | `<div class="table-responsive"><table class="portal-table …">` | cuộn ngang, không tràn trang |
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
