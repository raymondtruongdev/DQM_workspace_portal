# THADICO DQM - Workspace Portal (SPFx 1.24.1)

Ứng dụng **Workspace Portal** phát triển trên nền tảng **SharePoint Framework (SPFx 1.24.1)** và **React 18**, vận hành trực tiếp trên **SharePoint Online** bằng tài khoản Microsoft 365 tiêu chuẩn (miễn phí 100% bản quyền, không yêu cầu Power Apps Premium).

---

## 1. Các phân hệ chính (Modules)

### 1.1. Quản lý Báo cáo Ngày (`src/modules/report`)

- Lưu trữ trên SharePoint Document Library **`CapNhatCongViec`**; báo cáo ngày lưu dạng **JSON + HTML độc lập** (app chỉ hỗ trợ JSON, không còn đọc/ghi báo cáo hay template Excel):
  - Template báo cáo: `_template/template_DRP_<Unit>_<yyyyMMdd>.json` (ngày trong tên file là phiên bản, bản mới nhất thắng).
  - Mỗi báo cáo một thư mục `data/<Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/` gồm `<DRP|WRP>_<yyyyMMdd>_<Unit>.json` (dữ liệu), `<DRP|WRP>_<yyyyMMdd>_<Unit>.html` (bản xem/in độc lập, nhúng sẵn dữ liệu + ảnh + logo), `photos/<GroupID>/` (ảnh hiện trường) và `<DocCode>_signed.pdf` (bản đã ký).
  - Phân quyền theo người: list `ReportRole`; nhận diện đăng nhập: list `UserProfiles`; sổ quản lý số báo cáo/trạng thái: list `ReportRegistry`; danh mục dùng chung: list `CatalogItems` + file `ActivityP6.xlsx` (chỉ admin ghi).
- **Tính năng nổi bật:**
  - **Lịch 14 ngày và 3 trạng thái** (Chưa nhập / Chưa upload bản ký / Hoàn tất), tab **Tổng quan** cho admin (ma trận đơn vị × ngày, xuất Excel).
  - **Tự động kế thừa dữ liệu ngày trước:** ngày mới chưa có báo cáo thì lấy báo cáo ngày trước làm nền (áp dụng luỹ kế/đặt lại theo template); đơn vị chưa có báo cáo nào thì tạo từ template.
  - **Thêm/bớt mục, cột, trường chỉ cần sửa template JSON** (đổi ngày trong tên file để báo cập nhật), không sửa code. Dòng mẫu `locked` không xoá được.
  - **Danh mục dùng chung và mã tự tăng** (Event ID, Issue ID), dropdown có "Khác…" để thêm giá trị mới.
  - **Xem bản in / PDF A4 song ngữ Việt – Anh** và **bản đã ký** (upload PDF ký tay).
  - **Chế độ bảo trì** (Admin Config → Shutdown Server) chặn người thường ở mọi module.
- Báo cáo tuần (`WeeklyReport`) hiện chỉ là placeholder.
- Chi tiết nghiệp vụ và hợp đồng dữ liệu: [`AGENTS.md`](./AGENTS.md) mục 5.3, [`docs/REPORT_JSON_PLAN.md`](./docs/REPORT_JSON_PLAN.md).

---

## 2. Yêu cầu hệ thống & Cài đặt

- **Node.js:** `v22.x` (theo chuẩn SPFx `1.24.1`)
- **Trình quản lý gói:** `npm`

```bash
cd /Volumes/DATA/THADICO_local/REPO_DQM/DQM_workspace_portal
npm install
```

---

## 3. Thiết lập SharePoint (một lần cho mỗi site)

Cần **4 SharePoint List** và **1 Document Library** trên site đích (Dev: `https://ytf84.sharepoint.com/sites/metro-checkin-dev`):

| Tài nguyên | Loại | Dùng cho |
|---|---|---|
| `UserProfiles` | List | Nhận diện đăng nhập, vai trò `admin` / `user`, phòng ban |
| `ReportRole` | List | Ai được làm báo cáo nào (đơn vị × loại báo cáo) |
| `ReportRegistry` | List | Sổ báo cáo: số báo cáo tự tăng, trạng thái, bộ đếm mã |
| `CatalogItems` | List | Danh mục dùng chung (Element, Resource, Vật tư, thời tiết, trạng thái…) |
| `CapNhatCongViec` | Document Library | `ActivityP6.xlsx`, `_template/`, `resource/` (logo), `data/<Unit>/...` |

Cấu trúc thư viện:

```text
CapNhatCongViec/
├── ActivityP6.xlsx                  # Danh mục Activity (một sheet), người dùng chỉ đọc, chỉ admin ghi
├── _template/
│   └── template_DRP_<Unit>_<yyyyMMdd>.json
├── resource/                        # thaco_logo.jpg, crec_logo.jpg (logo bản in)
└── data/                            # App tự tạo
    └── <Unit>/<DRP|WRP>_<yyyyMMdd>_<Unit>/
        ├── <DRP|WRP>_<yyyyMMdd>_<Unit>.json
        ├── <DRP|WRP>_<yyyyMMdd>_<Unit>.html
        ├── photos/<GroupID>/
        └── <DocCode>_signed.pdf
```

- **Hướng dẫn từng bước:** [`docs/SHAREPOINT_SETUP.md`](./docs/SHAREPOINT_SETUP.md). File tạo list (`*_list.xlsx`) và script cài đặt tự động nằm ở [`docs/sharepoint_setup/`](./docs/sharepoint_setup/README.md).
- **File mẫu** (`ActivityP6.xlsx`, template JSON, logo, báo cáo mẫu) nằm ở `data_test/sharepoint_sample/`, tải lên thủ công (hoặc bằng script); thư mục `data/...` do app tự tạo.
- Tài khoản phải có trong **cả** `UserProfiles` và `ReportRole` mới làm được báo cáo; admin (`UserProfiles.Role = admin`) thấy mọi phân quyền.
- Người dùng báo cáo cần quyền **Edit** trên thư mục `data/<Unit>` của mình. Phân quyền SharePoint là bảo mật thật; kiểm tra quyền ở giao diện chỉ là tiện ích.

---

## 4. Chạy thử nghiệm ở môi trường Local (Development)

```bash
npm run start
```

1. Dịch vụ local chạy tại `https://localhost:4321`.
2. Nếu mở trên trình duyệt mới lần đầu, truy cập `https://localhost:4321/temp/build/manifests.js` $\rightarrow$ Chọn **Advanced** $\rightarrow$ **Proceed to localhost (unsafe)** để tin cậy chứng chỉ SSL local.
3. Mở SharePoint Hosted Workbench tại:
   - `https://ytf84.sharepoint.com/sites/metro-checkin-dev/_layouts/15/workbench.aspx?reset=true`
   - Nếu trang vẫn chạy bản đã deploy cũ (tên file web part có mã hash thay vì `localhost:4321`) hoặc Console báo CORS/loopback: cho phép **Local network access** cho trang SharePoint trong Site settings của Chrome, rồi dùng link dự phòng có `debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true` và bấm **Load debug scripts**.

**Chạy thử module Report không cần SharePoint:** `npm run dev:report` (SharePoint giả lập trong bộ nhớ, mở `temp/dev/index.html`); `npm run test:report` chạy kiểm thử chuyển đổi JSON ↔ mô hình ↔ template. Không chạy `npm run build` khi `npm run start` đang chạy.

**Kiểm tra giao diện/quy ước style:** `npm run lint:styles`, `npx tsc --noEmit -p tsconfig.json`, `npx eslint src --ext .ts,.tsx`.

---

## 5. Đóng gói bản Production (`.sppkg`) & Triển khai

```bash
npm run build
```

- Lệnh trên thực thi `heft test --clean --production` và `heft package-solution --production`.
- Gói cài đặt thành phẩm được xuất ra tại:
  `sharepoint/solution/workspace_portal.sppkg`

### Các bước triển khai lên SharePoint App Catalog:

1. Mở **SharePoint App Catalog** của site hoặc tenant (`Apps for SharePoint`).
2. Tải file `sharepoint/solution/workspace_portal.sppkg` lên và chọn **Deploy**.
3. Mở trang SharePoint đích $\rightarrow$ **Edit** $\rightarrow$ Bấm **`+`** $\rightarrow$ Chọn **Workspace Portal** $\rightarrow$ **Republish**.

---

## 6. Tài liệu tham khảo nội bộ

- [`AGENTS.md`](./AGENTS.md): Quy chuẩn kiến trúc, quy tắc UI đa thiết bị (Responsive) và hướng dẫn dành cho AI Agent / Developer.
- [`project_architecture.md`](./project_architecture.md): Sơ đồ kiến trúc phân tầng Multi-Module (`entry`, `app`, `core`, `modules`).
- [`docs/REPORT_JSON_PLAN.md`](./docs/REPORT_JSON_PLAN.md): Hợp đồng JSON/HTML của báo cáo ngày và lý do thiết kế.
- [`docs/SHAREPOINT_SETUP.md`](./docs/SHAREPOINT_SETUP.md) và [`docs/sharepoint_setup/README.md`](./docs/sharepoint_setup/README.md): Cài đặt list/thư viện trên SharePoint (thủ công và script).
