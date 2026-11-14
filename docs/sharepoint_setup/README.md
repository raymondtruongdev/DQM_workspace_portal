# Tự động cài đặt SharePoint (Workspace Portal)

Script `setup-sharepoint.mjs` thay cho các bước tạo tay ở mục 2 và 3 của `docs/SHAREPOINT_SETUP.md`: tạo 4 list, thư viện `DailyReportLibrary`, thư mục, upload file mẫu và nạp dữ liệu ban đầu. Chạy lại nhiều lần an toàn (chỉ bổ sung phần còn thiếu).

| File | Vai trò |
|---|---|
| `setup-sharepoint.mjs` | Script chạy (Node >= 22, không cần cài gói mới; đọc xlsx bằng gói `xlsx` có sẵn của repo) |
| `*_list.xlsx` (4 file) | File mẫu tạo list bằng tay (sheet `Huong_dan`) và là nguồn dữ liệu nạp mẫu của script |
| `setup.config.json` | Cấu hình: URL site, danh sách list + cột + index + dữ liệu nạp, thư viện, file cần upload |

## 1. Việc script làm / không làm

**Làm:**
- List `UserProfiles`, `ReportRole`, `ReportRegistry`, `CatalogItems`: tạo cột đúng tên nội bộ (không bị `field_N`), kiểu Text/Number(0 số lẻ)/Yes-No, `Title` Enforce unique values + Indexed (`ReportRegistry`, `CatalogItems`), index `Unit`/`ReportDate`/`ReportNo`/`Email`, bật versioning.
- Thư viện `DailyReportLibrary` + thư mục `_template`, `resource`, `data`, `data/<Unit>` (đơn vị lấy từ tên các file template).
- Upload `ActivityP6.xlsx`, `_template/template_DRP_*.json`, logo `resource/*.jpg` từ `docs/sharepoint_setup/DailyReportLibrary/`. File đã có trên SharePoint thì **bỏ qua** (tránh đè bản admin đã sửa), dùng `--force` để ghi đè.
- Nạp dữ liệu: `CatalogItems` (77 dòng mẫu), admin đầu tiên vào `UserProfiles` (`--admin`), `ReportRole` mẫu (chỉ khi có `--seed-sample`). `ReportRegistry` để trống (web part tự ghi).

**Không làm (làm tay):** phân quyền list/thư viện (mục 5), build + deploy `.sppkg`, xoá dòng mẫu nếu bạn đã dùng `--seed-sample`.

## 2. Đăng nhập (chọn một cách)

### Cách A: app Entra riêng (khuyên dùng, đăng nhập device code)
Làm một lần trong Entra ID (portal.azure.com) > App registrations > New registration:
1. Name: `Workspace Portal Setup`; Supported account types: *Single tenant*; không cần Redirect URI.
2. Authentication > **Allow public client flows = Yes** > Save.
3. API permissions > Add > **SharePoint** > *Delegated* > `AllSites.FullControl` > **Grant admin consent**.
4. Copy **Application (client) ID**, điền vào `setup.config.json` (`clientId`) hoặc truyền `--client-id`.

Chạy script, làm theo hướng dẫn in ra (mở trang `microsoft.com/devicelogin`, nhập mã, đăng nhập **tài khoản Site Owner/admin site**).

### Cách B: có sẵn access token
Nếu bạn dùng CLI khác (ví dụ CLI for Microsoft 365, PnP PowerShell) để lấy token SharePoint của site:
```bash
SP_TOKEN="<jwt>" node docs/sharepoint_setup/setup-sharepoint.mjs --admin you@company.com
```
Token hết hạn sau khoảng 1 giờ.

## 3. Chạy

Từ gốc repo:

```bash
# 1) Xem kế hoạch, không gọi SharePoint, không cần đăng nhập
node docs/sharepoint_setup/setup-sharepoint.mjs --dry-run --admin you@company.com

# 2) Cài thật lần đầu
node docs/sharepoint_setup/setup-sharepoint.mjs --client-id <GUID> --admin you@company.com

# Site khác với config
node docs/sharepoint_setup/setup-sharepoint.mjs --site https://tenant.sharepoint.com/sites/xyz --client-id <GUID> --admin you@company.com
```

| Tuỳ chọn | Ý nghĩa |
|---|---|
| `--admin <email>` | Admin đầu tiên (UPN M365) ghi vào `UserProfiles` với `Role = ADMIN`. **Bắt buộc lần đầu**: không có admin thì không ai vào được Admin Config |
| `--seed-sample` | Nạp thêm `ReportRole` mẫu (email trong file mẫu là của dự án dev, chỉ dùng thử) |
| `--no-seed-catalog` | Không nạp 77 dòng mẫu của `CatalogItems` |
| `--skip-lists` / `--skip-files` | Chỉ làm phần thư viện / chỉ làm phần list |
| `--force` | Ghi đè file đã có trong thư viện (ví dụ upload lại `ActivityP6.xlsx`) |
| `--config <file>` | Dùng file cấu hình khác |
| `--dry-run` | Chỉ in kế hoạch |

Sửa `setup.config.json` khi cần thêm cột, thêm file upload hay đổi dữ liệu nạp; tên list/cột phải khớp code (xem `_ghi_chu` trong file). Thêm đơn vị mới: đặt template `template_DRP_<Unit>_<yyyyMMdd>.json` vào `docs/sharepoint_setup/DailyReportLibrary/_template/` rồi chạy lại (script tạo `data/<Unit>` và upload template mới).

## 4. Sau khi chạy

- Mở site, kiểm tra 4 list và thư viện đã có; cột `Title` của `ReportRegistry`/`CatalogItems` có Enforce unique values.
- Thay dữ liệu mẫu bằng dữ liệu thật: `UserProfiles` (người dùng), `ReportRole` (phân quyền báo cáo), `ActivityP6.xlsx` (Activity thật).
- Nếu có dòng mẫu không dùng thì xoá (script không bao giờ xoá dữ liệu).

## 5. Phân quyền: làm tay trên SharePoint

Script chủ ý không đổi quyền (sai quyền khó phát hiện và là phần bảo mật thật). Cấu hình theo bảng:

| Đối tượng | Người báo cáo | Admin |
|---|---|---|
| `UserProfiles`, `ReportRole` | Chỉ đọc | Toàn quyền |
| `ReportRegistry` | Thêm + sửa, **không xoá** | Toàn quyền |
| `CatalogItems` | Chỉ **thêm** (không sửa/xoá) | Toàn quyền |
| Thư viện gốc, `_template`, `resource`, `ActivityP6.xlsx` | Chỉ đọc | Toàn quyền |
| `data/` hoặc `data/<Unit>` | Contribute (không xoá file bản ký) | Toàn quyền |
| `SiteAssets` (chỗ ghi `portal_config.json`) | Không ghi | Ghi |

Chi tiết và lý do: `docs/SHAREPOINT_SETUP.md` mục 2.4 và 3. Sau đó build/deploy (mục 5 tài liệu đó) và chạy bảng kiểm tra mục 8.

## 6. Lỗi thường gặp

| Hiện tượng | Xử lý |
|---|---|
| `401/403` | Tài khoản đăng nhập không phải Site Owner, hoặc chưa Grant admin consent cho `AllSites.FullControl` |
| `AADSTS7000218` / lỗi public client | Chưa bật *Allow public client flows* (bước 2 mục 2) |
| `Enforce unique values` lỗi khi thêm dữ liệu | List đã có dòng trùng `Title`; xoá dòng trùng rồi chạy lại |
| Không index được cột (cảnh báo) | List > 5000 dòng thì SharePoint chặn tạo index trực tiếp; tạo tay trong List settings |
| Cột tên khác (`Unit0`, `field_1`) | List đã được tạo tay trước đó với tên khác; web part tự tra tên nội bộ nên thường vẫn chạy, hoặc xoá list rồi chạy lại |
