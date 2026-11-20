# Hướng Dẫn Chi Tiết Cập Nhật Website Hue Sound Express Lên GitHub Pages (0đ)

## Bước 1: Kiểm tra những file SẼ đẩy lên GitHub (và file KHÔNG đẩy)
File `.gitignore` đã được tạo sẵn để **tự động chặn** không đẩy nhầm file tạm của Excel (`.~lock...`) và bảng chiết khấu Cộng tác viên nội bộ (`*ctv_discount_private*`).

**Các file/thư mục sẽ được đẩy lên repo `huesoundexpress.github.io`:**
- `index.html` — Trang chủ Phòng Nghe Số (4 danh mục, nghe thử 30s, giỏ hàng, báo giá PDF, Waitlist).
- `data/catalog_all.json` — Kho dữ liệu hợp nhất 4 nhóm sản phẩm (Đĩa than, CD, Máy hát, Phụ kiện & Hue Sound Vinyl Care).
- `data/vinyl_lp_enriched.csv` & `data/templates/` — Dữ liệu bảng chuẩn (trừ file CTV).
- `p/` — Thư mục chứa 472+ trang chi tiết SEO riêng cho từng đĩa/sản phẩm (`/p/<slug>/index.html`).
- `llms.txt` — Dữ liệu chuẩn B2A cho ChatGPT / Gemini / AI Agent.
- `sitemap.xml` — Sơ đồ trang giúp Google lập chỉ mục toàn bộ sản phẩm.
- `scripts/` & `.github/workflows/sync-catalog.yml` — Công cụ cập nhật tự động.
- *(Tùy chọn)* `og-image.jpg` — Ảnh bìa 1200×630px hiển thị khi chia sẻ link `https://huesoundexpress.github.io/` lên Zalo / Facebook.

---

## Bước 2: Đẩy code lên GitHub (Chọn 1 trong 2 cách)

### Cách 1: Dùng Git trên máy tính (Nhanh nhất – Khuyên dùng vì thư mục `p/` có 472 trang SEO)
Mở **PowerShell** tại thư mục `E:\Hue Sound Express\Hue_Sound_Express_GitHub_Bio_B2A`:

**Nếu lần đầu nối thư mục này với repo GitHub `huesoundexpress.github.io`:**
```powershell
git init
git branch -M main
git remote add origin https://github.com/huesoundexpress/huesoundexpress.github.io.git
git add .
git commit -m "feat: nâng cấp Phòng Nghe Số 4 danh mục, nghe thử 30s, báo giá PDF & trang SEO"
git push -u origin main --force
```
*(Lưu ý: Nếu repo cũ trên GitHub có sẵn file `og-image.jpg`, hãy tải file `og-image.jpg` về đặt cạnh `index.html` trước khi chạy `git push --force` để không bị mất ảnh chia sẻ Zalo/Facebook!)*

**Từ lần cập nhật thứ 2 trở đi** (khi bạn sửa giá, thêm đĩa CD/máy hát mới):
```powershell
# 1. Chay lai build de cap nhat catalog_all.json, llms.txt, trang /p/ va sitemap.xml
powershell -ExecutionPolicy Bypass -File .\scripts\build_site_data.ps1 -GenerateProductPages

# 2. Day len GitHub (web tu cap nhat sau 1-2 phut)
git add .
git commit -m "chore: cap nhat kho dia va gia niem yet"
git push
```

### Cách 2: Dùng phần mềm GitHub Desktop (Không cần gõ lệnh)
1. Tải **GitHub Desktop** (miễn phí), đăng nhập tài khoản GitHub `huesoundexpress`.
2. Chọn **Clone repository** `huesoundexpress.github.io` về máy.
3. Copy toàn bộ các file/thư mục ở Bước 1 dán đè vào thư mục vừa clone *(trừ file `ctv_discount_private_template.csv`)*.
4. Mở GitHub Desktop → nhập tiêu đề ngắn ở ô **Summary** (vd: `Cập nhật Phòng Nghe Số v2`) → bấm **Commit to main** → bấm **Push origin**.

---

## Bước 3: Kiểm tra GitHub Pages & Đăng ký Google Index (Làm 1 lần duy nhất)
1. Vào repo `huesoundexpress.github.io` trên GitHub.com → chọn tab **Settings** → mục **Pages** (cột trái):
   - **Source:** `Deploy from a branch`
   - **Branch:** `main` / `/ (root)` → bấm **Save**.
2. Chờ 1–2 phút, mở trình duyệt vào `https://huesoundexpress.github.io/` để kiểm tra:
   - Kiểm tra 4 tab danh mục, nút **▶ 30s**, nút **+ Giỏ / Báo giá**, nút **📄 Xuất Báo Giá PDF** và nút **🔔 Báo khi về**.
3. Để Google tìm thấy toàn bộ 472+ trang sản phẩm:
   - Truy cập [Google Search Console](https://search.google.com/search-console/) → thêm tài sản `https://huesoundexpress.github.io/` → vào mục **Sitemaps** → nhập `sitemap.xml` → bấm **Gửi (Submit)**.

---

## Bước 4: Bật thông báo Đơn hàng & Waitlist về điện thoại (0đ – Làm trong 3 phút)
Hiện tại nếu chưa gắn Webhook, khi khách bấm đặt hàng web sẽ tự mở Zalo kèm sẵn toàn bộ đơn hàng. Để khách đặt **không cần mở Zalo** mà điện thoại bạn vẫn ting-ting ngay lập tức:
1. Tạo 1 **Google Sheet** mới tên `Hue Sound Express - Đơn Hàng`.
2. Vào menu **Tiện ích mở rộng (Extensions)** → **Apps Script**.
3. Copy toàn bộ mã trong file `scripts/google_apps_script_order_notify.gs` dán vào.
4. Điền `NOTIFY_EMAIL` (email của bạn) và/hoặc `TELEGRAM_BOT_TOKEN` + `TELEGRAM_CHAT_ID` (tạo bot miễn phí với `@BotFather` trên Telegram).
5. Bấm **Deploy** → **New deployment** → chọn loại **Web app** → `Execute as: Me`, `Who has access: Anyone` → bấm **Deploy** và copy đường link `https://script.google.com/macros/s/.../exec`.
6. Mở `index.html`, tìm dòng `const WEBHOOK_URL = "";` → dán link vừa copy vào giữa dấu ngoặc kép → `git push` lên GitHub.

---

## Bước 5: Quy trình bổ sung các đĩa còn thiếu Ảnh bìa / Audio 30s
1. Mở file `data/missing_media_report.csv` bằng Excel.
2. Các đĩa Nhạc Việt hiếm hoặc bản nội địa chưa có trên iTunes/Deezer được liệt kê kèm sẵn cột link tìm nhanh trên **Discogs**, **Google Images**, **YouTube**.
3. Upload ảnh bìa hoặc file cắt MP3 30s (nhẹ ~300KB/file) lên **Cloudflare R2** (miễn phí 10GB) hoặc dán trực tiếp link ảnh vào cột `cover_url` / `audio_url` trong `data/vinyl_lp_enriched.csv`.
4. Chạy lại `.\scripts\build_site_data.ps1 -GenerateProductPages` và `git push`.
