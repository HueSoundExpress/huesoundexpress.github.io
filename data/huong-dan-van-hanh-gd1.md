# Hướng Dẫn Vận Hành Giai Đoạn 1 (0đ) – Hue Sound Express

## 1. Cấu trúc thư mục mới đã thiết lập
- `data/SCHEMA.md` — Quy chuẩn dữ liệu 4 nhóm hàng + quy tắc thương hiệu & báo giá.
- `data/templates/` — 4 file mẫu CSV chuẩn UTF-8 (mở trực tiếp bằng Excel không lỗi tiếng Việt hoặc import lên Google Sheets):
  - `cd_sacd_template.csv` (Đĩa CD / SACD)
  - `turntables_template.csv` (Máy hát / Mâm than)
  - `accessories_care_template.csv` (Phụ kiện & Hue Sound Vinyl Care)
  - `ctv_discount_private_template.csv` (Bảng chiết khấu Cộng tác viên — **chỉ dùng nội bộ, KHÔNG upload lên web**)
- `scripts/enrich_catalog.ps1` — Tự động kéo **Ảnh bìa 600px + Tracklist + Năm phát hành + Audio nghe thử 30s** từ iTunes (VN/US), Deezer, MusicBrainz (+ tùy chọn Discogs).
- `scripts/build_site_data.ps1` — Hợp nhất 4 nhóm hàng → xuất `data/catalog_all.json`, cập nhật `llms.txt` chuẩn B2A và sinh trang chi tiết riêng chuẩn SEO `/p/<slug>/index.html` + `sitemap.xml`.
- `scripts/google_apps_script_order_notify.gs` — Nhận đơn hàng / đăng ký Waitlist từ khách không cần đăng nhập → báo **Telegram 1 giây** + **Email** + ghi **Google Sheets** (0đ).
- `.github/workflows/sync-catalog.yml` — Tự động build lại kho & trang SEO mỗi 6 tiếng trên GitHub Actions (0đ).

## 2. Chạy lấy Ảnh bìa & Audio 30s cho toàn bộ 456 đĩa than
Mở PowerShell tại thư mục dự án và chạy:
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\enrich_catalog.ps1 -Resume
```
- Script tự lưu tiến độ sau mỗi 15 đĩa (nếu ngắt giữa chừng chỉ cần chạy lại lệnh trên với `-Resume`).
- Nếu bạn có tài khoản Discogs (miễn phí), tạo Personal Access Token tại `discogs.com/settings/developers` và thêm `-DiscogsToken "MA_TOKEN"` để tăng tỷ lệ tìm thấy ảnh bìa & tracklist các đĩa Audiophile / Japan Pressing hiếm.
- Sau khi chạy xong, mở file `data\missing_media_report.csv` bằng Excel: danh sách các đĩa còn thiếu ảnh bìa / audio 30s sẽ nằm ngay đầu bảng kèm sẵn link tìm nhanh trên Discogs / Google Images / YouTube để bổ sung thủ công (hoặc tải file MP3 nén nhẹ lên Cloudflare R2).

## 3. Cập nhật Website & Sinh trang chi tiết SEO
Sau khi điền dữ liệu CD, Máy hát, Phụ kiện (hoặc chỉnh sửa `data\vinyl_lp_enriched.csv`), chạy:
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_site_data.ps1 -GenerateProductPages
```
Lệnh này sẽ tự động cập nhật:
1. `data\catalog_all.json` (dữ liệu tổng cho giao diện web 4 danh mục)
2. `llms.txt` (chuẩn AI Agent / B2A mới nhất)
3. Thư mục `p\<slug>\index.html` + `sitemap.xml` (trang riêng cho từng đĩa/máy để Google index)
