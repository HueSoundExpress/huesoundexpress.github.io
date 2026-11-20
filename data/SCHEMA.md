# Kiến Trúc Dữ Liệu Hue Sound Express (Giai Đoạn 1 – 0đ)

Tài liệu này định nghĩa cấu trúc dữ liệu chuẩn hóa **Một Nguồn Duy Nhất (Single Source of Truth)** cho cả 4 nhóm sản phẩm của **Hue Sound Express** và bảng chiết khấu nội bộ cho Cộng tác viên (CTV).

---

## 1. Quy ước Thương hiệu & Vận hành Bắt buộc
1. **Tên thương hiệu:**
   - Viết chuẩn: `Hue Sound Express` và `Hue Sound Vinyl Care`.
   - Chỉ viết hoa toàn bộ (`HUE SOUND EXPRESS`, `HUE SOUND VINYL CARE`) khi cả cụm dùng chữ in hoa (CAPSLOCK).
2. **Địa chỉ giao dịch:**
   - **Không** quảng bá 298 Võ Chí Công là cửa hàng mở cửa trực tiếp trên Website, `llms.txt` hay AI Chatbot.
   - Khi khách có nhu cầu gặp trực tiếp bàn giao đĩa/máy, chuyển cuộc hội thoại cho người thật qua Zalo `09323.13918` để hẹn lịch.
3. **Giá & Trạng thái hiển thị:**
   - **Luôn hiển thị Giá niêm yết** (giá hãng phát hành / giá chuẩn).
   - 4 trạng thái chuẩn hóa trên web:
     - `New Seal` — Hàng mới nguyên seal, sẵn kho, đặt mua ngay.
     - `Hết hàng` — Tạm hết, hiện nút **"Báo tôi khi hàng về" (Waitlist)**.
     - `Sắp có hàng` — Đang về kho, hiện nút **"Báo tôi khi hàng về" (Waitlist)**.
     - `Liên hệ` — Dành cho bản chữ ký (`Sign`), hàng hiếm, bản giới hạn hoặc đĩa chưa có giá cố định → mở kênh tư vấn trực tiếp.
4. **Quy tắc Xuất Báo Giá (Instant Quote):**
   - Dùng cho cá nhân/doanh nghiệp làm thủ tục tạm ứng hoặc đối chiếu quà tặng.
   - **Không** ghi Hue Sound Express xuất hóa đơn VAT trực tiếp.
   - Trong bảng báo giá: cột **Tên hàng hóa** ghi nhóm chung là **"Đĩa nhạc"** (hoặc *"Thiết bị âm thanh"*, *"Phụ kiện audio"*), còn tên album, ca sĩ, mã SKU chi tiết đưa vào cột **Ghi chú / Chi tiết**.

---

## 2. Cấu trúc 4 Bảng Dữ liệu Công khai (Public Sheets / CSV)

Mọi file CSV đều lưu định dạng **UTF-8 with BOM** để mở trực tiếp bằng Microsoft Excel không lỗi dấu tiếng Việt, đồng thời dễ dàng import lên **Google Sheets**.

### Bảng 1: Đĩa Than (`vinyl_lp`)
| Cột | Ý nghĩa | Ví dụ |
| :--- | :--- | :--- |
| `id` | Số thứ tự duy nhất | `1` |
| `sku` | Mã SKU nội bộ | `VN-MYTAM-TAM9-LP` |
| `slug` | Đường dẫn URL SEO (`/p/<slug>/`) | `my-tam-tam-9-r-1` |
| `category` | Nhóm sản phẩm (cố định) | `vinyl` |
| `genre` | Thể loại chính | `Nhạc Việt`, `Jazz & Vocal`, `Rock & Metal`... |
| `mood` | Tâm trạng / Dịp nghe (ngăn cách `;`) | `Đêm muộn; Quà tặng; Hoài niệm` |
| `artist` | Ca sĩ / Nghệ sĩ | `Mỹ Tâm` |
| `title` | Tên Album | `Tâm 9 [R]` |
| `label` | Hãng phát hành | `Viết Tân Audio` |
| `format` | Định dạng vật lý | `LP`, `2xLP`, `3xLP`, `Gatefold` |
| `edition` | Phiên bản / Màu phôi | `Black`, `Coloured`, `Sign`, `Limited` |
| `year` | Năm phát hành | `2018` |
| `list_price` | Giá niêm yết (VND, số nguyên) | `1750000` |
| `status` | Trạng thái (`New Seal`, `Hết hàng`, `Sắp có hàng`, `Liên hệ`) | `New Seal` |
| `description` | Mô tả ngắn chuẩn Audiophile | `Định dạng LP chuẩn mực...` |
| `tracklist` | Danh sách bài hát (ngăn cách `\|`) | `Đừng Hỏi Em \| Cô Ấy Là Ai \| ...` |
| `cover_url` | Link ảnh bìa (từ iTunes/Deezer/R2/local) | `https://is1-ssl.mzstatic.com/.../600x600bb.jpg` |
| `audio_url` | Link nghe thử 30s (iTunes/Deezer/R2/local) | `https://audio-ssl.itunes.apple.com/...` |
| `audio_track` | Tên bài hát đang làm đoạn sample 30s | `Đừng Hỏi Em` |
| `media_source` | Nguồn dữ liệu tự động | `itunes-vn`, `deezer`, `manual-r2` |
| `missing_fields` | Các trường còn thiếu cần bổ sung tay | `MISSING_AUDIO; MISSING_TRACKLIST` |

### Bảng 2: Đĩa CD / SACD (`cd_sacd`)
Cấu trúc giống hệt Bảng 1, chỉ khác:
- `category` = `cd`
- `format` = `CD`, `2xCD`, `SACD`, `XRCD`, `HDCD`, `LPCD`, `Boxset`
- `edition` = `Phôi tiêu chuẩn`, `Phôi Nhật`, `Phôi Đức`, `Bản chữ ký`, `Limited`

### Bảng 3: Máy Hát / Mâm Than (`turntables`)
| Cột | Ý nghĩa | Ví dụ |
| :--- | :--- | :--- |
| `id`, `sku`, `slug` | Định danh & URL | `TT-AT-LP60XBT` |
| `category` | Cố định | `turntable` |
| `brand` | Thương hiệu | `Audio-Technica`, `Pro-Ject`, `TEAC`, `Rega` |
| `title` | Tên model | `Audio-Technica AT-LP60XBT` |
| `level` | Phân khúc người chơi | `Người mới bắt đầu`, `Trung cấp`, `Audiophile` |
| `drive_type` | Cơ chế truyền động | `Belt-drive`, `Direct-drive` |
| `has_phono` | Tích hợp Phono Preamp | `Có` / `Không` |
| `has_bluetooth` | Tích hợp Bluetooth | `Có` / `Không` |
| `cartridge_type` | Chuẩn kim hỗ trợ | `MM`, `MC`, `MM / MC` |
| `cartridge_model` | Kim đi kèm máy | `AT3600L`, `AT-VM95E`, `Ortofon OM10` |
| `list_price` | Giá niêm yết (VND) | `6900000` |
| `status` | Trạng thái | `New Seal`, `Hết hàng`, `Sắp có hàng`, `Liên hệ` |
| `specs_summary` | Tóm tắt thông số nổi bật | `Tự động hoàn toàn; Tốc độ 33/45 RPM; Có Phono & Bluetooth` |
| `recommended_skus` | SKU phụ kiện gợi ý mua kèm | `ACC-CARBON-01; ACC-SLEEVE-50` |
| `description` | Mô tả tư vấn thực tế | `Lựa chọn dễ dùng nhất cho người mới...` |
| `cover_url` | Ảnh sản phẩm chính | `https://...` |
| `video_demo_url` | Link clip test thực tế (nếu có) | `https://...` |

### Bảng 4: Phụ Kiện & Hue Sound Vinyl Care (`accessories_care`)
| Cột | Ý nghĩa | Ví dụ |
| :--- | :--- | :--- |
| `id`, `sku`, `slug` | Định danh & URL | `ACC-AQ-CARBON` |
| `category` | `accessory` hoặc `care` | `accessory` |
| `sub_type` | Nhóm con | `Vệ sinh đĩa`, `Bao lót & Bìa bảo vệ`, `Kim & Máng`, `Cân chỉnh`, `Hue Sound Vinyl Care` |
| `brand` | Thương hiệu | `AudioQuest`, `Audio-Technica`, `Ortofon`, `Hue Sound Vinyl Care` |
| `title` | Tên sản phẩm / Dịch vụ | `Chổi sợi Carbon chống tĩnh điện` |
| `compatibility` | Khả năng tương thích | `Mọi đĩa than 12-inch`, `Kim MM`, `Mọi mâm than` |
| `level` | Mức độ cần thiết | `Cơ bản (Nên có)`, `Nâng cấp` |
| `list_price` | Giá niêm yết (VND) | `550000` |
| `status` | Trạng thái | `New Seal`, `Hết hàng`, `Sắp có hàng`, `Liên hệ` |
| `description` | Công dụng & hướng dẫn ngắn | `Quét bụi mịn và triệt tiêu tĩnh điện trước khi hạ kim.` |
| `cover_url` | Ảnh sản phẩm | `https://...` |

---

## 3. Bảng Nội Bộ Cho Cộng Tác Viên (KHÔNG Public lên Web)
File: `data/templates/ctv_discount_private_template.csv`
- Lưu trên 1 **Google Sheet riêng biệt** (chỉ share quyền Xem cho CTV theo từng cấp, tuyệt đối không đẩy lên GitHub công khai).
- Gồm các cột: `sku`, `product_name`, `category`, `list_price`, `ctv_tier1_pct`, `ctv_tier2_pct`, `ctv_price_tier1`, `ctv_price_tier2`, `stock_qty`, `internal_note`.
- Khi cập nhật giá niêm yết ở Bảng công khai, dùng hàm `IMPORTRANGE` trên Google Sheets để tự động tính lại giá chiết khấu cho CTV theo thời gian thực.

---

## 4. Luồng Tự động hóa 0đ (Single Source of Truth)
```text
[Google Sheets Master (4 tab public)]
        │
        ├──► (IMPORTRANGE) ──► [Google Sheet CTV & Chiết khấu (Riêng tư)]
        │
        └──► (GitHub Actions / Script chạy 1 chạm)
                    ├──► scripts/enrich_catalog.ps1  ──► Kéo Ảnh bìa + Tracklist + Năm + Audio 30s
                    │                                ──► Xuất báo cáo thiếu: data/missing_media_report.csv
                    └──► scripts/build_site_data.ps1 ──► Xuất catalog.json (4 nhóm hàng)
                                                     ──► Xuất llms.txt (B2A chuẩn mới)
                                                     ──► Sinh trang chi tiết /p/<slug>/index.html + sitemap.xml
```
