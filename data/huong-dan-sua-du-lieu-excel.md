# Hướng Dẫn Sửa Dữ Liệu CSV Thủ Công & Đồng Bộ Website (Hue Sound Express)

Tài liệu này hướng dẫn quy trình chuẩn để tự chỉnh sửa danh mục sản phẩm (Đĩa than, CD, Máy hát, Phụ kiện) bằng **Microsoft Excel** và đồng bộ toàn bộ website, trang sản phẩm `/p/`, bộ sưu tập `/bo-suu-tap/`, `feed/products.json`, `llms.txt`, `sitemap.xml` lên GitHub Pages chỉ bằng **Windows PowerShell** có sẵn (0đ).

---

## 1. Vị trí 4 file CSV dữ liệu chính

Tất cả dữ liệu website đều nằm trong thư mục `data\` (lưu chuẩn **UTF-8 CSV**):

1. **Đĩa than (Vinyl LP):** `data\vinyl_lp_enriched.csv`
2. **Đĩa CD / SACD:** `data\templates\cd_sacd_template.csv`
3. **Máy hát / Mâm than:** `data\templates\turntables_template.csv`
4. **Phụ kiện & Dịch vụ (Hue Sound Vinyl Care):** `data\templates\accessories_care_template.csv`
5. *(Nội bộ – KHÔNG đẩy lên web)* Chiết khấu CTV: `data\templates\ctv_discount_private_template.csv`

---

## 2. Cách mở & lưu file CSV bằng Excel (Tránh lỗi font tiếng Việt)

### Mở nhanh từ PowerShell
Mở **Windows PowerShell** tại thư mục dự án rồi chạy:
```powershell
# Mở file Đĩa than bằng Excel
Start-Process excel.exe -ArgumentList "`"data\vinyl_lp_enriched.csv`""

# Hoặc mở file Đĩa CD bằng Excel
Start-Process excel.exe -ArgumentList "`"data\templates\cd_sacd_template.csv`""

# Hoặc mở thư mục chứa các file template
explorer.exe data\templates
```

### Quy tắc lưu file cực kỳ quan trọng
* Sau khi sửa xong, bấm **Ctrl + S**.
* Nếu Excel hỏi *"Keep using this format?"* → bấm **Yes**.
* Nếu lưu lần đầu / **Save As**, **bắt buộc** chọn định dạng **`CSV UTF-8 (Comma delimited) (*.csv)`** (không chọn `CSV (Comma delimited)` thường sẽ làm hỏng dấu tiếng Việt).
* **Đóng hẳn Excel** trước khi chạy lệnh đồng bộ bên dưới để tránh lỗi khóa file (*file in use*).

---

## 3. Quy ước giá trị nội bộ chuẩn `snake_case` (Không dấu, chữ thường)

> **Lưu ý:** Website tự động dịch các mã nội bộ bên dưới sang **Tiếng Việt / Tiếng Anh** trên giao diện. Khi nhập trong CSV, luôn điền **đúng mã tiếng Anh không dấu** để bộ lọc, AI Shopping Feed và Schema hoạt động chính xác 100%.

| Cột trong CSV | Ý nghĩa | Giá trị hợp lệ cần điền |
| :--- | :--- | :--- |
| `sku` | Mã sản phẩm (ASCII, không dấu) | VD: `VN-SSR-MOI-LUONG-DUYEN-LP`, `CD-VN-HUONGTRA-1001` *(tiền tố `VN-` / `CD-VN-` tự nhận Nhạc Việt, `NN-` tự nhận Nhạc Quốc Tế)* |
| `list_price` | Giá niêm yết (VND) | Chỉ gõ số nguyên (VD: `1200000`). Điền **`0`** để hiện **`Liên hệ`**. |
| `status` | Trạng thái kho | `in_stock` (*Còn hàng*) hoặc `out_of_stock` (*Hết hàng* → tự hiện nút **`🔔 Báo khi về`**). |
| `edition` | Tình trạng đĩa (Goldmine) | `mint_sealed` (*Mới 100% (M)*), `near_mint` (*NM*), `very_good_plus` (*VG+*), `very_good` (*VG*), `good_plus` (*G+*), `good` (*G*), `fair` (*F*), `poor` (*P*). |
| `genre` | Thể loại chính | **Đĩa than / CD:** `love_songs_1954_1975` (*Tình khúc 1954–1975*), `bolero_golden_oldies` (*Trữ tình / Bolero*), `vietnamese_pop_modern` (*Nhạc Việt trẻ*), `vietnamese_folk_patriotic` (*Dân ca & Cách mạng*), `instrumental` (*Instrument*), `jazz_vocal` (*Jazz & Vocal*), `audiophile_reference` (*Audiophile*), `rock_metal` (*Rock & Metal*), `pop_disco` (*Pop & Disco*), `classic_soundtrack` (*Classic & OST*).<br>**Mâm than:** `beginner`, `mid_range`, `audiophile`.<br>**Phụ kiện:** `record_cleaning`, `sleeves_protection`, `cartridge_stylus`, `calibration`, `vinyl_care_service`. |
| `mood` | Tâm trạng / Không gian (ngăn cách `;`) | `late_night_chill`, `nostalgic_classic`, `audiophile_reference`, `weekend_upbeat`, `refined_gifting`, `morning_coffee_acoustic`. |
| `contact_only` *(chỉ file CD)* | Chế độ nút đặt hàng | `0` (bình thường: còn hàng → Giỏ hàng; hết hàng → **`🔔 Báo khi về`**); `1` (ép hiện nút *Liên hệ đặt hàng* kể cả khi còn giá). |
| `year` | Năm phát hành | Số năm 4 chữ số (VD: `2024`). |
| `cover_url` | Link ảnh bìa | URL `https://...` (tên file không dấu, nối bằng dấu gạch ngang `-`). |
| `audio_url` | Link nghe thử 30s | URL `https://...` (MP3/M4A). |
| `tracklist` | Danh sách bài hát | Các bài hát ngăn cách bởi dấu gạch đứng `|`. |

### Mẹo chỉnh nhanh "Hết hàng – Giá Liên hệ – Báo khi về"
* Đặt `list_price` = `0`
* Đặt `status` = `out_of_stock`
* *(Với file CD)* đặt `contact_only` = `0`

### Mẹo bật lại "Còn hàng – Có giá – Thêm vào giỏ"
* Đặt `list_price` = giá niêm yết (VD: `1200000` → web tự hiển thị giá gạch ngang & giá ưu đãi giảm 50.000–60.000đ)
* Đặt `status` = `in_stock`
* *(Với file CD)* đặt `contact_only` = `0` (nếu muốn cho vào giỏ) hoặc `1` (nếu muốn hiện giá nhưng nút bấm là *Liên hệ đặt hàng*)

---

## 4. Quy trình 4 bước đồng bộ lên Website sau khi sửa CSV (PowerShell)

Mở **Windows PowerShell**, sao chép và chạy lần lượt các lệnh sau:

### Bước 1: Vào đúng thư mục dự án
```powershell
cd "E:\Hue Sound Express\Hue_Sound_Express_GitHub_Bio_B2A"
```

### Bước 2: Chạy script tự động đồng bộ toàn bộ dữ liệu & trang tĩnh
Lệnh này sẽ đọc lại 4 file CSV và tự động tái tạo đồng bộ:
- `data/catalog_all.json` (dữ liệu hiển thị trên trang chủ)
- Toàn bộ trang sản phẩm `/p/<slug>/index.html` + redirect URL cũ
- 10 trang bộ sưu tập `/bo-suu-tap/...` & 6 trang `/cam-nang/...`
- `feed/products.json` (AI Shopping Feed), `llms.txt` và `sitemap.xml`

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_site_data.ps1 -GenerateProductPages
```

### Bước 3: Kiểm tra nhanh trạng thái thay đổi (không có lỗi)
```powershell
git status -s
```
*(Nếu thấy `data/catalog_all.json`, `feed/products.json`, `sitemap.xml` và các file `/p/...` được cập nhật là chuẩn.)*

### Bước 4: Đẩy dữ liệu lên GitHub Pages (Live sau 1–2 phút)
```powershell
git add -A
git commit -m "chore: cap nhat du lieu san pham tu CSV"
git push
```

> **Xử lý khi `git push` báo `[rejected] (fetch first)`:**
> Chạy 2 lệnh sau để kéo bản mới nhất từ GitHub rồi đẩy lại:
> ```powershell
> git pull --rebase origin main
> git push
> ```

---

## 5. Checklist tự kiểm tra sau khi đồng bộ
- [ ] Đã đóng Excel trước khi chạy `build_site_data.ps1`.
- [ ] Lệnh `build_site_data.ps1 -GenerateProductPages` chạy xong không báo chữ đỏ (Error).
- [ ] Đã chạy `git add -A`, `git commit`, `git push` thành công.
- [ ] Mở `https://huesoundexpress.github.io/` (bấm **Ctrl + F5** xóa cache trình duyệt) và kiểm tra 1–2 sản phẩm vừa sửa đã hiển thị đúng giá, tình trạng và nút bấm.
