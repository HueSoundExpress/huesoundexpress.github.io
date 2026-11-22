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
| `list_price` | Giá niêm yết (VND) | Chỉ gõ số nguyên (VD: `1200000`). **Website luôn hiển thị giá** (kể cả khi `Hết hàng`) để khách tham khảo. Chỉ điền **`0`** nếu muốn ẩn giá và hiển thị **`Liên hệ`**. |
| `status` | Trạng thái kho | `in_stock` (*Còn hàng*) hoặc `out_of_stock` (*Hết hàng* → **luôn hiện nút `🔔 Báo khi về`**, kể cả khi có giá). |
| `edition` | Tình trạng đĩa (Goldmine) | `mint_sealed` (*Mới 100% (M)*), `near_mint` (*NM*), `very_good_plus` (*VG+*), `very_good` (*VG*), `good_plus` (*G+*), `good` (*G*), `fair` (*F*), `poor` (*P*). |
| `genre` | Thể loại chính | **Đĩa than / CD:** `love_songs_1954_1975` (*Tình khúc 1954–1975*), `bolero_golden_oldies` (*Trữ tình / Bolero*), `vietnamese_pop_modern` (*Nhạc Việt trẻ*), `vietnamese_folk_patriotic` (*Dân ca & Cách mạng*), `instrumental` (*Instrument*), `jazz_vocal` (*Jazz & Vocal*), `audiophile_reference` (*Audiophile*), `rock_metal` (*Rock & Metal*), `pop_disco` (*Pop & Disco*), `classic_soundtrack` (*Classic & OST*).<br>**Mâm than:** `beginner`, `mid_range`, `audiophile`.<br>**Phụ kiện:** `record_cleaning`, `sleeves_protection`, `cartridge_stylus`, `calibration`, `vinyl_care_service`. |
| `mood` | Tâm trạng / Không gian (ngăn cách `;`) | `late_night_chill`, `nostalgic_classic`, `audiophile_reference`, `weekend_upbeat`, `refined_gifting`, `morning_coffee_acoustic`. |
| `contact_only` *(áp dụng cho cả Đĩa than, CD, Máy hát, Phụ kiện)* | Chế độ nút bấm khi **Còn hàng** (`in_stock`) | `0` (hoặc để trống): hiện nút **`+ Giỏ / Báo giá`** (thêm vào giỏ hàng).<br>`1`: hiện nút **`💬 Liên hệ đặt hàng`** (bật popup Zalo / WhatsApp).<br>*(Lưu ý: Khi `status = out_of_stock`, website luôn ưu tiên nút **`🔔 Báo khi về`**).* |
| `year` | Năm phát hành | Số năm 4 chữ số (VD: `2024`). |
| `cover_url` | Link ảnh bìa | URL `https://...` (tên file không dấu, nối bằng dấu gạch ngang `-`). |
| `audio_url` | Link nghe thử 30s | URL `https://...` (MP3/M4A). |
| `tracklist` | Danh sách bài hát | Các bài hát ngăn cách bởi dấu gạch đứng `|`. |

### Bảng 3 chế độ hiển thị nhanh (Áp dụng chung cho mọi file CSV)

| Mục tiêu hiển thị trên Website | `list_price` | `status` | `contact_only` | Kết quả hiển thị trên Web & Trang `/p/` |
| :--- | :--- | :--- | :---: | :--- |
| **1. Hết hàng → Báo khi hàng về (Luôn hiện giá)** | Giá niêm yết (VD: `1200000`) *(chỉ nhập `0` nếu muốn hiện "Liên hệ")* | `out_of_stock` | `0` hoặc `1` | Luôn hiện **Giá** (hoặc *Liên hệ* nếu = `0`) • Nhãn **`Hết hàng`** • Nút **`🔔 Báo khi về`** |
| **2. Còn hàng → Thêm vào giỏ / Báo giá** | Giá niêm yết (VD: `1200000` → tự gạch giá gốc & giảm 50k–60k) | `in_stock` | `0` | Hiện **Giá gạch ngang + Giá ưu đãi** • Nhãn **`Còn hàng`** • Nút **`+ Giỏ / Báo giá`** |
| **3. Còn hàng → Liên hệ đặt hàng** | Giá niêm yết (VD: `1200000`) | `in_stock` | `1` | Hiện **Giá gạch ngang + Giá ưu đãi** • Nhãn **`Còn hàng`** • Nút **`💬 Liên hệ đặt hàng`** |

> **Lưu ý:** Nếu đặt `list_price = 0` thì kể cả đang ghi `in_stock`, hệ thống vẫn tự chuyển trạng thái về **`Hết hàng` → `🔔 Báo khi về`** để tránh lỗi cộng tiền 0đ trong giỏ hàng/báo giá PDF.

---

## 4. Đồng bộ lên Website sau khi sửa CSV (PowerShell)

### Cách nhanh 1 lệnh duy nhất (Khuyên dùng – không lo gõ nhầm / dính lệnh)
Mở **Windows PowerShell**, chạy đúng 2 dòng sau:
```powershell
cd "E:\Hue Sound Express\Hue_Sound_Express_GitHub_Bio_B2A"
powershell -ExecutionPolicy Bypass -File .\dong-bo.ps1
```
Script [dong-bo.ps1](../dong-bo.ps1) sẽ tự động kiểm tra xem Excel đã đóng chưa → chạy [kiem-tra-va-sua-csv.ps1](../scripts/kiem-tra-va-sua-csv.ps1) để **tự động phát hiện & cắt sạch thẻ HTML rác / cảnh báo link ảnh dán nhầm / xóa slug lỗi** → chạy `build_site_data.ps1 -GenerateProductPages` → `git add -A` → `git commit` → `git push` lên GitHub Pages (tự xử lý cả trường hợp cần `git pull --rebase`).

### Hoặc chạy thủ công từng bước (nhớ bấm `Enter` sau mỗi dòng)

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
- [ ] Đã đóng Excel (và các trình mở file CSV/Markdown) trước khi chạy `build_site_data.ps1`.
- [ ] Lệnh `build_site_data.ps1 -GenerateProductPages` chạy xong không báo chữ đỏ (Error).
- [ ] Đã chạy `git add -A`, `git commit`, `git push` thành công.
- [ ] Mở `https://huesoundexpress.github.io/` (bấm **Ctrl + F5** xóa cache trình duyệt) và kiểm tra đúng 3 chế độ: **Hết hàng → hiện giá + `🔔 Báo khi về`**, **Còn hàng (`contact_only = 0`) → `+ Giỏ / Báo giá`**, **Còn hàng (`contact_only = 1`) → `💬 Liên hệ đặt hàng`**.
