# Hướng Dẫn Chi Tiết Tìm & Sửa Dữ Liệu Sản Phẩm Bằng Excel (Hue Sound Express)

## 1. Tại sao bạn không thấy đuôi `.xlsx` trong thư mục `data`?
Dữ liệu web của Hue Sound Express được lưu ở định dạng **`.csv`** (*Comma Separated Values* — mở và chỉnh sửa bình thường bằng **Microsoft Excel**, nhưng nhẹ và cho phép PowerShell tự động build web 0đ).

Trên Windows của bạn đang ẩn đuôi file nên chúng hiển thị như sau:
1. **Đĩa Than (LP):** Nằm ngay trong `E:\Hue Sound Express\Hue_Sound_Express_GitHub_Bio_B2A\data\` → file tên **`vinyl_lp_enriched`** (cột Type ghi *Comma Separated Values*).
2. **Đĩa CD / SACD:** Mở tiếp thư mục vàng **`templates`** bên trong `data\` → file **`cd_sacd_template`**.
3. **Máy Hát / Mâm Than:** Trong `data\templates\` → file **`turntables_template`**.
4. **Phụ Kiện:** Trong `data\templates\` → file **`accessories_care_template`**.
5. **Chiết khấu CTV (Nội bộ – KHÔNG đẩy lên web):** Trong `data\templates\` → file **`ctv_discount_private_template`**.

---

## 2. Cách mở nhanh bằng Excel (Không lỗi dấu tiếng Việt)
* **Cách 1 (Chuột phải):** Bấm **chuột phải** vào file `vinyl_lp_enriched` (hoặc file trong `templates`) → chọn **Open with** → chọn **Excel**.
* **Cách 2 (Mở ngay từ PowerShell):** Mở PowerShell tại thư mục dự án và chạy 1 trong các lệnh sau:
  ```powershell
  # Mo file Dia Than bang Excel:
  Start-Process excel.exe -ArgumentList "`"data\vinyl_lp_enriched.csv`""

  # Mo thu muc templates chua CD, May hat, Phu kien:
  explorer.exe data\templates
  ```
* **Khi lưu lại trong Excel:** Chỉ cần bấm **Ctrl + S** (nếu Excel hỏi *"Keep using this format?"* / *"Có muốn giữ định dạng CSV UTF-8 không?"* → bấm **Yes**).

---

## 3. Ý nghĩa các cột cần sửa
* `artist` / `brand`: Ca sĩ / Nghệ sĩ / Thương hiệu máy hát.
* `title`: Tên album / model máy / tên phụ kiện.
* `label`: Hãng phát hành (VD: `Viết Tân Studio`, `Gia Định Audio`, `Hãng Đĩa Thời Đại`).
* `list_price`: **Giá niêm yết** (chỉ gõ số, VD `1750000`; nhập `0` để hiển thị **`Liên hệ`**).
* `edition`: **Tình trạng** — chỉ điền 1 trong 3 giá trị: `New Seal`, `NM`, `Liên hệ`.
* `status`: **Trạng thái** — chỉ điền 1 trong 2 giá trị: `Còn hàng`, `Liên hệ`.
* `cover_url`: Link ảnh bìa sản phẩm.
* `audio_url`: Link file nghe thử 30s (MP3/M4A).
* `tracklist`: Danh sách bài hát, ngăn cách bởi dấu `|`.

---

## 4. Cập nhật lên Website sau khi sửa xong (PowerShell)
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_site_data.ps1 -GenerateProductPages
git add .
git commit -m "chore: cap nhat du lieu san pham"
git push
```
