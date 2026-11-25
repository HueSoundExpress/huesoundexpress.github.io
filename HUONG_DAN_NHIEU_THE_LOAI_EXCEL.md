# Hướng Dẫn Gán 1 Sản Phẩm Vào Nhiều Thể Loại Trong Excel (Hue Sound Express)

Từ bản cập nhật mới, hệ thống dữ liệu CSV của **Hue Sound Express** đã bổ sung cột **`extra_genres`** (Thể loại bổ sung) ngay sau cột **`genre`** (Thể loại chính) trong cả 2 file:
- Đĩa than (Vinyl LP): `data/vinyl_lp_enriched.csv`
- Đĩa CD / SACD: `data/templates/cd_sacd_template.csv`

Nhờ đó, bạn chỉ cần nhập **1 dòng sản phẩm duy nhất** trong Excel nhưng sản phẩm đó vẫn xuất hiện đầy đủ ở **2, 3 hoặc nhiều thể loại khác nhau** trên Phòng Nghe Số và trên trang chi tiết `/p/.../` (mỗi thể loại đều có thẻ bấm lọc chéo riêng).

---

## 1. Cách nhập trong Excel (Cực kỳ đơn giản)

| Cột trong Excel | Ý nghĩa | Ví dụ (2 Cellos – Score) | Ví dụ (Đĩa nhạc Trịnh hòa tấu Piano) |
| :--- | :--- | :--- | :--- |
| **`genre`** | Thể loại **chính** (hiển thị đầu tiên trong breadcrumb) | `classic_soundtrack` | `love_songs_1954_1975` |
| **`extra_genres`** | Các thể loại **phụ / bổ sung**, ngăn cách bằng dấu chấm phẩy `;` | `instrumental` | `instrumental; audiophile_reference` |

> **Mẹo tiện lợi:** Ở cả cột `genre` và `extra_genres`, bạn có thể nhập **Mã tiếng Anh chuẩn** HOẶC **Tên tiếng Việt hiển thị trên web** đều được — bộ build sẽ tự động nhận diện và chuẩn hóa!
> - Ví dụ: gõ `Instrument` hay `instrumental` đều cho kết quả giống nhau.
> - Ví dụ: gõ `Classic & OST; Instrument` hoặc `classic_soundtrack; instrumental` đều hoạt động chính xác.

---

## 2. Bảng Tra Cứu Nhanh Mã Thể Loại Chuẩn

Bạn cũng có thể mở file tham chiếu sẵn bằng Excel tại `data/templates/bang_ma_the_loai_excel.csv`.

| Phân loại | Tên hiển thị trên Website | Mã chuẩn (`genre` / `extra_genres`) | Gõ nhanh tiếng Việt được hỗ trợ |
| :--- | :--- | :--- | :--- |
| Nhạc Việt | Tình khúc 1954–1975 | `love_songs_1954_1975` | `Tình khúc 1954-1975`, `Nhạc xưa`, `Tiền chiến` |
| Nhạc Việt | Nhạc Việt / Bolero (Trữ tình) | `bolero_golden_oldies` | `Bolero`, `Nhạc vàng`, `Trữ tình` |
| Nhạc Việt | Nhạc Việt trẻ | `vietnamese_pop_modern` | `Nhạc Việt trẻ`, `Nhạc Việt`, `V-Pop` |
| Nhạc Việt | Dân ca & Cách mạng | `vietnamese_folk_patriotic` | `Dân ca & Cách mạng`, `Dân ca`, `Quê hương` |
| Chung / Quốc tế | Instrument (Hòa tấu) | `instrumental` | `Instrument`, `Hòa tấu`, `Nhạc cụ` |
| Quốc tế | Jazz & Vocal | `jazz_vocal` | `Jazz & Vocal`, `Jazz`, `Vocal` |
| Quốc tế | Audiophile & Khác | `audiophile_reference` | `Audiophile`, `Audiophile & Khác` |
| Quốc tế | Rock & Metal | `rock_metal` | `Rock & Metal`, `Rock`, `Metal` |
| Quốc tế | Pop & Disco | `pop_disco` | `Pop & Disco`, `Pop`, `Disco` |
| Quốc tế | Classic & OST (Nhạc phim) | `classic_soundtrack` | `Classic & OST`, `Classic & Soundtrack`, `OST`, `Nhạc phim`, `Cổ điển` |

---

## 3. Quy Trình 3 Bước Cập Nhật Bằng Excel & Đồng Bộ Lên Web

1. **Mở file CSV bằng Excel** (khuyến nghị lưu dưới định dạng **CSV UTF-8 (Comma delimited) (*.csv)** để không lỗi dấu tiếng Việt).
2. Điền thể loại chính vào cột **`genre`**, điền thêm các thể loại phụ vào cột **`extra_genres`** (nếu nhiều hơn 1 thể loại phụ thì ngăn cách bằng dấu `;`, ví dụ: `instrumental; jazz_vocal`). Lưu file lại.
3. Mở **Windows PowerShell** tại thư mục dự án và chạy lệnh 1 dòng quen thuộc:

```powershell
powershell -ExecutionPolicy Bypass -File .\dong-bo.ps1
```

Hệ thống sẽ tự động:
- Làm sạch lỗi định dạng từ Excel (ngoặc kép thừa, khoảng trắng).
- Đưa sản phẩm xuất hiện ở **tất cả** các bộ lọc thể loại bạn đã gán.
- Tự động tạo đủ các badge thể loại bấm được trên trang `/p/<slug>/`.
- Tự động đưa sản phẩm vào các **Bộ Sưu Tập** tương ứng (ví dụ: thêm `instrumental` là đĩa tự xuất hiện thêm trong *Bộ sưu tập Instrument & Hòa tấu nhạc cụ*!).

---

## 4. Lưu Ý Khi Lưu File Từ Excel
- Không xóa hoặc đổi tên hàng tiêu đề đầu tiên (`id`, `sku`, `genre`, `extra_genres`, ...).
- Nếu 1 sản phẩm chỉ thuộc **1 thể loại duy nhất**, cứ **để trống** ô `extra_genres`.
- Tuyệt đối không dùng dấu phẩy `,` để định dạng tiền ở cột `list_price` (nhập số liền, ví dụ `1350000`; nếu lỡ gõ dấu phẩy, script `kiem-tra-va-sua-csv.ps1` vẫn tự động làm sạch hộ bạn).
