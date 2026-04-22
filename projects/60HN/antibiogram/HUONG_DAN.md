# HƯỚNG DẪN SỬ DỤNG — Bộ công cụ làm sạch dữ liệu iSharp AMS
===========================================================================

## CẤU TRÚC THƯ MỤC

```
project/
├── app_whonet.R        → Ứng dụng 1: Làm sạch dữ liệu WHONET phiên giải
├── app_danh_muc.R      → Ứng dụng 2: Cập nhật Danh mục tham chiếu
├── run_whonet.bat      → Nhấn đúp để chạy Ứng dụng 1
├── run_danh_muc.bat    → Nhấn đúp để chạy Ứng dụng 2
├── renv.lock           → Danh sách gói R (dùng chung cho cả 2 ứng dụng)
├── renv/               → Thư mục renv (không xóa)
└── HUONG_DAN.md        → File này
```

===========================================================================

## YÊU CẦU HỆ THỐNG

- Hệ điều hành: Windows 10 trở lên
- R: phiên bản 4.2.0 trở lên — tải tại https://cran.r-project.org
- Không bắt buộc cài RStudio, nhưng khuyến khích

> Lần đầu chạy: chương trình sẽ tự động cài các gói R cần thiết
> thông qua renv. Quá trình này có thể mất 5–15 phút tùy tốc độ mạng.
> Các lần sau sẽ khởi động nhanh hơn.

===========================================================================

## ỨNG DỤNG 1: LÀM SẠCH DỮ LIỆU WHONET PHIÊN GIẢI

Khởi chạy: nhấn đúp vào `run_whonet.bat`

### Mục đích
Đọc một hoặc nhiều file WHONET xuất từ phần mềm WHONET (.xlsx), ghép
với danh mục vi sinh vật và danh mục cơ chế kháng, rồi xuất ra một
file dữ liệu dài (long-format) đã chuẩn hóa, sẵn sàng để phân tích.

### Các bước sử dụng

**Bước 1 — Chọn file đầu vào**
- Nhấn "📂 Chọn file(s) WHONET" → chọn một hoặc nhiều file .xlsx xuất
  từ WHONET (có thể chọn nhiều file cùng lúc bằng Ctrl+Click).
- Nhấn "📘 Chọn Danh mục VSV" → chọn file danh mục vi sinh vật.
  File phải có các cột: `ma_hoa`, `ten_vsv`.
- Nhấn "📙 Chọn Danh mục cơ chế kháng" → chọn file danh mục cơ chế kháng.
  File phải có các cột: `ten_vsv`, `khang_sinh`, `ket_qua_ksd`.
  (Có thể bỏ qua nếu chưa có danh mục cơ chế kháng.)

**Bước 2 — Xử lý**
- Nhấn "▶ Xử lý dữ liệu" và chờ thanh tiến trình hoàn tất.
- Kết quả xem trước hiện ở bảng bên phải (100 dòng đầu).
- Tổng quan cho biết số dòng, số cột và khoảng năm của dữ liệu.

**Bước 3 — Xuất file**
- Nhấn "💾 Xuất file đã làm sạch..." → hộp thoại Save-As của Windows mở ra.
- Chọn thư mục và tên file muốn lưu → nhấn Save.
- Tên file mặc định gợi ý: `whonet_phien_giai_da_lam_sach_YYYY_YYYY.xlsx`

### Định dạng file đầu ra
File đầu ra là dữ liệu dạng dài (long-format), mỗi dòng đại diện cho
một kết quả kháng sinh đơn của một phân lập:

| Cột             | Mô tả                                      |
|-----------------|--------------------------------------------|
| ma_bn           | Mã bệnh nhân                               |
| ho_dem, ten_bn  | Họ và tên bệnh nhân                        |
| gioi_tinh       | Giới tính                                  |
| tuoi            | Tuổi                                       |
| ma_khoa         | Mã khoa lâm sàng                           |
| ngay_nhap_vien  | Ngày nhập viện                             |
| ngay_nuoi_cay   | Ngày lấy bệnh phẩm nuôi cấy               |
| nam_nuoi_cay    | Năm nuôi cấy (trích từ ngày)               |
| ten_benh_pham   | Loại bệnh phẩm                             |
| ma_vsv          | Mã vi khuẩn (từ WHONET, viết thường)       |
| ten_vsv         | Tên vi sinh vật (từ Danh mục VSV)          |
| khang_sinh      | Tên kháng sinh (viết hoa, bỏ hậu tố _nd)  |
| kq_ksd          | Kết quả KSD: S / R / I hoặc giá trị MIC   |
| source_file     | Tên file nguồn                             |

### Lưu ý kỹ thuật
- Chương trình tự động tìm dòng tiêu đề trong 20 dòng đầu của file WHONET.
- Phân lập trùng lặp (cùng bệnh nhân, vi khuẩn, loại bệnh phẩm, năm)
  được loại bỏ — chỉ giữ lần nuôi cấy sớm nhất.
- Kết quả SIR (S/R/I) và giá trị MIC được ép về kiểu chuỗi thống nhất.

===========================================================================

## ỨNG DỤNG 2: CẬP NHẬT DANH MỤC THAM CHIẾU

Khởi chạy: nhấn đúp vào `run_danh_muc.bat`

### Mục đích
Phát hiện các mã chưa được chuẩn hóa trong ba danh mục tham chiếu
(Vi sinh vật, Tên khoa, Tên bệnh phẩm), cho phép điền thông tin còn
thiếu trực tiếp trên giao diện, rồi lưu đè file danh mục gốc.

Ứng dụng gồm 3 tab, mỗi tab xử lý một danh mục riêng.

---

### TAB "Tên VSV" — Danh mục vi sinh vật

**File danh mục cần có các cột:**

| Cột          | Mô tả                              |
|--------------|------------------------------------|
| `ma_hoa`     | Mã vi khuẩn trong WHONET (khóa)   |
| `ten_vsv`    | Tên vi sinh vật đầy đủ            |
| `loai_vsv`   | Loại vi sinh vật (Gram+, Gram-, …) |
| `ten_viet_tat` | Tên viết tắt                   |

**Quy trình:**
1. Chọn file(s) WHONET + file Danh mục VSV → nhấn **Kiểm tra**
2. Bảng mã chưa chuẩn hóa xuất hiện → nhấn vào ô để sửa trực tiếp
3. Nhấn **Áp dụng cập nhật** → xem trước kết quả ở bảng bên phải
4. Nhấn **Lưu đè** → hộp thoại mở ra, chọn đúng tên file danh mục gốc

---

### TAB "Tên khoa" — Danh mục tên khoa lâm sàng

**File danh mục cần có các cột:**

| Cột             | Mô tả                           |
|-----------------|---------------------------------|
| `ma_hoa`        | Mã khoa trong WHONET (khóa)    |
| `ten_khoa`      | Tên khoa đầy đủ                |
| `ten_khoa_nhom` | Nhóm khoa (Nội, Ngoại, ICU, …) |

**Quy trình:** tương tự tab Tên VSV (trích xuất từ cột `Location`)

---

### TAB "Tên bệnh phẩm" — Danh mục loại bệnh phẩm

**File danh mục cần có các cột:**

| Cột                  | Mô tả                              |
|----------------------|------------------------------------|
| `ma_hoa`             | Mã bệnh phẩm trong WHONET (khóa)  |
| `ten_benh_pham`      | Tên bệnh phẩm đầy đủ             |
| `ten_benh_pham_nhom` | Nhóm bệnh phẩm (Máu, Nước tiểu, …)|

**Quy trình:** tương tự tab Tên VSV (trích xuất từ cột `Specimen type`)

---

### Lưu ý quan trọng khi lưu đè danh mục

⚠️ Khi hộp thoại Save-As mở ra, **bắt buộc phải chọn đúng tên file
danh mục gốc** (tên file gợi ý được điền sẵn). Nếu đổi tên, chương
trình sẽ từ chối lưu và hiển thị cảnh báo — đây là cơ chế bảo vệ
để tránh ghi nhầm file.

Cơ chế so khớp mã (`ma_hoa`) không phân biệt chữ hoa/thường và dấu
tiếng Việt, giúp tránh trùng lặp khi WHONET xuất mã không nhất quán.

===========================================================================

## XỬ LÝ SỰ CỐ THƯỜNG GẶP

| Sự cố | Nguyên nhân | Cách xử lý |
|-------|-------------|------------|
| Cửa sổ đen xuất hiện rồi đóng ngay | R chưa được cài hoặc renv.lock thiếu | Cài R tại cran.r-project.org, kiểm tra thư mục dự án |
| "Không file nào có cột Organism" | File WHONET không đúng định dạng xuất chuẩn | Xuất lại từ phần mềm WHONET, chọn định dạng .xlsx |
| Lỗi "Thiếu cột ma_hoa" | File danh mục không đúng cấu trúc | Kiểm tra tên cột trong file Excel (không có dấu cách thừa) |
| Hộp thoại Save-As không xuất hiện | PowerShell bị chặn bởi Group Policy | Liên hệ quản trị IT, hoặc lưu thủ công qua R console |
| Dữ liệu bị mất sau khi đóng app | Chưa nhấn Lưu đè trước khi đóng | Luôn lưu file trước khi đóng trình duyệt |
| Lần đầu chạy rất chậm | renv đang cài gói từ internet | Chờ quá trình hoàn tất, không đóng cửa sổ |

===========================================================================

## THÔNG TIN KỸ THUẬT

- Ngôn ngữ: R + Shiny
- Quản lý gói: renv (file renv.lock khóa phiên bản cụ thể)
- Giao diện: chạy trên trình duyệt web mặc định (localhost)
- Lưu file: sử dụng PowerShell SaveFileDialog (chỉ hỗ trợ Windows)
- So khớp danh mục: normalize_key() — bỏ dấu, viết thường, trim

**Các gói R chính được sử dụng:**

| Gói       | Mục đích                                     |
|-----------|----------------------------------------------|
| shiny     | Framework giao diện web tương tác            |
| DT        | Bảng dữ liệu có thể sửa trực tiếp           |
| dplyr     | Xử lý và biến đổi data frame                |
| tidyr     | Pivot dữ liệu sang dạng long-format          |
| openxlsx  | Đọc và ghi file Excel (.xlsx)                |
| stringi   | Chuẩn hóa chuỗi, bỏ dấu tiếng Việt         |
| janitor   | Chuẩn hóa tên cột (clean_names)             |
| writexl   | Ghi file Excel nhẹ hơn                      |
| tibble    | Tạo data frame nhanh                        |

===========================================================================

Phiên bản tài liệu: tháng 4/2026 · iSharp AMS · OUCRU
