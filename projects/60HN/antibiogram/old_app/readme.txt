================================================================================
 Chuẩn hóa Danh mục Vi sinh vật (Antibiogram Organism Name Standardizer)
 Dự án: 60HN 
================================================================================
CHỨC NĂNG
---------
Ứng dụng Shiny chạy cục bộ, dùng để chuẩn hóa tên vi sinh vật trong các file
Excel dựa trên danh mục tham chiếu (danh_muc_vsv.xlsx). Ứng dụng đánh dấu các
tên không nhận dạng được, cho phép xem xét, chỉnh sửa và xuất dữ liệu đã làm sạch.

CÁCH CHẠY ỨNG DỤNG
-------------------
1. Đảm bảo đã cài R (>= 4.4.0): https://cran.r-project.org
2. Nhấp đúp vào run_app.bat — tệp này sẽ tự động:
   - Tìm Rscript tự động
   - Khôi phục tất cả các thư viện (packages) cần thiết từ renv.lock
     (chỉ lần đầu chạy, mất vài phút)
   - Khởi chạy ứng dụng trên trình duyệt mặc định

CẤU TRÚC TỆP
-------------
  app.R                  Ứng dụng Shiny chính
  setup_and_run.R        Được gọi bởi run_app.bat — xử lý khôi phục renv + khởi chạy
  run_app.bat            Trình khởi chạy bằng nhấp đúp (Windows)
  renv.lock              Phiên bản của các thư viện (packages) đã ghim (không chỉnh sửa thủ công)
  renv/                  Nội bộ renv (không chỉnh sửa thủ công)

YÊU CẦU HỆ THỐNG
-----------------
  - Hệ điều hành Windows
  - R >= 4.4.0
  - Kết nối Internet ở lần chạy đầu tiên (để cài đặt gói)

LƯU Ý
------
  - Ở lần chạy đầu, renv::restore() sẽ cài 60+ thư viện (packages), vui long đợi.
  - Các lần chạy sau sẽ bỏ qua bước khôi phục và khởi chạy ngay.
  - Không di chuyển app.R hoặc renv.lock ra khỏi thư mục này.
================================================================================