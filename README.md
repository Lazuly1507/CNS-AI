# Portfolio Năng lực số và Trí tuệ nhân tạo

Website portfolio của Nguyễn Ngọc Hiếu cho học phần **Nhập môn Công nghệ số và Ứng dụng Trí tuệ nhân tạo** tại VNU-UET.

## Cấu trúc website

- `index.html`: trang chủ gốc, giữ nguyên ngoại hình của bản đã xuất PDF.
- `styles.css`, `script.js`: giao diện và hành vi của trang chủ.
- `pages/b1.html` đến `pages/b7.html`: bảy trang case study chi tiết.
- Mỗi trang B1–B7 có bản báo cáo trực tuyến đầy đủ; ảnh minh chứng được đặt ngay cạnh nội dung liên quan, kèm bảng dữ liệu và PDF đính kèm.
- `pages/tong-ket.html`: cảm nghĩ và tổng kết học phần.
- `portfolio.css`, `portfolio.js`: giao diện, chuyển cảnh và hiệu ứng của các trang chi tiết.
- `assets/`: ảnh đại diện và ảnh minh chứng.
- `deliverables/Google-Drive-PDF/`: bảy báo cáo PDF đi kèm.
- `reports/`: các báo cáo DOCX nguồn để tiếp tục chỉnh sửa.
- `site_builder.py`: sinh lại tám trang chi tiết từ dữ liệu nội dung.

## Chạy thử trên máy

Tại thư mục dự án:

```powershell
py -m http.server 8000
```

Sau đó mở `http://localhost:8000`.

## Xuất bản bằng GitHub Pages

1. Đẩy toàn bộ mã nguồn lên nhánh `main` của repository GitHub.
2. Mở `Settings` → `Pages`.
3. Trong `Build and deployment`, chọn `Deploy from a branch`.
4. Chọn nhánh `main`, thư mục `/root`, sau đó nhấn `Save`.
5. Website sẽ được xuất bản tại:

```text
https://lazuly1507.github.io/CNS-AI/
```

PDF được phục vụ trực tiếp từ repository và mở trong tab mới. Không cần Google Sites hoặc Google Drive.

## Cập nhật nội dung trang chi tiết

Sửa dữ liệu trong `site_builder.py`, sau đó chạy:

```powershell
py -m pip install -r requirements.txt
py site_builder.py
```

Lệnh này chỉ tạo lại `pages/b1.html` đến `pages/b7.html` và `pages/tong-ket.html`; `index.html` không bị thay đổi.
