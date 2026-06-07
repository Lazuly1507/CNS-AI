# Portfolio Năng lực số và Trí tuệ nhân tạo

Google Sites không cho tải trực tiếp một thư mục HTML để chạy như website. Cách ổn định nhất là xuất bản thư mục này bằng một dịch vụ hosting tĩnh, sau đó nhúng URL vào Google Sites.

## Cấu trúc thư mục

- `index.html`
- `styles.css`
- `script.js`
- Thư mục `assets`
- `reports/`: bảy báo cáo DOCX nguồn để tiếp tục chỉnh sửa
- `documents/`: đề bài, rubric và tài liệu tham khảo
- `archive/`: bản sao lưu, không cần đưa lên hosting
- `deliverables/`: gói ZIP hoàn chỉnh để bàn giao hoặc tải lên hosting
- `deliverables/Google-Drive-PDF/`: bảy báo cáo PDF cuối để tải lên Google Drive
- `tools/`: công cụ tạo, chuẩn hóa và xuất PDF báo cáo

## Ảnh đại diện

Lưu ảnh chân dung với tên `profile.jpg` tại đường dẫn:

```text
assets/profile.jpg
```

Nếu chưa có ảnh, website sẽ tự hiển thị chữ viết tắt `NH`.

## Cách xuất bản đề xuất

1. Tải bảy PDF trong `deliverables/Google-Drive-PDF/` lên một thư mục Google Drive.
2. Đặt quyền thư mục thành `Bất kỳ ai có đường liên kết` → `Người xem`.
3. Thay các liên kết báo cáo trong `index.html` bằng liên kết xem Google Drive của từng PDF.
4. Tạo repository GitHub và tải `index.html`, `styles.css`, `script.js`, `assets/` và `deliverables/Google-Drive-PDF/` lên. Có thể bỏ thư mục PDF sau khi đã thay đủ liên kết bằng URL Google Drive.
5. Vào `Settings` → `Pages`.
6. Chọn `Deploy from a branch`, nhánh `main`, thư mục `/root`.
7. Sau khi có URL GitHub Pages, mở Google Sites.
8. Chọn `Insert` → `Embed` → `By URL`, dán URL GitHub Pages.
9. Kéo khung nhúng đủ rộng và xuất bản Google Site.

Có thể mở trực tiếp `index.html` để xem thử trên máy tính trước khi xuất bản.

Chi tiết xem tại `HUONG-DAN-GOOGLE-SITES.md`.
