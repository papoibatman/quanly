# Hệ Thống Quản Lý Nhân Sự (quanly)

Ứng dụng web quản lý nhân sự được xây dựng bằng Python Flask và SQLite.

## Tính năng chính
- Quản trị nhân viên (Thêm, sửa, xóa, tìm kiếm nhân viên)
- Quản lý công việc (Tasks)
- Lịch nghỉ phép (Leave calendar)
- Bảng điều khiển quản trị (Admin Dashboard) & Giao diện người dùng (User Dashboard)
- Đăng nhập & phân quyền người dùng

## Hướng dẫn cài đặt & chạy ứng dụng

1. **Cài đặt thư viện phụ thuộc:**
   ```bash
   pip install flask flask-sqlalchemy
   ```

2. **Khởi chạy ứng dụng:**
   ```bash
   python app.py
   ```

3. **Truy cập ứng dụng:**
   Mở trình duyệt và truy cập: `http://127.0.0.1:5000`
