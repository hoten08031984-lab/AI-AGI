# CLAUDE.md — Project Context for AI AGI_2

## Thông tin dự án

- **Tên dự án**: AI AGI_2
- **Thư mục**: `D:\AI AGI_2`
- **GitHub**: https://github.com/hoten08031984-lab/AI-AGI
- **Branch mặc định**: `main`
- **Chủ sở hữu**: hoten08031984-lab

## Mô tả

Dự án web app mới được xây dựng trong thư mục `D:\AI AGI_2`.  
Đây là dự án độc lập, tách biệt với `D:\AI AGI`.

## Git

```bash
# Push lên GitHub
git add .
git commit -m "mô tả thay đổi"
git push

# Git path (nếu chưa có trong PATH)
"C:\Users\Hoang Tien\AppData\Local\Programs\Git\bin\git.exe"
```

## Quy tắc quan trọng

- **KHÔNG xóa** file nào khi không được yêu cầu
- **KHÔNG sửa** file nào khi không được yêu cầu  
- **KHÔNG đề xuất** ngoài phạm vi yêu cầu
- Luôn hỏi trước khi thực hiện thay đổi lớn

## Nguyên tắc thiết kế & Vận hành Dashboard

1. **Làm Mới & Đồng Bộ Dữ Liệu (Live Data Sync & Anti-Cache)**:
   - Dashboard có nút "Làm Mới Dữ Liệu" kết nối tới API `/api/refresh` ép Server đọc lại file Excel mới nhất lập tức.
   - Sử dụng query timestamp động (`Date.now()`) cho file script dữ liệu để chống cache trình duyệt 100%, đảm bảo dữ liệu mới nhất được hiển thị khi người dùng lưu file Excel và F5/làm mới.

2. **Khả Năng Chạy Trên Bất Kỳ Máy Tính Nào (Portability)**:
   - Mã nguồn (`server.py`, kịch bản khởi động) bắt buộc sử dụng **đường dẫn tương đối** (`BASE_DIR`), không hardcode đường dẫn ổ đĩa cố định.
   - Cung cấp sẵn file `CAI_DAT_TU_DONG.bat` để copy sang máy tính khác chỉ cần chạy 1 lần duy nhất là Server tự động đăng ký chạy ngầm cùng Windows.
   - Người dùng chỉ cần mở trình duyệt gõ `http://localhost:8080` là sử dụng bình thường.

3. **Cấu Trúc Bảng Danh Sách Chi Tiết**:
   - Ẩn các cột trùng lặp đã lọc ở thanh công cụ trên.
   - Hiển thị đúng 9 cột: `Tháng`, `Tiểu mục CP`, `Số HĐ`, `Ngày HĐ`, `Lý do thanh toán`, `Chi tiết HĐ`, `Số tiền chưa VAT`, `VAT`, `Số tiền VAT`.
   - Có dòng TỔNG CỘNG cố định ở chân bảng (`tfoot`) tính tổng 3 cột số tiền.

## Cấu trúc dự án

```
D:\AI AGI_2\
├── CLAUDE.md              ← file này (context cho Claude)
├── README.md              ← mô tả dự án
├── server.py              ← Python Web Server & tự động đọc Excel
├── index.html             ← Giao diện Dashboard HTML5
├── index.css              ← Hệ thống giao diện CSS (Dark Mode, Neon Badges)
├── app.js                 ← Logic lọc, ma trận, biểu đồ, bảng chi tiết
├── dashboard_data.js      ← Data cache JS (window.RAW_DATA)
├── dashboard_data.json    ← Data cache JSON
├── CAI_DAT_TU_DONG.bat    ← File 1-Click cài đặt chạy ngầm cho máy tính mới
├── RUN_DASHBOARD.bat      ← Batch runner thủ công
└── start_server_silent.vbs← Script chạy ngầm Server không hiện cửa sổ đen
```

## Môi trường

- **OS**: Windows
- **Shell**: PowerShell
- **IDE**: VS Code + Antigravity
- **Git**: `C:\Users\Hoang Tien\AppData\Local\Programs\Git\bin\git.exe`
- **Node/npm**: kiểm tra bằng `node -v` và `npm -v`
