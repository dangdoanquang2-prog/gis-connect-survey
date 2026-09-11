# 📋 Hệ Thống Khảo Sát Thực Địa & Thiết Kế Biểu Mẫu (FormsMobile - GIS Connect)

Ứng dụng khảo sát thực địa di động (PWA) và thiết kế biểu mẫu trực quan, hỗ trợ hoạt động offline, định vị GPS, chụp ảnh hiện trường, đồng bộ dữ liệu đám mây (Supabase) và tự động hóa xuất báo cáo.

---

## 🌟 Tính Năng Nổi Bật

- **Thiết Kế Biểu Mẫu Trực Quan (Form Builder):**
  - Đầy đủ các dạng câu hỏi: Trắc nghiệm, Hộp kiểm, Dropdown, Thang đo, Lưới ma trận, Ngày & Giờ.
  - Hỗ trợ Logic rẽ nhánh (Conditional Skip Routing), Xáo trộn đáp án.
  - Phân trang nhiều phần (Multi-section) với thanh công cụ điều hướng nổi (Floating Toolbar).
- **Thu Thập Dữ Liệu Hiện Trường (Field Data Collection):**
  - Tự động bắt tọa độ GPS với độ chính xác cao.
  - Chụp ảnh hiện trường trực tiếp từ Camera, tự động gắn với mã điểm và phiếu khảo sát.
  - Ghi âm bằng chứng khảo sát (AI Voice Survey).
  - Hoạt động 100% Offline qua Service Worker, tự động đồng bộ khi có kết nối mạng.
- **Lưu Trữ & Đồng Bộ Đám Mây:**
  - Kết nối Supabase Cloud (PostgreSQL) theo thời gian thực.
  - Hỗ trợ LocalStorage dự phòng an toàn.
- **Phân Tích & Xuất Báo Cáo:**
  - Tab thống kê câu trả lời với biểu đồ trực quan.
  - Xuất dữ liệu ra Excel (.xlsx), CSV và liên kết tự động với mẫu báo cáo Word/PDF.

---

## 🚀 Hướng Dẫn Cài Đặt & Chạy Cục Bộ

### 1. Yêu Cầu Hệ Thống
- Python 3.8+ hoặc bất kỳ web server tĩnh nào (Node.js, Live Server...).
- Trình duyệt hiện đại (Chrome, Edge, Safari, Firefox).

### 2. Khởi Động Ứng Dụng
Chạy file batch có sẵn:
```cmd
CHAY_APP.bat
```
Hoặc khởi động thủ công qua Python:
```bash
python server.py
```
Mở trình duyệt và truy cập:
👉 **`http://localhost:8080`**

---

## 📂 Cấu Trúc Thư Mục

```text
gis_connect/
├── css/                  # Giao diện & Tailwind styles
├── js/
│   ├── app.js            # Logic giao diện, form builder & khảo sát
│   └── store.js          # Quản lý dữ liệu, LocalStorage & Supabase Cloud
├── index.html            # Giao diện ứng dụng chính
├── manifest.json         # Cấu hình PWA (cài đặt ra màn hình điện thoại)
├── sw.js                 # Service Worker hỗ trợ chạy Offline
├── server.py             # HTTP Server không lưu cache (No-cache)
└── CHAY_APP.bat          # File bấm chạy nhanh trên Windows
```

---

## ☁️ Đồng Bộ Supabase Cloud

Hệ thống được tích hợp sẵn với Supabase:
- Bảng `surveys`: Lưu trữ cấu trúc câu hỏi và cài đặt biểu mẫu.
- Bảng `responses`: Lưu trữ kết quả khảo sát từ hiện trường.

---

## 📱 Cài Đặt Trên Điện Thoại (PWA)
1. Mở liên kết ứng dụng trên trình duyệt Chrome (Android) hoặc Safari (iOS).
2. Nhấn vào nút menu hoặc biểu tượng chia sẻ -> Chọn **"Thêm vào màn hình chính" (Add to Home Screen)**.
3. Ứng dụng sẽ xuất hiện trên màn hình điện thoại với biểu tượng riêng và chạy toàn màn hình (Standalone).
