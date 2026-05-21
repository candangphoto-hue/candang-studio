# Can Đăng Studio — Workspace

Bộ công cụ HTML cho Can Đăng Studio, hoạt động trực tiếp trên trình duyệt, không cần cài đặt.

🌐 **Live:** https://candangphoto-hue.github.io/candang-studio/

---

## Công cụ

| Tool | URL | Mô tả |
|------|-----|--------|
| 🏠 Launcher | `/` | Trang chủ — điều hướng tất cả tool |
| 📊 Studio Dashboard | `/dashboard/` | Quản lý dự án, khách hàng, doanh thu |
| 🎬 Hợp đồng Quay phim | `/hop-dong-quay-phim/` | Tạo bộ 3 văn bản hợp đồng quay phim |
| 📄 Bộ Văn bản v2 | `/hop-dong-v2/` | Hợp đồng chụp ảnh (cơ bản, sản phẩm, tổng hợp, cao cấp) |
| 💰 Báo giá Studio | `/bao-gia/` | Tạo báo giá chụp ảnh xuất file Word |

---

## Cấu trúc

```
candang-studio/
├── index.html                  ← Launcher (trang chủ)
├── dashboard/
│   └── index.html              ← Studio Dashboard
├── hop-dong-quay-phim/
│   └── index.html              ← Hợp đồng Quay phim
├── hop-dong-v2/
│   └── index.html              ← Bộ Văn bản v2
├── bao-gia/
│   └── index.html              ← Báo giá Studio
└── README.md
```

---

## Dùng trên mọi thiết bị

Tất cả tool chạy hoàn toàn trên client-side (HTML + JS thuần), không cần server hay backend.  
Các file `.docx` được tạo và tải về ngay trên trình duyệt.

---

*Can Đăng Studio © 2026*
