# Order Management Web API + UI

Hệ thống quản lý đơn hàng nến (candle orders) — hỗ trợ import Excel (bảng tiếng Trung), giao diện song ngữ **Tiếng Việt / English / 中文**, và REST API.

## Tính năng

- **Import Excel** (`.xlsx` / `.xls`) với header tiếng Trung
- Lưu trữ SQLite
- Giao diện web: upload, tìm kiếm, xem chi tiết
- Tạo, lưu và xem lại phiếu cân giả định theo mã hàng (vỏ cốc, rót lần 1/lần 2, ghi chú)
- Quản lý tài khoản và phân quyền Admin / Nhân viên / Chỉ xem
- Chuyển ngôn ngữ: VI / EN / 中文
- REST API đầy đủ (Swagger tại `/docs`)

## Tài khoản và phân quyền

Tài khoản admin hiện có trong `data/credentials.json` được tự động chuyển vào bảng `user_accounts` khi khởi động; mật khẩu hiện tại được giữ nguyên. Đăng nhập bằng tài khoản admin, sau đó chọn **Quản lý user** để tạo tài khoản, đổi vai trò, khóa/mở khóa hoặc đặt lại mật khẩu. Mật khẩu user mới cần ít nhất 8 ký tự.

| Vai trò | Quyền |
|---------|-------|
| Admin | Toàn quyền: quản lý user, nhập/sửa/xóa đơn, phiếu cân, backup và restore |
| Nhân viên | Xem dữ liệu, lưu/sửa phiếu cân thực tế và xuất Excel phiếu cân |
| Chỉ xem | Xem dữ liệu và xuất Excel phiếu cân |

Lưu phiếu cân giả định vẫn không cần đăng nhập. Tài khoản admin cần chạy SQL trong `supabase_schema.sql` nếu triển khai trên Supabase với quyền database không cho phép ứng dụng tự tạo bảng.

## Chạy nhanh

```bash
cd order_api
pip install fastapi uvicorn openpyxl sqlalchemy python-multipart jinja2
python main.py
```

Mở trình duyệt: **http://localhost:8000**

API docs: **http://localhost:8000/docs**

## API chính

| Method | Endpoint | Mô tả |
|--------|----------|-------|
| `POST` | `/api/import` | Upload file Excel (multipart form, field `file`) |
| `GET`  | `/api/orders` | Danh sách đơn (query: `q`, `order_number`, `skip`, `limit`) |
| `GET`  | `/api/orders/{id}` | Chi tiết 1 dòng |
| `GET`  | `/api/orders/by-number/{order_number}` | Tất cả dòng theo số đơn |
| `GET`  | `/api/scale-tickets/{item_code}` | Danh sách phiếu cân đã lưu theo mã hàng |
| `POST` | `/api/scale-tickets` | Lưu phiếu cân giả định |
| `DELETE` | `/api/scale-tickets/{id}` | Xóa phiếu cân (yêu cầu đăng nhập) |
| `GET` | `/api/users` | Danh sách user (admin) |
| `POST` | `/api/users` | Tạo user (admin) |
| `PUT` | `/api/users/{username}` | Đổi vai trò, trạng thái hoặc đặt lại mật khẩu (admin) |
| `DELETE` | `/api/orders` | Xóa toàn bộ |
| `DELETE` | `/api/orders/{id}` | Xóa 1 dòng |
| `GET`  | `/api/labels?lang=vi` | Nhãn trường theo ngôn ngữ |
| `GET`  | `/api/stats` | Thống kê |

## Cấu trúc thư mục

```
order_api/
├── main.py          # FastAPI app
├── models.py        # SQLAlchemy models
├── importer.py      # Parse Excel → DB
├── translations.py  # Labels VI/EN/ZH
├── templates/
│   └── index.html   # Giao diện web
├── static/          # (assets nếu cần)
└── data/
    └── orders.db    # SQLite (tự tạo)
```

## Ví dụ import bằng curl

```bash
curl -X POST "http://localhost:8000/api/import" \
  -F "file=@YN26-0240订单明细表.xlsx"
```

## Ghi chú

- Ngày Excel (serial number) được chuyển sang `YYYY-MM-DD`
- Các giá trị `/` hoặc trống được coi là `null`
- Hỗ trợ header có xuống dòng (ví dụ `香精比例`, `留样要求`)
