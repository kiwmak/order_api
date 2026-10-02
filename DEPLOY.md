# Deploy Order API → Koyeb + Supabase

## 1. Tạo project Supabase (database)

1. Vào [https://supabase.com](https://supabase.com) → **New project**
2. Đặt tên, mật khẩu database (ghi lại)
3. **Project Settings → Database → Connection string**
4. Chọn **URI** / **Session pooler** (port **5432**) hoặc **Transaction pooler** (port **6543**)

Ví dụ (Session pooler – khuyến nghị cho Koyeb):

```
postgresql://postgres.XXXXXXXX:[YOUR-PASSWORD]@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres
```

> Nếu mật khẩu có ký tự đặc biệt (`@ # %`), cần **URL-encode**.

App sẽ tự tạo bảng `order_items` khi khởi động (`init_db()`).

---

## 2. Đẩy code lên GitHub

Trong thư mục `order_api`:

```bash
git init
git add .
git commit -m "Order API ready for Koyeb + Supabase"
# Tạo repo trống trên GitHub rồi:
git remote add origin https://github.com/YOUR_USER/order-api.git
git branch -M main
git push -u origin main
```

---

## 3. Deploy trên Koyeb

1. Đăng nhập [https://app.koyeb.com](https://app.koyeb.com)
2. **Create Service** → **GitHub** → chọn repo `order-api`
3. **Builder**: **Dockerfile** (tự nhận `Dockerfile`)
4. **Port**: `8000` (HTTP)
5. **Environment variables** (Secrets):

| Name | Value | Ghi chú |
|------|--------|---------|
| `DATABASE_URL` | `postgresql://postgres.xxx:PASSWORD@...pooler.supabase.com:5432/postgres` | Connection string Supabase |
| `ORDER_ADMIN_USER` | `admin` | User đăng nhập (lần đầu) |
| `ORDER_ADMIN_PASS` | `mat-khau-manh` | Mật khẩu lần đầu |
| `ORDER_API_SECRET` | chuỗi ngẫu nhiên dài | Ký token JWT-like |

6. **Deploy**

Sau khi xanh, URL dạng: `https://your-app-xxxxx.koyeb.app`

---

## 4. Kiểm tra

- Mở `https://your-app.koyeb.app/` → Dashboard
- `https://your-app.koyeb.app/docs` → Swagger API
- Đăng nhập bằng `ORDER_ADMIN_USER` / `ORDER_ADMIN_PASS`
- Import Excel thử

Trong Supabase → **Table Editor** sẽ thấy bảng `order_items`.

---

## 5. Lưu ý quan trọng

### Ảnh sản phẩm
Disk trên Koyeb **không bền** (mất khi redeploy). Ảnh trong `static/uploads` có thể mất.  
Dữ liệu đơn hàng vẫn an toàn trên **Supabase Postgres**.  
Nếu cần ảnh bền vững: cấu hình Supabase Storage (bước mở rộng sau).

### Đổi mật khẩu
Sau khi đổi MK trên web, file `data/credentials.json` nằm trên disk container → có thể mất khi redeploy.  
Nên đặt `ORDER_ADMIN_PASS` mạnh trên Koyeb, hoặc sau này lưu hash vào Postgres.

### Local vẫn dùng SQLite
Không set `DATABASE_URL` → app dùng `data/orders.db` như cũ.

---

## 6. Biến môi trường tóm tắt

```env
DATABASE_URL=postgresql://postgres.PROJECT:PASSWORD@HOST:5432/postgres
ORDER_ADMIN_USER=admin
ORDER_ADMIN_PASS=your-strong-password
ORDER_API_SECRET=random-long-secret-string
PORT=8000
```

## Supabase Storage (product images)

1. Supabase → **Storage** → New bucket: `order-images`
2. Set bucket to **Public** (so `<img src>` works without signed URLs)
3. Koyeb environment variables:

| Name | Value |
|------|--------|
| `SUPABASE_URL` | `https://YOUR_PROJECT.supabase.co` |
| `SUPABASE_SERVICE_ROLE_KEY` | Project Settings → API → `service_role` (secret) |
| `SUPABASE_BUCKET` | `order-images` |

4. Redeploy → import Excel again. Images go to Storage; DB only stores the public URL.

Fallback: if env not set, images still save as base64 in DB (old behavior).

Check: `GET /api/storage/status` (requires login).
