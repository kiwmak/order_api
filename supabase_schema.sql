-- Order Management API — schema for Supabase (PostgreSQL)
-- Chạy trong Supabase: SQL Editor → New query → Paste → Run

CREATE TABLE IF NOT EXISTS order_items (
    id SERIAL PRIMARY KEY,
    customer_name VARCHAR(200),
    order_number VARCHAR(100),
    order_date DATE,
    delivery_date DATE,
    customer_order_number VARCHAR(100),
    product_image TEXT,
    item_code VARCHAR(100),
    customer_item_code VARCHAR(100),
    sub_item_code VARCHAR(100),
    main_category VARCHAR(100),
    sub_category VARCHAR(100),
    description TEXT,
    container_size VARCHAR(100),
    container_process VARCHAR(100),
    container_color VARCHAR(200),
    order_qty INTEGER,
    retail_pack_rate INTEGER,
    unit_qty INTEGER,
    unit_wax_weight_g DOUBLE PRECISION,
    fragrance_net_content_ml VARCHAR(50),
    total_wax_weight_kg VARCHAR(50),
    wax_material VARCHAR(100),
    solid_or_bubble_wax VARCHAR(50),
    wick_count INTEGER,
    wax_color VARCHAR(50),
    lid_process VARCHAR(100),
    fragrance_name VARCHAR(200),
    fragrance_code VARCHAR(200),
    fragrance_company VARCHAR(100),
    fragrance_ratio TEXT,
    quality_requirement VARCHAR(50),
    inspection_type VARCHAR(100),
    inspection_requirement VARCHAR(200),
    test_requirement VARCHAR(200),
    sample_requirement VARCHAR(200),
    salesperson VARCHAR(100),
    remarks TEXT,
    packaging_detail TEXT,
    merchandiser VARCHAR(100),
    outer_box_barcode VARCHAR(100),
    outer_box_pack_rate INTEGER,
    inner_box_barcode VARCHAR(100),
    retail_barcode VARCHAR(100),
    imported_at TIMESTAMP WITHOUT TIME ZONE DEFAULT (NOW() AT TIME ZONE 'utc'),
    source_file VARCHAR(300)
);

CREATE INDEX IF NOT EXISTS ix_order_items_id ON order_items (id);
CREATE INDEX IF NOT EXISTS ix_order_items_order_number ON order_items (order_number);
CREATE INDEX IF NOT EXISTS ix_order_items_item_code ON order_items (item_code);

-- Giữ phiếu cân giả định riêng theo mã hàng để lịch sử vẫn tồn tại khi đơn hàng thay đổi.
CREATE TABLE IF NOT EXISTS scale_tickets (
    id SERIAL PRIMARY KEY,
    item_code VARCHAR(100) NOT NULL,
    order_number VARCHAR(100),
    customer_name VARCHAR(200),
    description TEXT,
    spout_count INTEGER NOT NULL,
    cup_min_g INTEGER NOT NULL,
    cup_max_g INTEGER NOT NULL,
    pour1_min_g INTEGER NOT NULL,
    pour1_max_g INTEGER NOT NULL,
    pour2_min_g INTEGER NOT NULL,
    pour2_max_g INTEGER NOT NULL,
    readings_json TEXT NOT NULL,
    average_wax_weight_g DOUBLE PRECISION NOT NULL,
    notes TEXT,
    created_by VARCHAR(100),
    created_at TIMESTAMP WITHOUT TIME ZONE DEFAULT (NOW() AT TIME ZONE 'utc') NOT NULL
);

ALTER TABLE scale_tickets ADD COLUMN IF NOT EXISTS created_by VARCHAR(100);

CREATE INDEX IF NOT EXISTS ix_scale_tickets_item_code ON scale_tickets (item_code);

-- Tài khoản đăng nhập và vai trò phân quyền.
CREATE TABLE IF NOT EXISTS user_accounts (
    id SERIAL PRIMARY KEY,
    username VARCHAR(100) NOT NULL UNIQUE,
    salt VARCHAR(64) NOT NULL,
    password_hash VARCHAR(128) NOT NULL,
    role VARCHAR(20) NOT NULL CHECK (role IN ('admin', 'staff', 'viewer')),
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMP WITHOUT TIME ZONE DEFAULT (NOW() AT TIME ZONE 'utc') NOT NULL
);

CREATE INDEX IF NOT EXISTS ix_user_accounts_id ON user_accounts (id);
CREATE INDEX IF NOT EXISTS ix_user_accounts_username ON user_accounts (username);

-- (Tùy chọn) Cho phép app đọc/ghi qua role mặc định
-- GRANT ALL ON order_items TO postgres;
-- GRANT USAGE, SELECT ON SEQUENCE order_items_id_seq TO postgres;
-- GRANT ALL ON user_accounts TO postgres;
-- GRANT USAGE, SELECT ON SEQUENCE user_accounts_id_seq TO postgres;

-- Nếu bảng đã tồn tại với product_image VARCHAR(500), chạy thêm:
-- ALTER TABLE order_items ALTER COLUMN product_image TYPE TEXT;
