-- Chạy trên Supabase SQL Editor (bắt buộc nếu product_image vẫn trống)
ALTER TABLE order_items ALTER COLUMN product_image TYPE TEXT;

-- Kiểm tra
SELECT id, item_code,
       CASE WHEN product_image IS NULL THEN 'NULL'
            WHEN length(product_image) < 50 THEN product_image
            ELSE left(product_image, 40) || '... len=' || length(product_image)::text
       END AS product_image_preview
FROM order_items
ORDER BY id DESC
LIMIT 20;
