"""Excel order importer – maps Chinese headers to DB fields + extracts embedded images."""
from datetime import datetime, date, timedelta
from typing import List, Dict, Any, Optional, Tuple
import os
import uuid
import openpyxl
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
from models import OrderItem
from translations import HEADER_MAP

# Where product images are saved (relative URL path used in DB)
UPLOAD_SUBDIR = "uploads"


def _excel_date(val) -> Optional[date]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=int(val))).date()
        except Exception:
            return None
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _clean(val):
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip().replace("\r\n", "\n").replace("\r", "\n")
        return s if s and s != "/" else None
    return val


def _to_int(val) -> Optional[int]:
    if val is None or val == "" or val == "/":
        return None
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None


def _to_float(val) -> Optional[float]:
    if val is None or val == "" or val == "/":
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _extract_images(ws, upload_dir: str = None) -> Dict[int, str]:
    """
    Extract embedded images from worksheet.
    Returns dict: excel_row_1based -> data:image/...;base64,... (or static path fallback)

    Primary storage is base64 data-URI so images persist in Postgres/Supabase
    without depending on server disk.
    """
    import base64
    import io

    if upload_dir:
        os.makedirs(upload_dir, exist_ok=True)
    row_to_url: Dict[int, str] = {}

    images = list(getattr(ws, "_images", []) or [])
    print(f"[importer] ws._images count={len(images)}")

    orphan_uris = []  # images without valid row anchor

    for idx, img in enumerate(images):
        try:
            # --- resolve Excel row (1-based) ---
            row_0 = None
            anchor = getattr(img, "anchor", None)
            if anchor is not None:
                if hasattr(anchor, "_from") and anchor._from is not None:
                    row_0 = getattr(anchor._from, "row", None)
                elif hasattr(anchor, "from_"):  # rare
                    row_0 = getattr(anchor.from_, "row", None)

            # --- binary data ---
            data = None
            if hasattr(img, "_data") and callable(img._data):
                try:
                    data = img._data()
                except Exception as e:
                    print(f"[importer] img._data() failed idx={idx}: {e}")
            if not data and hasattr(img, "ref"):
                try:
                    # Some openpyxl versions expose path into archive
                    data = None
                except Exception:
                    pass
            if not data and hasattr(img, "path"):
                try:
                    # Image on disk (unlikely for embedded)
                    with open(img.path, "rb") as f:
                        data = f.read()
                except Exception:
                    pass

            if not data:
                print(f"[importer] no data for image idx={idx}")
                continue

            fmt = (getattr(img, "format", None) or "jpeg").lower().replace(".", "")
            mime_map = {
                "jpeg": "image/jpeg", "jpg": "image/jpeg",
                "png": "image/png", "gif": "image/gif", "bmp": "image/bmp",
            }
            ext_map = {
                "jpeg": ".jpg", "jpg": ".jpg",
                "png": ".png", "gif": ".gif", "bmp": ".bmp",
            }
            mime = mime_map.get(fmt, "image/jpeg")
            ext = ext_map.get(fmt, ".jpg")

            b64 = base64.b64encode(bytes(data)).decode("ascii")
            data_uri = f"data:{mime};base64,{b64}"

            if upload_dir:
                try:
                    filename = f"{uuid.uuid4().hex}{ext}"
                    with open(os.path.join(upload_dir, filename), "wb") as f:
                        f.write(data)
                except Exception as fe:
                    print(f"[importer] file save skip {idx}: {fe}")

            if row_0 is not None:
                excel_row = int(row_0) + 1
                row_to_url[excel_row] = data_uri
                print(f"[importer] image idx={idx} -> row {excel_row}, bytes={len(data)}")
            else:
                orphan_uris.append(data_uri)
                print(f"[importer] image idx={idx} no anchor, bytes={len(data)}")
        except Exception as e:
            print(f"[importer] skip image {idx}: {e}")
            continue

    # Assign orphan images to first data rows that lack images (row 2, 3, 4...)
    if orphan_uris:
        next_row = 2
        for uri in orphan_uris:
            while next_row in row_to_url:
                next_row += 1
            row_to_url[next_row] = uri
            print(f"[importer] orphan image assigned to row {next_row}")
            next_row += 1

    return row_to_url



def parse_excel(
    file_path: str,
    source_name: str = None,
    static_dir: str = None,
) -> List[Dict[str, Any]]:
    """
    Parse Excel order file.
    If static_dir is provided, embedded images are extracted to static_dir/uploads
    and product_image is set to the public URL.
    """
    # data_only=False keeps drawings/images; cell values still readable
    wb = openpyxl.load_workbook(file_path, data_only=False)
    ws = wb.active

    # ALWAYS extract embedded images → base64 data-URI (works on Koyeb/Supabase)
    upload_dir = os.path.join(static_dir, UPLOAD_SUBDIR) if static_dir else None
    row_images: Dict[int, str] = _extract_images(ws, upload_dir)
    if not row_images:
        # Second pass: some files only expose images after full load
        try:
            wb3 = openpyxl.load_workbook(file_path, data_only=False, keep_vba=False)
            row_images = _extract_images(wb3.active, upload_dir)
            wb3.close()
        except Exception as e:
            print(f"[importer] image reload failed: {e}")
    print(f"[importer] extracted images for rows: {list(row_images.keys())}")

    # Header row
    headers = []
    for cell in ws[1]:
        h = cell.value
        if h is None:
            headers.append(None)
            continue
        h = str(h).strip().replace("\r\n", "\n").replace("\r", "\n")
        headers.append(h)

    field_names = []
    for h in headers:
        if h is None:
            field_names.append(None)
        else:
            field_names.append(
                HEADER_MAP.get(h) or HEADER_MAP.get(h.replace("\n", ""))
            )

    rows = []
    for excel_row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(c is None or (isinstance(c, str) and not c.strip()) for c in row):
            continue

        data: Dict[str, Any] = {"source_file": source_name}

        # Attach image from same Excel row if any
        if excel_row_idx in row_images:
            data["product_image"] = row_images[excel_row_idx]

        for i, val in enumerate(row):
            if i >= len(field_names) or field_names[i] is None:
                continue
            field = field_names[i]
            # Don't overwrite image extracted from drawing with empty cell value
            if field == "product_image" and data.get("product_image"):
                continue

            cleaned = _clean(val)

            if field in ("order_date", "delivery_date"):
                data[field] = _excel_date(val)
            elif field in (
                "order_qty",
                "retail_pack_rate",
                "unit_qty",
                "wick_count",
                "outer_box_pack_rate",
            ):
                data[field] = _to_int(val)
            elif field == "unit_wax_weight_g":
                data[field] = _to_float(val)
            else:
                data[field] = cleaned

        # 货号 (col G) vs 子货号 (col I) — both are product codes.
        # Rule: if column I (子货号 / sub_item_code) has data → use I as item_code (mã hàng).
        #        otherwise → keep column G (货号).
        i_code = data.get("sub_item_code")
        if i_code is not None and str(i_code).strip() and str(i_code).strip() != "/":
            data["item_code"] = str(i_code).strip()

        if data.get("order_number") or data.get("item_code"):
            rows.append(data)

    return rows


def rows_to_models(rows: List[Dict[str, Any]]) -> List[OrderItem]:
    items = []
    for r in rows:
        item = OrderItem(
            customer_name=r.get("customer_name"),
            order_number=r.get("order_number"),
            order_date=r.get("order_date"),
            delivery_date=r.get("delivery_date"),
            customer_order_number=r.get("customer_order_number"),
            product_image=r.get("product_image"),
            item_code=str(r.get("item_code")) if r.get("item_code") is not None else None,
            customer_item_code=r.get("customer_item_code"),
            sub_item_code=r.get("sub_item_code"),
            main_category=r.get("main_category"),
            sub_category=r.get("sub_category"),
            description=r.get("description"),
            container_size=r.get("container_size"),
            container_process=r.get("container_process"),
            container_color=r.get("container_color"),
            order_qty=r.get("order_qty"),
            retail_pack_rate=r.get("retail_pack_rate"),
            unit_qty=r.get("unit_qty"),
            unit_wax_weight_g=r.get("unit_wax_weight_g"),
            fragrance_net_content_ml=r.get("fragrance_net_content_ml"),
            total_wax_weight_kg=r.get("total_wax_weight_kg"),
            wax_material=r.get("wax_material"),
            solid_or_bubble_wax=r.get("solid_or_bubble_wax"),
            wick_count=r.get("wick_count"),
            wax_color=r.get("wax_color"),
            lid_process=r.get("lid_process"),
            fragrance_name=r.get("fragrance_name"),
            fragrance_code=r.get("fragrance_code"),
            fragrance_company=r.get("fragrance_company"),
            fragrance_ratio=r.get("fragrance_ratio"),
            quality_requirement=r.get("quality_requirement"),
            inspection_type=r.get("inspection_type"),
            inspection_requirement=r.get("inspection_requirement"),
            test_requirement=r.get("test_requirement"),
            sample_requirement=r.get("sample_requirement"),
            salesperson=r.get("salesperson"),
            remarks=r.get("remarks"),
            packaging_detail=r.get("packaging_detail"),
            merchandiser=r.get("merchandiser"),
            outer_box_barcode=r.get("outer_box_barcode"),
            outer_box_pack_rate=r.get("outer_box_pack_rate"),
            inner_box_barcode=r.get("inner_box_barcode"),
            retail_barcode=r.get("retail_barcode"),
            source_file=r.get("source_file"),
        )
        items.append(item)
    return items
