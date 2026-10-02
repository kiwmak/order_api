"""Excel order importer – maps Chinese headers to DB fields + extracts embedded images."""
from datetime import datetime, date, timedelta
from typing import List, Dict, Any, Optional, Tuple
import os
import uuid
import zipfile
import base64
import re
from xml.etree import ElementTree as ET
import openpyxl
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
from models import OrderItem
from translations import HEADER_MAP
from storage import image_ref_from_bytes, storage_configured

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



def _mime_for_name(name: str) -> str:
    n = name.lower()
    if n.endswith('.png'):
        return 'image/png'
    if n.endswith('.gif'):
        return 'image/gif'
    if n.endswith('.bmp'):
        return 'image/bmp'
    if n.endswith('.webp'):
        return 'image/webp'
    return 'image/jpeg'


def _extract_images_from_zip(file_path: str) -> Dict[int, str]:
    """
    Reliable image extraction: read xlsx as ZIP.
    Parse drawing XML for row anchors + media files → data-URI map {excel_row: data_uri}.
    """
    row_to_url: Dict[int, str] = {}
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            names = z.namelist()
            media = sorted([n for n in names if n.startswith('xl/media/') and not n.endswith('/')])
            if not media:
                print('[importer] zip: no xl/media/* found')
                return {}

            # Load media BYTES first — upload to Storage only when linked to a data row
            print(f"[importer] supabase storage configured={storage_configured()}")
            media_bytes: Dict[str, tuple] = {}  # key -> (raw, mime, fname)
            for mpath in media:
                raw = z.read(mpath)
                fname = mpath.split('/')[-1]
                mime = _mime_for_name(fname)
                media_bytes[fname] = (raw, mime, fname)
                media_bytes[mpath] = (raw, mime, fname)
                media_bytes['../media/' + fname] = (raw, mime, fname)
                media_bytes['media/' + fname] = (raw, mime, fname)
                print(f'[importer] zip media loaded {fname} bytes={len(raw)}')

            def _uri_for(target: str):
                """Resolve target path/name to Storage URL or base64 (upload once per file)."""
                if not target:
                    return None
                info = media_bytes.get(target) or media_bytes.get(target.split('/')[-1])
                if not info:
                    return None
                raw, mime, fname = info
                # cache on fname so same image not re-uploaded
                cache_key = f"__uri__{fname}"
                if cache_key in media_bytes:
                    return media_bytes[cache_key]  # type: ignore
                uri = image_ref_from_bytes(raw, mime=mime, filename_hint=fname)
                media_bytes[cache_key] = uri  # type: ignore
                kind = "storage" if str(uri).startswith("http") else "base64"
                print(f'[importer] image ready {fname} -> {kind} len={len(str(uri))}')
                return uri

            # Parse relationships: rId -> media path
            rid_to_file: Dict[str, str] = {}
            for rel_name in [n for n in names if n.startswith('xl/drawings/_rels/') and n.endswith('.rels')]:
                try:
                    root = ET.fromstring(z.read(rel_name))
                    for rel in root:
                        rid = rel.attrib.get('Id') or rel.attrib.get('id')
                        target = rel.attrib.get('Target') or rel.attrib.get('target')
                        if rid and target:
                            # Target like ../media/image1.jpeg
                            rid_to_file[rid] = target
                except Exception as e:
                    print(f'[importer] rels parse {rel_name}: {e}')

            # Parse drawing XML: twoCellAnchor / oneCellAnchor → row + r:embed
            NS = {
                'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            }
            for draw_name in [n for n in names if n.startswith('xl/drawings/drawing') and n.endswith('.xml')]:
                try:
                    root = ET.fromstring(z.read(draw_name))
                    anchors = list(root.findall('xdr:twoCellAnchor', NS)) + list(root.findall('xdr:oneCellAnchor', NS))
                    for anc in anchors:
                        from_el = anc.find('xdr:from', NS)
                        if from_el is None:
                            continue
                        row_el = from_el.find('xdr:row', NS)
                        if row_el is None or row_el.text is None:
                            continue
                        row_0 = int(row_el.text)
                        excel_row = row_0 + 1  # 0-based in XML → 1-based Excel

                        blip = anc.find('.//a:blip', NS)
                        if blip is None:
                            continue
                        rid = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if not rid:
                            rid = blip.attrib.get('embed')
                        if not rid:
                            continue
                        target = rid_to_file.get(rid)
                        if not target:
                            continue
                        # resolve + upload only for this row
                        uri = _uri_for(target)
                        if uri:
                            row_to_url[excel_row] = uri
                            print(f'[importer] zip map rId={rid} -> excel row {excel_row}')
                except Exception as e:
                    print(f'[importer] drawing parse {draw_name}: {e}')

            # Fallback: sequential assign if no row mapping
            if not row_to_url and media:
                for i, mpath in enumerate(media):
                    fname = mpath.split('/')[-1]
                    uri = _uri_for(fname)
                    if uri:
                        row_to_url[2 + i] = uri  # data starts at row 2
                        print(f'[importer] zip sequential row {2+i} <- {fname}')
    except Exception as e:
        print(f'[importer] zip extract failed: {e}')
    return row_to_url


def _extract_images(ws, upload_dir: str = None) -> Dict[int, str]:
    """Fallback: openpyxl ws._images (may be empty on some builds)."""
    row_to_url: Dict[int, str] = {}
    images = list(getattr(ws, '_images', []) or [])
    print(f'[importer] openpyxl ws._images count={len(images)}')
    for idx, img in enumerate(images):
        try:
            row_0 = None
            anchor = getattr(img, 'anchor', None)
            if anchor is not None and hasattr(anchor, '_from') and anchor._from is not None:
                row_0 = getattr(anchor._from, 'row', None)
            data = None
            if hasattr(img, '_data') and callable(img._data):
                data = img._data()
            if not data:
                continue
            fmt = (getattr(img, 'format', None) or 'jpeg').lower().replace('.', '')
            mime = 'image/png' if fmt == 'png' else 'image/jpeg'
            uri = image_ref_from_bytes(bytes(data), mime=mime)
            if row_0 is not None:
                row_to_url[int(row_0) + 1] = uri
            if upload_dir:
                try:
                    os.makedirs(upload_dir, exist_ok=True)
                    ext = '.png' if fmt == 'png' else '.jpg'
                    with open(os.path.join(upload_dir, f'{uuid.uuid4().hex}{ext}'), 'wb') as f:
                        f.write(data)
                except Exception:
                    pass
        except Exception as e:
            print(f'[importer] openpyxl image skip {idx}: {e}')
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

    # PRIMARY: zip-based extraction (reliable on all platforms / Koyeb)
    row_images: Dict[int, str] = _extract_images_from_zip(file_path)
    # Fallback: openpyxl drawings
    if not row_images:
        upload_dir = os.path.join(static_dir, UPLOAD_SUBDIR) if static_dir else None
        row_images = _extract_images(ws, upload_dir)
    print(f"[importer] FINAL image rows: {list(row_images.keys())} count={len(row_images)}")

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

    # Backfill: rows without image get remaining extracted images (in order)
    assigned = {r.get("product_image") for r in rows if r.get("product_image")}
    leftover = [uri for _, uri in sorted(row_images.items()) if uri not in assigned]
    if leftover:
        li = 0
        for r in rows:
            if not r.get("product_image") and li < len(leftover):
                r["product_image"] = leftover[li]
                li += 1
                print(f"[importer] backfill image onto item {r.get('item_code')}")

    with_img = sum(1 for r in rows if r.get("product_image"))
    storage_n = sum(1 for r in rows if str(r.get("product_image") or "").startswith("http"))
    print(f"[importer] rows={len(rows)} with_images={with_img} storage_urls={storage_n}")
    for r in rows:
        pi = r.get("product_image") or ""
        if pi:
            print(f"[importer]   item={r.get('item_code')} image={pi[:80]}...")
    return rows


def _safe_int(val):
    if val is None or val == "" or val == "/":
        return None
    try:
        return int(float(str(val).strip().replace(",", "")))
    except (ValueError, TypeError):
        return None


def _safe_float(val):
    if val is None or val == "" or val == "/":
        return None
    try:
        return float(str(val).strip().replace(",", ""))
    except (ValueError, TypeError):
        return None


def _safe_str(val):
    if val is None:
        return None
    s = str(val).strip()
    return s if s and s != "/" else None



def rows_to_models(rows: List[Dict[str, Any]]) -> List[OrderItem]:
    items = []
    for r in rows:
        item = OrderItem(
            customer_name=_safe_str(r.get("customer_name")),
            order_number=_safe_str(r.get("order_number")),
            order_date=r.get("order_date"),
            delivery_date=r.get("delivery_date"),
            customer_order_number=_safe_str(r.get("customer_order_number")),
            product_image=r.get("product_image"),
            item_code=str(r.get("item_code")) if r.get("item_code") is not None else None,
            customer_item_code=_safe_str(r.get("customer_item_code")),
            sub_item_code=_safe_str(r.get("sub_item_code")),
            main_category=_safe_str(r.get("main_category")),
            sub_category=_safe_str(r.get("sub_category")),
            description=_safe_str(r.get("description")),
            container_size=_safe_str(r.get("container_size")),
            container_process=_safe_str(r.get("container_process")),
            container_color=_safe_str(r.get("container_color")),
            order_qty=_safe_int(r.get("order_qty")),
            retail_pack_rate=_safe_int(r.get("retail_pack_rate")),
            unit_qty=_safe_int(r.get("unit_qty")),
            unit_wax_weight_g=_safe_float(r.get("unit_wax_weight_g")),
            fragrance_net_content_ml=_safe_str(r.get("fragrance_net_content_ml")),
            total_wax_weight_kg=_safe_str(r.get("total_wax_weight_kg")),
            wax_material=_safe_str(r.get("wax_material")),
            solid_or_bubble_wax=_safe_str(r.get("solid_or_bubble_wax")),
            wick_count=_safe_int(r.get("wick_count")),
            wax_color=_safe_str(r.get("wax_color")),
            lid_process=_safe_str(r.get("lid_process")),
            fragrance_name=_safe_str(r.get("fragrance_name")),
            fragrance_code=_safe_str(r.get("fragrance_code")),
            fragrance_company=_safe_str(r.get("fragrance_company")),
            fragrance_ratio=_safe_str(r.get("fragrance_ratio")),
            quality_requirement=_safe_str(r.get("quality_requirement")),
            inspection_type=_safe_str(r.get("inspection_type")),
            inspection_requirement=_safe_str(r.get("inspection_requirement")),
            test_requirement=_safe_str(r.get("test_requirement")),
            sample_requirement=_safe_str(r.get("sample_requirement")),
            salesperson=_safe_str(r.get("salesperson")),
            remarks=_safe_str(r.get("remarks")),
            packaging_detail=_safe_str(r.get("packaging_detail")),
            merchandiser=_safe_str(r.get("merchandiser")),
            outer_box_barcode=_safe_str(r.get("outer_box_barcode")),
            outer_box_pack_rate=_safe_int(r.get("outer_box_pack_rate")),
            inner_box_barcode=_safe_str(r.get("inner_box_barcode")),
            retail_barcode=_safe_str(r.get("retail_barcode")),
            source_file=_safe_str(r.get("source_file")),
        )
        items.append(item)
    return items
