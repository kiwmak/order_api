"""
Order Management Web API + UI
- Upload Excel order detail files (Chinese headers supported)
- Store in SQLite
- Bilingual UI (中文 / English / Tiếng Việt)
- REST API for programmatic access
"""
import os
import json
import io
import shutil
import tempfile
from typing import List, Optional, Any
from datetime import date, datetime
from collections import defaultdict

from fastapi import FastAPI, UploadFile, File, Depends, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from sqlalchemy import func, desc
from pydantic import BaseModel

from models import init_db, get_db, OrderItem, SessionLocal
from importer import parse_excel, rows_to_models
from translations import FIELD_LABELS, UI_TEXT
from auth import (
    verify_password,
    create_token,
    get_current_user,
    get_optional_user,
    change_password,
)

app = FastAPI(
    title="Order Management API",
    description="Import & manage candle order details from Excel. Bilingual support.",
    version="1.0.0",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

# Create static folder if missing (avoids RuntimeError on Windows when folder not copied)
os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# ---------- Pydantic schemas ----------
class OrderItemOut(BaseModel):
    id: int
    customer_name: Optional[str] = None
    order_number: Optional[str] = None
    order_date: Optional[date] = None
    delivery_date: Optional[date] = None
    customer_order_number: Optional[str] = None
    product_image: Optional[str] = None  # URL e.g. /static/uploads/xxx.jpg
    item_code: Optional[str] = None
    customer_item_code: Optional[str] = None
    sub_item_code: Optional[str] = None
    main_category: Optional[str] = None
    sub_category: Optional[str] = None
    description: Optional[str] = None
    container_size: Optional[str] = None
    container_process: Optional[str] = None
    container_color: Optional[str] = None
    order_qty: Optional[int] = None
    retail_pack_rate: Optional[int] = None
    unit_qty: Optional[int] = None
    unit_wax_weight_g: Optional[float] = None
    fragrance_net_content_ml: Optional[str] = None
    total_wax_weight_kg: Optional[str] = None
    wax_material: Optional[str] = None
    solid_or_bubble_wax: Optional[str] = None
    wick_count: Optional[int] = None
    wax_color: Optional[str] = None
    lid_process: Optional[str] = None
    fragrance_name: Optional[str] = None
    fragrance_code: Optional[str] = None
    fragrance_company: Optional[str] = None
    fragrance_ratio: Optional[str] = None
    quality_requirement: Optional[str] = None
    inspection_type: Optional[str] = None
    inspection_requirement: Optional[str] = None
    test_requirement: Optional[str] = None
    sample_requirement: Optional[str] = None
    salesperson: Optional[str] = None
    remarks: Optional[str] = None
    packaging_detail: Optional[str] = None
    merchandiser: Optional[str] = None
    outer_box_barcode: Optional[str] = None
    outer_box_pack_rate: Optional[int] = None
    inner_box_barcode: Optional[str] = None
    retail_barcode: Optional[str] = None
    source_file: Optional[str] = None

    class Config:
        from_attributes = True


class FileImportDetail(BaseModel):
    filename: str
    imported_count: int = 0  # new rows inserted
    updated_count: int = 0   # existing rows updated (same order_number + item_code)
    image_count: int = 0
    error: Optional[str] = None


class ImportResult(BaseModel):
    success: bool
    imported_count: int = 0  # total new inserts
    updated_count: int = 0   # total updates
    file_count: int = 1
    message: str
    details: List[FileImportDetail] = []


class LoginRequest(BaseModel):
    username: str
    password: str


class LoginResponse(BaseModel):
    success: bool
    token: str
    username: str
    message: str


class OrderItemUpdate(BaseModel):
    """Fields allowed to edit (all optional)."""
    customer_name: Optional[str] = None
    order_number: Optional[str] = None
    order_date: Optional[date] = None
    delivery_date: Optional[date] = None
    customer_order_number: Optional[str] = None
    item_code: Optional[str] = None
    customer_item_code: Optional[str] = None
    sub_item_code: Optional[str] = None
    main_category: Optional[str] = None
    sub_category: Optional[str] = None
    description: Optional[str] = None
    container_size: Optional[str] = None
    container_process: Optional[str] = None
    container_color: Optional[str] = None
    order_qty: Optional[int] = None
    retail_pack_rate: Optional[int] = None
    unit_qty: Optional[int] = None
    unit_wax_weight_g: Optional[float] = None
    fragrance_net_content_ml: Optional[str] = None
    total_wax_weight_kg: Optional[str] = None
    wax_material: Optional[str] = None
    solid_or_bubble_wax: Optional[str] = None
    wick_count: Optional[int] = None
    wax_color: Optional[str] = None
    lid_process: Optional[str] = None
    fragrance_name: Optional[str] = None
    fragrance_code: Optional[str] = None
    fragrance_company: Optional[str] = None
    fragrance_ratio: Optional[str] = None
    quality_requirement: Optional[str] = None
    inspection_type: Optional[str] = None
    inspection_requirement: Optional[str] = None
    test_requirement: Optional[str] = None
    sample_requirement: Optional[str] = None
    salesperson: Optional[str] = None
    remarks: Optional[str] = None
    packaging_detail: Optional[str] = None
    merchandiser: Optional[str] = None
    outer_box_barcode: Optional[str] = None
    outer_box_pack_rate: Optional[int] = None
    inner_box_barcode: Optional[str] = None
    retail_barcode: Optional[str] = None


# ---------- Startup ----------
@app.on_event("startup")
def on_startup():
    init_db()


# ---------- Web UI ----------
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html", {
        "request": request,
        "field_labels": FIELD_LABELS,
        "ui": UI_TEXT,
    })


# ---------- Auth ----------
@app.post("/api/login", response_model=LoginResponse)
def login(body: LoginRequest):
    if not verify_password(body.username.strip(), body.password):
        raise HTTPException(401, "Invalid username or password")
    token = create_token(body.username.strip())
    return LoginResponse(
        success=True,
        token=token,
        username=body.username.strip(),
        message="Login OK",
    )


@app.get("/api/me")
def me(user: Optional[str] = Depends(get_optional_user)):
    if not user:
        return {"authenticated": False, "username": None}
    return {"authenticated": True, "username": user}


class ChangePasswordRequest(BaseModel):
    old_password: str
    new_password: str


@app.post("/api/change-password")
def api_change_password(
    body: ChangePasswordRequest,
    user: str = Depends(get_current_user),
):
    ok, msg = change_password(user, body.old_password, body.new_password)
    if not ok:
        raise HTTPException(400, msg)
    return {"success": True, "message": msg}



# Fields that can be updated on existing rows (exclude id / primary key)
_UPDATE_FIELDS = [
    "customer_name", "order_date", "delivery_date", "customer_order_number",
    "product_image", "customer_item_code", "sub_item_code",
    "main_category", "sub_category", "description",
    "container_size", "container_process", "container_color",
    "order_qty", "retail_pack_rate", "unit_qty", "unit_wax_weight_g",
    "fragrance_net_content_ml", "total_wax_weight_kg", "wax_material",
    "solid_or_bubble_wax", "wick_count", "wax_color", "lid_process",
    "fragrance_name", "fragrance_code", "fragrance_company", "fragrance_ratio",
    "quality_requirement", "inspection_type", "inspection_requirement",
    "test_requirement", "sample_requirement", "salesperson", "remarks",
    "packaging_detail", "merchandiser", "outer_box_barcode", "outer_box_pack_rate",
    "inner_box_barcode", "retail_barcode", "source_file",
]


def _upsert_rows(rows: list, db: Session) -> tuple:
    """
    Insert or update by (order_number, item_code).
    If order_number + item_code already exist → update all fields.
    Otherwise → insert new row.
    Returns (inserted, updated, image_count).
    """
    from datetime import datetime as _dt

    items = rows_to_models(rows)
    inserted = 0
    updated = 0
    image_count = 0

    for item in items:
        if item.product_image:
            image_count += 1

        order_no = (item.order_number or "").strip()
        item_code = (item.item_code or "").strip() if item.item_code else ""

        existing = None
        if order_no and item_code:
            existing = (
                db.query(OrderItem)
                .filter(
                    OrderItem.order_number == order_no,
                    OrderItem.item_code == item_code,
                )
                .first()
            )
        elif order_no:
            # fallback: order_number only when item_code empty
            from sqlalchemy import or_
            existing = (
                db.query(OrderItem)
                .filter(
                    OrderItem.order_number == order_no,
                    or_(OrderItem.item_code.is_(None), OrderItem.item_code == ""),
                )
                .first()
            )

        if existing:
            for field in _UPDATE_FIELDS:
                new_val = getattr(item, field, None)
                # Keep old image if new row has no image
                if field == "product_image" and not new_val:
                    continue
                setattr(existing, field, new_val)
            existing.imported_at = _dt.utcnow()
            updated += 1
        else:
            db.add(item)
            inserted += 1

    return inserted, updated, image_count


def _import_bytes(content: bytes, filename: str, db: Session) -> FileImportDetail:
    """Parse one Excel from bytes, upsert rows. Caller commits."""
    if not filename:
        return FileImportDetail(filename="(empty)", imported_count=0, error="No filename")

    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".xlsx", ".xls"):
        return FileImportDetail(
            filename=filename,
            imported_count=0,
            error="Only .xlsx / .xls supported",
        )
    if not content:
        return FileImportDetail(filename=filename, imported_count=0, error="Empty file")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        rows = parse_excel(
            tmp_path,
            source_name=filename,
            static_dir=STATIC_DIR,
        )
        if not rows:
            return FileImportDetail(
                filename=filename,
                imported_count=0,
                error="No valid order rows found",
            )

        inserted, updated, with_img = _upsert_rows(rows, db)
        return FileImportDetail(
            filename=filename,
            imported_count=inserted,
            updated_count=updated,
            image_count=with_img,
        )
    except Exception as e:
        return FileImportDetail(
            filename=filename or "(unknown)",
            imported_count=0,
            error=str(e),
        )
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


# ---------- API: Import (single or multiple files) — requires login ----------
@app.post("/api/import", response_model=ImportResult)
async def import_excel(
    files: List[UploadFile] = File(..., description="One or more Excel files"),
    db: Session = Depends(get_db),
    user: str = Depends(get_current_user),
):
    # Normalize: filter empty / browser ghost slots
    uploads = [f for f in files if f is not None and (f.filename or "").strip()]
    if not uploads:
        raise HTTPException(400, "No file uploaded")

    details: List[FileImportDetail] = []
    total_new = 0
    total_upd = 0
    total_img = 0
    errors = 0

    # Process & COMMIT each file separately so one failure does not roll back others
    for f in uploads:
        fname = f.filename or "unknown.xlsx"
        try:
            content = await f.read()
        except Exception as e:
            details.append(FileImportDetail(filename=fname, imported_count=0, error=f"Read failed: {e}"))
            errors += 1
            continue

        detail = _import_bytes(content, fname, db)
        if detail.error:
            db.rollback()
            details.append(detail)
            errors += 1
            continue

        try:
            db.commit()
            total_new += detail.imported_count
            total_upd += detail.updated_count
            total_img += detail.image_count
            details.append(detail)
            db.expunge_all()
        except Exception as e:
            db.rollback()
            details.append(FileImportDetail(
                filename=fname,
                imported_count=0,
                updated_count=0,
                image_count=0,
                error=f"DB commit failed: {e}",
            ))
            errors += 1

    if total_new == 0 and total_upd == 0 and errors == len(uploads):
        msgs = "; ".join(f"{d.filename}: {d.error}" for d in details if d.error)
        raise HTTPException(400, f"Import failed for all files. {msgs}")

    ok_files = len(uploads) - errors
    parts = []
    if total_new:
        parts.append(f"{total_new} new")
    if total_upd:
        parts.append(f"{total_upd} updated")
    summary = ", ".join(parts) if parts else "0 rows"
    msg = f"{summary} from {ok_files}/{len(uploads)} file(s)"
    if total_img:
        msg += f" ({total_img} with images)"
    if errors:
        msg += f" — {errors} file(s) failed: "
        msg += "; ".join(f"{d.filename}: {d.error}" for d in details if d.error)

    return ImportResult(
        success=errors == 0,
        imported_count=total_new,
        updated_count=total_upd,
        file_count=len(uploads),
        message=msg,
        details=details,
    )


# ---------- API: List / Search ----------
@app.get("/api/orders", response_model=List[OrderItemOut])
def list_orders(
    q: Optional[str] = Query(None, description="Search order_number / item_code / customer_name"),
    order_number: Optional[str] = Query(None),
    skip: int = 0,
    limit: int = 200,
    db: Session = Depends(get_db),
):
    query = db.query(OrderItem)
    if order_number:
        query = query.filter(OrderItem.order_number == order_number)
    if q:
        like = f"%{q}%"
        query = query.filter(
            (OrderItem.order_number.ilike(like))
            | (OrderItem.item_code.ilike(like))
            | (OrderItem.customer_name.ilike(like))
            | (OrderItem.customer_item_code.ilike(like))
            | (OrderItem.description.ilike(like))
        )
    items = query.order_by(OrderItem.id.desc()).offset(skip).limit(limit).all()
    return items


@app.get("/api/orders/{item_id}", response_model=OrderItemOut)
def get_order(item_id: int, db: Session = Depends(get_db)):
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Order item not found")
    return item


@app.get("/api/orders/by-number/{order_number}", response_model=List[OrderItemOut])
def get_by_order_number(order_number: str, db: Session = Depends(get_db)):
    items = db.query(OrderItem).filter(OrderItem.order_number == order_number).all()
    return items


@app.delete("/api/orders")
def clear_all(
    db: Session = Depends(get_db),
    user: str = Depends(get_current_user),
):
    count = db.query(OrderItem).delete()
    db.commit()
    return {"deleted": count}


@app.delete("/api/orders/{item_id}")
def delete_item(
    item_id: int,
    db: Session = Depends(get_db),
    user: str = Depends(get_current_user),
):
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Not found")
    db.delete(item)
    db.commit()
    return {"deleted": item_id}


@app.put("/api/orders/{item_id}", response_model=OrderItemOut)
def update_item(
    item_id: int,
    body: OrderItemUpdate,
    db: Session = Depends(get_db),
    user: str = Depends(get_current_user),
):
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Order item not found")
    data = body.model_dump(exclude_unset=True) if hasattr(body, "model_dump") else body.dict(exclude_unset=True)
    for k, v in data.items():
        setattr(item, k, v)
    db.commit()
    db.refresh(item)
    return item



# ---------- Backup / Restore ----------
@app.get("/api/backup")
def backup_database(
    db: Session = Depends(get_db),
    user: str = Depends(get_current_user),
):
    """Download full database backup as JSON (includes product images)."""
    items = db.query(OrderItem).order_by(OrderItem.id).all()
    rows = []
    for x in items:
        rows.append({
            "customer_name": x.customer_name,
            "order_number": x.order_number,
            "order_date": x.order_date.isoformat() if x.order_date else None,
            "delivery_date": x.delivery_date.isoformat() if x.delivery_date else None,
            "customer_order_number": x.customer_order_number,
            "product_image": x.product_image,
            "item_code": x.item_code,
            "customer_item_code": x.customer_item_code,
            "sub_item_code": x.sub_item_code,
            "main_category": x.main_category,
            "sub_category": x.sub_category,
            "description": x.description,
            "container_size": x.container_size,
            "container_process": x.container_process,
            "container_color": x.container_color,
            "order_qty": x.order_qty,
            "retail_pack_rate": x.retail_pack_rate,
            "unit_qty": x.unit_qty,
            "unit_wax_weight_g": x.unit_wax_weight_g,
            "fragrance_net_content_ml": x.fragrance_net_content_ml,
            "total_wax_weight_kg": x.total_wax_weight_kg,
            "wax_material": x.wax_material,
            "solid_or_bubble_wax": x.solid_or_bubble_wax,
            "wick_count": x.wick_count,
            "wax_color": x.wax_color,
            "lid_process": x.lid_process,
            "fragrance_name": x.fragrance_name,
            "fragrance_code": x.fragrance_code,
            "fragrance_company": x.fragrance_company,
            "fragrance_ratio": x.fragrance_ratio,
            "quality_requirement": x.quality_requirement,
            "inspection_type": x.inspection_type,
            "inspection_requirement": x.inspection_requirement,
            "test_requirement": x.test_requirement,
            "sample_requirement": x.sample_requirement,
            "salesperson": x.salesperson,
            "remarks": x.remarks,
            "packaging_detail": x.packaging_detail,
            "merchandiser": x.merchandiser,
            "outer_box_barcode": x.outer_box_barcode,
            "outer_box_pack_rate": x.outer_box_pack_rate,
            "inner_box_barcode": x.inner_box_barcode,
            "retail_barcode": x.retail_barcode,
            "imported_at": x.imported_at.isoformat() if x.imported_at else None,
            "source_file": x.source_file,
        })
    payload = {
        "format": "order_api_backup",
        "version": 1,
        "exported_at": datetime.utcnow().isoformat() + "Z",
        "exported_by": user,
        "count": len(rows),
        "items": rows,
    }
    body = json.dumps(payload, ensure_ascii=False, indent=2)
    filename = f"order_backup_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.json"
    return Response(
        content=body.encode("utf-8"),
        media_type="application/json; charset=utf-8",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        },
    )


@app.post("/api/restore")
async def restore_database(
    file: UploadFile = File(..., description="Backup JSON file"),
    mode: str = Query("replace", description="replace = clear then import; merge = upsert by order_number+item_code"),
    db: Session = Depends(get_db),
    user: str = Depends(get_current_user),
):
    """Restore database from a JSON backup file."""
    if not file.filename or not file.filename.lower().endswith(".json"):
        raise HTTPException(400, "Please upload a .json backup file")
    try:
        raw = await file.read()
        data = json.loads(raw.decode("utf-8"))
    except Exception as e:
        raise HTTPException(400, f"Invalid JSON: {e}")

    items = data.get("items") if isinstance(data, dict) else data
    if not isinstance(items, list):
        raise HTTPException(400, "Backup must contain an items array")

    mode = (mode or "replace").lower().strip()
    if mode not in ("replace", "merge"):
        raise HTTPException(400, "mode must be replace or merge")

    deleted = 0
    if mode == "replace":
        deleted = db.query(OrderItem).delete()
        db.commit()

    inserted = 0
    updated = 0
    errors = 0

    for row in items:
        if not isinstance(row, dict):
            errors += 1
            continue
        try:
            # parse dates
            def _d(v):
                if not v:
                    return None
                if isinstance(v, date) and not isinstance(v, datetime):
                    return v
                s = str(v)[:10]
                try:
                    return datetime.strptime(s, "%Y-%m-%d").date()
                except Exception:
                    return None

            def _dt(v):
                if not v:
                    return None
                try:
                    return datetime.fromisoformat(str(v).replace("Z", ""))
                except Exception:
                    return None

            order_no = (row.get("order_number") or "")
            if isinstance(order_no, str):
                order_no = order_no.strip()
            else:
                order_no = str(order_no).strip() if order_no is not None else ""
            item_code = row.get("item_code")
            if item_code is not None:
                item_code = str(item_code).strip()

            existing = None
            if mode == "merge" and order_no:
                q = db.query(OrderItem).filter(OrderItem.order_number == order_no)
                if item_code:
                    q = q.filter(OrderItem.item_code == item_code)
                existing = q.first()

            fields = {
                "customer_name": row.get("customer_name"),
                "order_number": order_no or None,
                "order_date": _d(row.get("order_date")),
                "delivery_date": _d(row.get("delivery_date")),
                "customer_order_number": row.get("customer_order_number"),
                "product_image": row.get("product_image"),
                "item_code": item_code,
                "customer_item_code": row.get("customer_item_code"),
                "sub_item_code": row.get("sub_item_code"),
                "main_category": row.get("main_category"),
                "sub_category": row.get("sub_category"),
                "description": row.get("description"),
                "container_size": row.get("container_size"),
                "container_process": row.get("container_process"),
                "container_color": row.get("container_color"),
                "order_qty": row.get("order_qty"),
                "retail_pack_rate": row.get("retail_pack_rate"),
                "unit_qty": row.get("unit_qty"),
                "unit_wax_weight_g": row.get("unit_wax_weight_g"),
                "fragrance_net_content_ml": row.get("fragrance_net_content_ml"),
                "total_wax_weight_kg": row.get("total_wax_weight_kg"),
                "wax_material": row.get("wax_material"),
                "solid_or_bubble_wax": row.get("solid_or_bubble_wax"),
                "wick_count": row.get("wick_count"),
                "wax_color": row.get("wax_color"),
                "lid_process": row.get("lid_process"),
                "fragrance_name": row.get("fragrance_name"),
                "fragrance_code": row.get("fragrance_code"),
                "fragrance_company": row.get("fragrance_company"),
                "fragrance_ratio": row.get("fragrance_ratio"),
                "quality_requirement": row.get("quality_requirement"),
                "inspection_type": row.get("inspection_type"),
                "inspection_requirement": row.get("inspection_requirement"),
                "test_requirement": row.get("test_requirement"),
                "sample_requirement": row.get("sample_requirement"),
                "salesperson": row.get("salesperson"),
                "remarks": row.get("remarks"),
                "packaging_detail": row.get("packaging_detail"),
                "merchandiser": row.get("merchandiser"),
                "outer_box_barcode": row.get("outer_box_barcode"),
                "outer_box_pack_rate": row.get("outer_box_pack_rate"),
                "inner_box_barcode": row.get("inner_box_barcode"),
                "retail_barcode": row.get("retail_barcode"),
                "source_file": row.get("source_file") or f"restore:{file.filename}",
            }
            if existing:
                for k, v in fields.items():
                    if k == "product_image" and not v:
                        continue
                    setattr(existing, k, v)
                existing.imported_at = _dt(row.get("imported_at")) or datetime.utcnow()
                updated += 1
            else:
                obj = OrderItem(**fields)
                obj.imported_at = _dt(row.get("imported_at")) or datetime.utcnow()
                db.add(obj)
                inserted += 1
        except Exception:
            errors += 1
            continue

    try:
        db.commit()
    except Exception as e:
        db.rollback()
        raise HTTPException(500, f"Restore commit failed: {e}")

    return {
        "success": True,
        "mode": mode,
        "deleted": deleted,
        "inserted": inserted,
        "updated": updated,
        "errors": errors,
        "message": f"Restore OK ({mode}): {inserted} new, {updated} updated" + (f", cleared {deleted}" if deleted else ""),
    }


@app.get("/api/labels")
def get_labels(lang: str = Query("en", pattern="^(zh|en|vi)$")):
    """Return field labels in the requested language."""
    return {k: v.get(lang, v.get("en")) for k, v in FIELD_LABELS.items()}


@app.get("/api/stats")
def stats(db: Session = Depends(get_db)):
    total = db.query(OrderItem).count()
    orders = db.query(OrderItem.order_number).distinct().count()
    customers = db.query(OrderItem.customer_name).distinct().count()
    max_id = db.query(func.max(OrderItem.id)).scalar() or 0
    last_import = db.query(func.max(OrderItem.imported_at)).scalar()
    qty_sum = db.query(func.coalesce(func.sum(OrderItem.order_qty), 0)).scalar() or 0
    return {
        "total_items": total,
        "distinct_orders": orders,
        "distinct_customers": customers,
        "total_qty": int(qty_sum),
        "max_id": max_id,
        "last_imported_at": last_import.isoformat() if last_import else None,
    }


@app.get("/api/dashboard")
def dashboard(
    top: int = Query(10, ge=1, le=50),
    db: Session = Depends(get_db),
):
    """Full dashboard data: KPIs, charts by customer / item / order, recent rows."""
    total_items = db.query(OrderItem).count()
    distinct_orders = db.query(OrderItem.order_number).distinct().count()
    distinct_customers = db.query(OrderItem.customer_name).distinct().count()
    distinct_items = db.query(OrderItem.item_code).distinct().count()
    total_qty = int(db.query(func.coalesce(func.sum(OrderItem.order_qty), 0)).scalar() or 0)
    max_id = db.query(func.max(OrderItem.id)).scalar() or 0
    last_import = db.query(func.max(OrderItem.imported_at)).scalar()

    # --- By customer ---
    cust_rows = (
        db.query(
            OrderItem.customer_name,
            func.count(OrderItem.id).label("line_count"),
            func.coalesce(func.sum(OrderItem.order_qty), 0).label("qty"),
            func.count(func.distinct(OrderItem.order_number)).label("order_count"),
        )
        .group_by(OrderItem.customer_name)
        .order_by(desc("qty"))
        .limit(top)
        .all()
    )
    by_customer = [
        {
            "name": r.customer_name or "(N/A)",
            "line_count": r.line_count,
            "qty": int(r.qty),
            "order_count": r.order_count,
        }
        for r in cust_rows
    ]

    # --- By item code ---
    item_rows = (
        db.query(
            OrderItem.item_code,
            func.count(OrderItem.id).label("line_count"),
            func.coalesce(func.sum(OrderItem.order_qty), 0).label("qty"),
            func.count(func.distinct(OrderItem.customer_name)).label("customer_count"),
            func.count(func.distinct(OrderItem.order_number)).label("order_count"),
        )
        .group_by(OrderItem.item_code)
        .order_by(desc("qty"))
        .limit(top)
        .all()
    )
    by_item = []
    for r in item_rows:
        code = r.item_code
        # Distinct customers & order numbers for tooltip
        cust_list = [
            x[0] for x in db.query(OrderItem.customer_name)
            .filter(OrderItem.item_code == code)
            .distinct()
            .limit(15)
            .all()
            if x[0]
        ]
        order_list = [
            x[0] for x in db.query(OrderItem.order_number)
            .filter(OrderItem.item_code == code)
            .distinct()
            .limit(15)
            .all()
            if x[0]
        ]
        by_item.append({
            "code": code or "(N/A)",
            "line_count": r.line_count,
            "qty": int(r.qty),
            "customer_count": r.customer_count,
            "order_count": r.order_count,
            "customers": cust_list,
            "orders": order_list,
        })

    # --- By order number ---
    order_rows = (
        db.query(
            OrderItem.order_number,
            OrderItem.customer_name,
            func.count(OrderItem.id).label("line_count"),
            func.coalesce(func.sum(OrderItem.order_qty), 0).label("qty"),
            func.max(OrderItem.order_date).label("order_date"),
            func.max(OrderItem.delivery_date).label("delivery_date"),
        )
        .group_by(OrderItem.order_number, OrderItem.customer_name)
        .order_by(desc("qty"))
        .limit(top)
        .all()
    )
    by_order = [
        {
            "order_number": r.order_number or "(N/A)",
            "customer_name": r.customer_name,
            "line_count": r.line_count,
            "qty": int(r.qty),
            "order_date": r.order_date.isoformat() if r.order_date else None,
            "delivery_date": r.delivery_date.isoformat() if r.delivery_date else None,
        }
        for r in order_rows
    ]

    # --- Recent imports ---
    recent_q = (
        db.query(OrderItem)
        .order_by(OrderItem.id.desc())
        .limit(15)
        .all()
    )
    recent = [
        {
            "id": x.id,
            "order_number": x.order_number,
            "item_code": x.item_code,
            "customer_name": x.customer_name,
            "order_qty": x.order_qty,
            "product_image": x.product_image,
            "imported_at": x.imported_at.isoformat() if x.imported_at else None,
        }
        for x in recent_q
    ]

    return {
        "kpis": {
            "total_items": total_items,
            "distinct_orders": distinct_orders,
            "distinct_customers": distinct_customers,
            "distinct_items": distinct_items,
            "total_qty": total_qty,
            "max_id": max_id,
            "last_imported_at": last_import.isoformat() if last_import else None,
        },
        "by_customer": by_customer,
        "by_item": by_item,
        "by_order": by_order,
        "recent": recent,
    }


@app.get("/api/notifications")
def notifications(
    since_id: int = Query(0, ge=0, description="Return items with id > since_id"),
    db: Session = Depends(get_db),
):
    """Poll for new order lines since last seen id (for live notifications)."""
    max_id = db.query(func.max(OrderItem.id)).scalar() or 0
    new_count = 0
    newest = []
    if since_id < max_id:
        q = (
            db.query(OrderItem)
            .filter(OrderItem.id > since_id)
            .order_by(OrderItem.id.desc())
            .limit(20)
            .all()
        )
        new_count = (
            db.query(OrderItem).filter(OrderItem.id > since_id).count()
        )
        newest = [
            {
                "id": x.id,
                "order_number": x.order_number,
                "item_code": x.item_code,
                "customer_name": x.customer_name,
                "order_qty": x.order_qty,
            }
            for x in q
        ]
    return {
        "max_id": max_id,
        "new_count": new_count,
        "has_new": new_count > 0,
        "items": newest,
    }


# ---------- Run ----------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
