"""
Order Management Web API + UI
- Upload Excel order detail files (Chinese headers supported)
- Store in SQLite
- Bilingual UI (中文 / English / Tiếng Việt)
- REST API for programmatic access
"""
import os
import shutil
import tempfile
from typing import List, Optional, Any
from datetime import date, datetime
from collections import defaultdict

from fastapi import FastAPI, UploadFile, File, Depends, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse
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
    imported_count: int
    image_count: int = 0
    error: Optional[str] = None


class ImportResult(BaseModel):
    success: bool
    imported_count: int
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



def _import_bytes(content: bytes, filename: str, db: Session) -> FileImportDetail:
    """Parse one Excel from bytes, add rows to session. Caller commits."""
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

        items = rows_to_models(rows)
        for item in items:
            db.add(item)

        with_img = sum(1 for r in rows if r.get("product_image"))
        return FileImportDetail(
            filename=filename,
            imported_count=len(items),
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
    total = 0
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
            total += detail.imported_count
            total_img += detail.image_count
            details.append(detail)
            # free identity map (large base64 images)
            db.expunge_all()
        except Exception as e:
            db.rollback()
            details.append(FileImportDetail(
                filename=fname,
                imported_count=0,
                image_count=0,
                error=f"DB commit failed: {e}",
            ))
            errors += 1

    if total == 0 and errors == len(uploads):
        msgs = "; ".join(f"{d.filename}: {d.error}" for d in details if d.error)
        raise HTTPException(400, f"Import failed for all files. {msgs}")

    ok_files = len(uploads) - errors
    msg = f"Imported {total} order items from {ok_files}/{len(uploads)} file(s)"
    if total_img:
        msg += f" ({total_img} with images)"
    if errors:
        msg += f" — {errors} file(s) failed: "
        msg += "; ".join(f"{d.filename}: {d.error}" for d in details if d.error)

    return ImportResult(
        success=errors == 0,
        imported_count=total,
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
