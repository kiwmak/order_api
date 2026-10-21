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
import math
import shutil
import tempfile
from typing import List, Optional, Any, Literal
from datetime import date, datetime
from collections import defaultdict

from fastapi import FastAPI, UploadFile, File, Depends, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from sqlalchemy import func, desc, cast, String
from sqlalchemy.exc import IntegrityError
from pydantic import BaseModel, Field, ValidationError
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from models import init_db, get_db, OrderItem, ScaleTicket, UserAccount, SessionLocal
from importer import parse_excel, rows_to_models
from translations import FIELD_LABELS, UI_TEXT
from auth import (
    verify_password,
    create_token,
    get_current_user,
    get_optional_user,
    get_user_role,
    ensure_admin_account,
    require_roles,
    set_user_password,
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


class OrderListResponse(BaseModel):
    total: int
    skip: int
    limit: int
    items: List[OrderItemOut]


class LoginRequest(BaseModel):
    username: str
    password: str


class LoginResponse(BaseModel):
    success: bool
    token: str
    username: str
    role: str
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


class ScaleTicketReading(BaseModel):
    spout: int = Field(ge=1, le=100)
    cup_weight_g: int = Field(ge=0)
    pour1_weight_g: int = Field(ge=0)
    pour2_weight_g: int = Field(ge=0)


class ScaleTicketReadingOut(BaseModel):
    spout: int
    cup_weight_g: Optional[int] = None
    pour1_weight_g: Optional[int] = None
    pour2_weight_g: Optional[int] = None


class ScaleTicketCreate(BaseModel):
    item_code: str = Field(min_length=1, max_length=100)
    order_number: Optional[str] = Field(None, max_length=100)
    spout_count: int = Field(ge=1, le=100)
    cup_min_g: int = Field(ge=0)
    cup_max_g: int = Field(ge=0)
    pour1_min_g: int = Field(ge=0)
    pour1_max_g: int = Field(ge=0)
    pour2_min_g: int = Field(ge=0)
    pour2_max_g: int = Field(ge=0)
    readings: List[ScaleTicketReading]
    notes: Optional[str] = Field(None, max_length=2000)


class ScaleTicketUIReading(BaseModel):
    spout: int = Field(ge=1, le=100)
    cup: Optional[int] = Field(None, ge=0)
    pour1: Optional[int] = Field(None, ge=0)
    pour2: Optional[int] = Field(None, ge=0)


class ScaleTicketUIInput(BaseModel):
    mode: Optional[Literal["actual"]] = None
    spout_count: int = Field(ge=1, le=100)
    cup_min: int = Field(ge=0)
    cup_max: int = Field(ge=0)
    pour1_min: int = Field(ge=0)
    pour1_max: int = Field(ge=0)
    pour2_min: int = Field(ge=0)
    pour2_max: int = Field(ge=0)
    readings: List[ScaleTicketUIReading]
    notes: Optional[str] = Field(None, max_length=2000)


class ScaleTicketOut(BaseModel):
    id: int
    item_code: str
    order_number: Optional[str] = None
    customer_name: Optional[str] = None
    description: Optional[str] = None
    spout_count: int
    cup_min_g: int
    cup_max_g: int
    pour1_min_g: int
    pour1_max_g: int
    pour2_min_g: int
    pour2_max_g: int
    readings: List[ScaleTicketReadingOut]
    average_wax_weight_g: Optional[int] = None
    notes: Optional[str] = None
    created_at: datetime
    mode: str = "hypothetical"


class ScaleTicketUIOut(BaseModel):
    id: int
    item_code: str
    order_number: Optional[str] = None
    customer_name: Optional[str] = None
    description: Optional[str] = None
    spout_count: int
    cup_min: int
    cup_max: int
    pour1_min: int
    pour1_max: int
    pour2_min: int
    pour2_max: int
    readings: List[ScaleTicketUIReading]
    average_wax_weight_g: Optional[int] = None
    notes: Optional[str] = None
    created_at: datetime
    mode: str


def _validate_scale_ticket(body: ScaleTicketCreate) -> int:
    ranges = (
        (body.cup_min_g, body.cup_max_g),
        (body.pour1_min_g, body.pour1_max_g),
        (body.pour2_min_g, body.pour2_max_g),
    )
    if any(minimum > maximum for minimum, maximum in ranges):
        raise ValueError("Each minimum weight must be less than or equal to its maximum")
    if len(body.readings) != body.spout_count:
        raise ValueError("Reading count must match spout count")
    for index, reading in enumerate(body.readings, start=1):
        if reading.spout != index:
            raise ValueError("Spout numbers must be consecutive and start at 1")
        if not body.cup_min_g <= reading.cup_weight_g <= body.cup_max_g:
            raise ValueError("Cup weight is outside the saved range")
        if not body.pour1_min_g <= reading.pour1_weight_g <= body.pour1_max_g:
            raise ValueError("Pour 1 weight is outside the saved range")
        if not body.pour2_min_g <= reading.pour2_weight_g <= body.pour2_max_g:
            raise ValueError("Pour 2 weight is outside the saved range")
    wax_weights = [row.pour1_weight_g + row.pour2_weight_g for row in body.readings]
    return math.floor(sum(wax_weights) / len(wax_weights) + 0.5)


def _scale_ticket_response(ticket: ScaleTicket) -> ScaleTicketOut:
    readings_data, mode = _scale_ticket_readings(ticket)
    return ScaleTicketOut(
        id=ticket.id,
        item_code=ticket.item_code,
        order_number=ticket.order_number,
        customer_name=ticket.customer_name,
        description=ticket.description,
        spout_count=ticket.spout_count,
        cup_min_g=ticket.cup_min_g,
        cup_max_g=ticket.cup_max_g,
        pour1_min_g=ticket.pour1_min_g,
        pour1_max_g=ticket.pour1_max_g,
        pour2_min_g=ticket.pour2_min_g,
        pour2_max_g=ticket.pour2_max_g,
        readings=[
            ScaleTicketReadingOut(
                spout=row["spout"],
                cup_weight_g=row.get("cup", row.get("cup_weight_g")),
                pour1_weight_g=row.get("pour1", row.get("pour1_weight_g")),
                pour2_weight_g=row.get("pour2", row.get("pour2_weight_g")),
            )
            for row in readings_data
        ],
        average_wax_weight_g=_scale_ticket_average(ticket, readings_data, mode),
        notes=ticket.notes,
        created_at=ticket.created_at,
        mode=mode,
    )


def _scale_ticket_readings(ticket: ScaleTicket):
    data = json.loads(ticket.readings_json)
    if isinstance(data, dict):
        return data.get("readings", []), data.get("mode", "hypothetical")
    return data, "hypothetical"


def _scale_ticket_average(ticket: ScaleTicket, readings, mode: str):
    if mode == "actual":
        wax_weights = [
            row["pour2"] - row["cup"]
            for row in readings
            if row.get("cup") is not None and row.get("pour2") is not None
        ]
        return math.floor(sum(wax_weights) / len(wax_weights) + 0.5) if wax_weights else None
    return ticket.average_wax_weight_g


def _scale_ticket_ui_response(ticket: ScaleTicket) -> ScaleTicketUIOut:
    readings, mode = _scale_ticket_readings(ticket)
    normalized_readings = [
        ScaleTicketUIReading(
            spout=row["spout"],
            cup=row.get("cup", row.get("cup_weight_g")),
            pour1=row.get("pour1", row.get("pour1_weight_g")),
            pour2=row.get("pour2", row.get("pour2_weight_g")),
        )
        for row in readings
    ]
    return ScaleTicketUIOut(
        id=ticket.id,
        item_code=ticket.item_code,
        order_number=ticket.order_number,
        customer_name=ticket.customer_name,
        description=ticket.description,
        spout_count=ticket.spout_count,
        cup_min=ticket.cup_min_g,
        cup_max=ticket.cup_max_g,
        pour1_min=ticket.pour1_min_g,
        pour1_max=ticket.pour1_max_g,
        pour2_min=ticket.pour2_min_g,
        pour2_max=ticket.pour2_max_g,
        readings=normalized_readings,
        average_wax_weight_g=_scale_ticket_average(ticket, readings, mode),
        notes=ticket.notes,
        created_at=ticket.created_at,
        mode=mode,
    )


# ---------- Startup ----------
@app.on_event("startup")
def on_startup():
    init_db()
    ensure_admin_account()


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
    username = body.username.strip()
    if not verify_password(username, body.password):
        raise HTTPException(401, "Invalid username or password")
    role = get_user_role(username)
    if not role:
        raise HTTPException(401, "Invalid or inactive user")
    token = create_token(username)
    return LoginResponse(
        success=True,
        token=token,
        username=username,
        role=role,
        message="Login OK",
    )


@app.get("/api/me")
def me(user: Optional[str] = Depends(get_optional_user)):
    if not user:
        return {"authenticated": False, "username": None, "role": None}
    return {"authenticated": True, "username": user, "role": get_user_role(user)}


class UserCreateRequest(BaseModel):
    username: str = Field(min_length=1, max_length=100)
    password: str = Field(min_length=8, max_length=200)
    role: Literal["admin", "staff", "viewer"]


class UserUpdateRequest(BaseModel):
    role: Optional[Literal["admin", "staff", "viewer"]] = None
    is_active: Optional[bool] = None
    password: Optional[str] = Field(None, min_length=8, max_length=200)


def _user_response(account: UserAccount) -> dict:
    return {
        "username": account.username,
        "role": account.role,
        "is_active": account.is_active,
        "created_at": account.created_at,
    }


@app.get("/api/users")
def list_users(
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin")),
):
    return [_user_response(account) for account in db.query(UserAccount).order_by(UserAccount.username).all()]


@app.post("/api/users", status_code=201)
def create_user(
    body: UserCreateRequest,
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin")),
):
    username = body.username.strip()
    if not username:
        raise HTTPException(400, "Username is required")
    if db.query(UserAccount).filter(UserAccount.username == username).first():
        raise HTTPException(409, "Username already exists")
    account = UserAccount(username=username, salt="", password_hash="", role=body.role, is_active=True)
    set_user_password(account, body.password)
    db.add(account)
    try:
        db.commit()
    except IntegrityError as exc:
        db.rollback()
        raise HTTPException(409, "Username already exists") from exc
    db.refresh(account)
    return _user_response(account)


@app.put("/api/users/{username}")
def update_user(
    username: str,
    body: UserUpdateRequest,
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin")),
):
    account = db.query(UserAccount).filter(UserAccount.username == username).first()
    if not account:
        raise HTTPException(404, "User not found")
    updated_role = body.role if body.role is not None else account.role
    updated_active = body.is_active if body.is_active is not None else account.is_active
    removes_admin = account.role == "admin" and account.is_active and (
        updated_role != "admin" or not updated_active
    )
    if removes_admin:
        active_admins = db.query(UserAccount).filter(
            UserAccount.role == "admin",
            UserAccount.is_active.is_(True),
        ).count()
        if active_admins <= 1:
            raise HTTPException(400, "At least one active administrator is required")
    account.role = updated_role
    account.is_active = updated_active
    if body.password is not None:
        set_user_password(account, body.password)
    db.commit()
    db.refresh(account)
    return _user_response(account)


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
            print(f"[upsert] item_code={item.item_code} product_image={str(item.product_image)[:90]}")

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
                error="No valid order rows found (need order_number or item_code). Check Excel headers match mapping.",
            )
        print(f"[import] {filename}: parsed {len(rows)} rows, images={sum(1 for r in rows if r.get('product_image'))}")

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
    user: str = Depends(require_roles("admin")),
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
def _apply_order_search(query, q: Optional[str]):
    if q:
        like = f"%{q}%"
        query = query.filter(
            (OrderItem.order_number.ilike(like))
            | (OrderItem.item_code.ilike(like))
            | (OrderItem.customer_name.ilike(like))
            | (OrderItem.customer_item_code.ilike(like))
            | (OrderItem.description.ilike(like))
        )
    return query


def _decode_column_filters(raw: Optional[str]) -> dict:
    if not raw:
        return {}
    try:
        filters = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise HTTPException(400, "Invalid filters JSON") from exc
    if not isinstance(filters, dict):
        raise HTTPException(400, "Filters must be an object")

    valid_columns = set(OrderItem.__table__.columns.keys())
    decoded = {}
    for field, values in filters.items():
        if field not in valid_columns or field == "id":
            raise HTTPException(400, f"Invalid filter field: {field}")
        if not isinstance(values, list):
            raise HTTPException(400, f"Filter values for {field} must be a list")
        decoded[field] = [str(value) for value in values]
    return decoded


def _apply_column_filters(query, filters: dict, exclude_field: Optional[str] = None):
    for field, values in filters.items():
        if field == exclude_field:
            continue
        column = getattr(OrderItem, field)
        if not values:
            query = query.filter(False)
        else:
            query = query.filter(func.coalesce(cast(column, String), "").in_(values))
    return query


@app.get("/api/orders/filter-values/{field}")
def get_order_filter_values(
    field: str,
    q: Optional[str] = Query(None),
    filters: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    if field not in OrderItem.__table__.columns or field == "id":
        raise HTTPException(404, "Filter field not found")

    query = _apply_order_search(db.query(OrderItem), q)
    query = _apply_column_filters(query, _decode_column_filters(filters), exclude_field=field)
    column = getattr(OrderItem, field)
    rows = query.with_entities(column, func.count(OrderItem.id)).group_by(column).all()
    items = [{"value": "" if value is None else str(value), "count": count} for value, count in rows]
    return {"items": items}


@app.get("/api/orders", response_model=OrderListResponse)
def list_orders(
    q: Optional[str] = Query(None, description="Search order_number / item_code / customer_name"),
    order_number: Optional[str] = Query(None),
    filters: Optional[str] = Query(None),
    skip: int = Query(0, ge=0),
    limit: int = Query(100, ge=1, le=500),
    include_images: bool = Query(False, description="Include full product_image (heavy). Default false for list."),
    db: Session = Depends(get_db),
):
    """Paginated order list. Default limit 100, max 500. Total count always returned."""
    query = db.query(OrderItem)
    if order_number:
        query = query.filter(OrderItem.order_number == order_number)
    query = _apply_order_search(query, q)
    query = _apply_column_filters(query, _decode_column_filters(filters))
    total = query.count()
    items = query.order_by(OrderItem.order_number.desc(), OrderItem.id.desc()).offset(skip).limit(limit).all()

    # Lighten payload: base64 images are huge — mark presence only unless requested
    out = []
    for x in items:
        row = OrderItemOut.model_validate(x) if hasattr(OrderItemOut, "model_validate") else OrderItemOut.from_orm(x)
        if not include_images:
            pi = row.product_image
            # only strip heavy base64; keep http(s) Storage URLs for thumbnails
            if pi and str(pi).startswith("data:"):
                row.product_image = "__has_image__"
        out.append(row)

    return OrderListResponse(total=total, skip=skip, limit=limit, items=out)


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
    user: str = Depends(require_roles("admin")),
):
    count = db.query(OrderItem).delete()
    db.commit()
    return {"deleted": count}


@app.delete("/api/orders/{item_id}")
def delete_item(
    item_id: int,
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin")),
):
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Not found")
    db.delete(item)
    db.commit()
    return {"deleted": item_id}


def _validate_scale_ticket_ui(body: ScaleTicketUIInput) -> int:
    ranges = (
        (body.cup_min, body.cup_max),
        (body.pour1_min, body.pour1_max),
        (body.pour2_min, body.pour2_max),
    )
    if any(minimum > maximum for minimum, maximum in ranges):
        raise HTTPException(400, "Each minimum weight must be less than or equal to its maximum")
    if len(body.readings) != body.spout_count:
        raise HTTPException(400, "Reading count must match spout count")
    if body.mode == "actual" and not any(
        reading.cup is not None or reading.pour1 is not None or reading.pour2 is not None
        for reading in body.readings
    ):
        raise HTTPException(400, "Enter at least one reading before saving")

    for index, reading in enumerate(body.readings, start=1):
        if reading.spout != index:
            raise HTTPException(400, "Spout numbers must be consecutive and start at 1")
        values = (reading.cup, reading.pour1, reading.pour2)
        if body.mode != "actual" and any(value is None for value in values):
            raise HTTPException(400, "Hypothetical readings must include all weights")
        if body.mode == "actual":
            if reading.cup is not None and reading.pour1 is not None and reading.pour1 < reading.cup:
                raise HTTPException(400, "First pour must be greater than or equal to cup tare")
            if reading.pour1 is not None and reading.pour2 is not None and reading.pour2 < reading.pour1:
                raise HTTPException(400, "Final pour must be greater than or equal to first pour")
            if reading.cup is not None and reading.pour2 is not None and reading.pour2 < reading.cup:
                raise HTTPException(400, "Final pour must be greater than or equal to cup tare")
        else:
            for value, (minimum, maximum) in zip(values, ranges):
                if value is not None and not minimum <= value <= maximum:
                    raise HTTPException(400, "Reading is outside the saved range")

    if body.mode == "actual":
        wax_weights = [
            reading.pour2 - reading.cup
            for reading in body.readings
            if reading.cup is not None and reading.pour2 is not None
        ]
    else:
        wax_weights = [reading.pour1 + reading.pour2 for reading in body.readings]
    return math.floor(sum(wax_weights) / len(wax_weights) + 0.5) if wax_weights else 0


def _save_scale_ticket_ui(item: OrderItem, body: ScaleTicketUIInput, ticket: Optional[ScaleTicket] = None):
    average_weight = _validate_scale_ticket_ui(body)
    if ticket is None:
        ticket = ScaleTicket(
            item_code=item.item_code,
            order_number=item.order_number,
            customer_name=item.customer_name,
            description=item.description,
        )
    ticket.spout_count = body.spout_count
    ticket.cup_min_g = body.cup_min
    ticket.cup_max_g = body.cup_max
    ticket.pour1_min_g = body.pour1_min
    ticket.pour1_max_g = body.pour1_max
    ticket.pour2_min_g = body.pour2_min
    ticket.pour2_max_g = body.pour2_max
    ticket.readings_json = json.dumps(
        {
            "mode": body.mode or "hypothetical",
            "readings": [
                reading.model_dump() if hasattr(reading, "model_dump") else reading.dict()
                for reading in body.readings
            ],
        },
        ensure_ascii=False,
    )
    ticket.average_wax_weight_g = average_weight
    ticket.notes = body.notes
    return ticket


@app.get("/api/orders/{item_id}/scale-tickets", response_model=List[ScaleTicketUIOut])
def list_order_scale_tickets(item_id: int, db: Session = Depends(get_db)):
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Order item not found")
    tickets = (
        db.query(ScaleTicket)
        .filter(
            ScaleTicket.item_code == item.item_code,
            ScaleTicket.order_number == item.order_number,
        )
        .order_by(ScaleTicket.created_at.desc(), ScaleTicket.id.desc())
        .all()
    )
    return [_scale_ticket_ui_response(ticket) for ticket in tickets]


@app.post("/api/orders/{item_id}/scale-tickets", response_model=ScaleTicketUIOut)
def create_order_scale_ticket(
    item_id: int,
    body: ScaleTicketUIInput,
    db: Session = Depends(get_db),
    user: Optional[str] = Depends(get_optional_user),
):
    if body.mode == "actual":
        if not user:
            raise HTTPException(401, "Login required")
        if get_user_role(user) not in ("admin", "staff"):
            raise HTTPException(403, "Insufficient permissions")
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Order item not found")
    ticket = _save_scale_ticket_ui(item, body)
    db.add(ticket)
    db.commit()
    db.refresh(ticket)
    return _scale_ticket_ui_response(ticket)


@app.put("/api/orders/{item_id}/scale-tickets/{ticket_id}", response_model=ScaleTicketUIOut)
def update_order_scale_ticket(
    item_id: int,
    ticket_id: int,
    body: ScaleTicketUIInput,
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin", "staff")),
):
    if body.mode != "actual":
        raise HTTPException(400, "Only actual weighing tickets can be edited")
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Order item not found")
    ticket = (
        db.query(ScaleTicket)
        .filter(
            ScaleTicket.id == ticket_id,
            ScaleTicket.item_code == item.item_code,
            ScaleTicket.order_number == item.order_number,
        )
        .first()
    )
    if not ticket:
        raise HTTPException(404, "Scale ticket not found")
    _save_scale_ticket_ui(item, body, ticket)
    db.commit()
    db.refresh(ticket)
    return _scale_ticket_ui_response(ticket)


@app.post("/api/orders/{item_id}/scale-tickets/export")
def export_order_scale_ticket(
    item_id: int,
    body: ScaleTicketUIInput,
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin", "staff", "viewer")),
):
    if body.mode != "actual":
        raise HTTPException(400, "Only actual weighing tickets can be exported")
    _validate_scale_ticket_ui(body)
    item = db.query(OrderItem).filter(OrderItem.id == item_id).first()
    if not item:
        raise HTTPException(404, "Order item not found")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Phieu can thuc te"
    target_weight = item.unit_wax_weight_g
    has_target_weight = target_weight is not None and math.isfinite(target_weight) and target_weight > 0
    sheet.append(["Phiếu cân thực tế"])
    sheet.append(["Mã hàng", item.item_code or ""])
    sheet.append(["Đơn hàng", item.order_number or ""])
    sheet.append(["Trọng lượng mã hàng (g)", target_weight if has_target_weight else ""])
    sheet.append([])
    headers = [
        "STT / Vòi",
        "Vỏ cốc (g)",
        "Rót lần đầu (gồm cốc+sáp) (g)",
        "Sáp lần 1 (g)",
        "Rót lần cuối (gồm cốc+sáp) (g)",
        "Sáp rót thêm (g)",
        "Trọng lượng sáp nến (g)",
    ]
    sheet.append(headers)
    header_row = sheet.max_row
    for reading in body.readings:
        cup = reading.cup
        pour1 = reading.pour1
        pour2 = reading.pour2
        sheet.append([
            reading.spout,
            cup,
            pour1,
            pour1 - cup if pour1 is not None and cup is not None else None,
            pour2,
            pour2 - pour1 if pour2 is not None and pour1 is not None else None,
            pour2 - cup if pour2 is not None and cup is not None else None,
        ])
        wax_cell = sheet.cell(sheet.max_row, 7)
        wax_weight = wax_cell.value
        if has_target_weight and wax_weight is not None:
            if wax_weight < target_weight:
                wax_cell.fill = PatternFill(fill_type="solid", fgColor="FECACA")
                wax_cell.font = Font(color="DC2626", bold=True)
            elif wax_weight > target_weight * 1.05:
                wax_cell.fill = PatternFill(fill_type="solid", fgColor="FEF3C7")
                wax_cell.font = Font(color="A16207", bold=True)
            elif wax_weight > target_weight:
                wax_cell.fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
                wax_cell.font = Font(color="15803D", bold=True)

    readings_start = header_row + 1
    readings_end = sheet.max_row
    complete_wax_count = sum(
        reading.cup is not None and reading.pour2 is not None
        for reading in body.readings
    )
    average = _validate_scale_ticket_ui(body)
    sheet.append(["", "", "", "", "", "Trọng lượng trung bình (g)", average if complete_wax_count else None])
    for cell in sheet[sheet.max_row]:
        cell.font = Font(bold=True)
    for cell in sheet[header_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="334155")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in sheet.iter_rows(min_row=readings_start, max_row=readings_end, min_col=1, max_col=7):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    column_widths = [12, 16, 32, 16, 32, 18, 24]
    for index, width in enumerate(column_widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = f"A{readings_start}"
    sheet.auto_filter.ref = f"A{header_row}:G{readings_end}"
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="phieu_can_{item.id}.xlsx"'
        },
    )


@app.put("/api/orders/{item_id}", response_model=OrderItemOut)
def update_item(
    item_id: int,
    body: OrderItemUpdate,
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin")),
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


@app.get("/api/scale-tickets/{item_code}", response_model=List[ScaleTicketOut])
def list_scale_tickets(
    item_code: str,
    order_number: Optional[str] = Query(None),
    db: Session = Depends(get_db),
):
    query = (
        db.query(ScaleTicket)
        .filter(ScaleTicket.item_code == item_code)
    )
    if order_number:
        query = query.filter(ScaleTicket.order_number == order_number)
    tickets = query.order_by(ScaleTicket.created_at.desc(), ScaleTicket.id.desc()).all()
    return [_scale_ticket_response(ticket) for ticket in tickets]


@app.post("/api/scale-tickets", response_model=ScaleTicketOut)
def create_scale_ticket(
    body: ScaleTicketCreate,
    db: Session = Depends(get_db),
    user: Optional[str] = Depends(get_optional_user),
):
    item_code = body.item_code.strip()
    if not item_code:
        raise HTTPException(400, "Item code is required")
    item_query = db.query(OrderItem).filter(OrderItem.item_code == item_code)
    if body.order_number:
        item_query = item_query.filter(OrderItem.order_number == body.order_number)
    item = item_query.order_by(OrderItem.id.desc()).first()
    if not item:
        raise HTTPException(404, "Product code not found")
    try:
        average_weight = _validate_scale_ticket(body)
    except ValueError as e:
        raise HTTPException(400, str(e))

    ticket = ScaleTicket(
        item_code=item.item_code,
        order_number=item.order_number,
        customer_name=item.customer_name,
        description=item.description,
        spout_count=body.spout_count,
        cup_min_g=body.cup_min_g,
        cup_max_g=body.cup_max_g,
        pour1_min_g=body.pour1_min_g,
        pour1_max_g=body.pour1_max_g,
        pour2_min_g=body.pour2_min_g,
        pour2_max_g=body.pour2_max_g,
        readings_json=json.dumps(
            [reading.model_dump() if hasattr(reading, "model_dump") else reading.dict() for reading in body.readings],
            ensure_ascii=False,
        ),
        average_wax_weight_g=average_weight,
        notes=body.notes,
    )
    db.add(ticket)
    db.commit()
    db.refresh(ticket)
    return _scale_ticket_response(ticket)


@app.delete("/api/scale-tickets/{ticket_id}")
def delete_scale_ticket(
    ticket_id: int,
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin")),
):
    ticket = db.query(ScaleTicket).filter(ScaleTicket.id == ticket_id).first()
    if not ticket:
        raise HTTPException(404, "Scale ticket not found")
    db.delete(ticket)
    db.commit()
    return {"deleted": ticket_id}


def _scale_ticket_backup(ticket: ScaleTicket) -> dict:
    readings, mode = _scale_ticket_readings(ticket)
    return {
        "item_code": ticket.item_code,
        "order_number": ticket.order_number,
        "customer_name": ticket.customer_name,
        "description": ticket.description,
        "spout_count": ticket.spout_count,
        "cup_min_g": ticket.cup_min_g,
        "cup_max_g": ticket.cup_max_g,
        "pour1_min_g": ticket.pour1_min_g,
        "pour1_max_g": ticket.pour1_max_g,
        "pour2_min_g": ticket.pour2_min_g,
        "pour2_max_g": ticket.pour2_max_g,
        "readings": [
            {
                "spout": row["spout"],
                "cup_weight_g": row.get("cup", row.get("cup_weight_g")),
                "pour1_weight_g": row.get("pour1", row.get("pour1_weight_g")),
                "pour2_weight_g": row.get("pour2", row.get("pour2_weight_g")),
            }
            for row in readings
        ],
        "average_wax_weight_g": _scale_ticket_average(ticket, readings, mode),
        "mode": mode,
        "notes": ticket.notes,
        "created_at": ticket.created_at.isoformat() if ticket.created_at else None,
    }



# ---------- Backup / Restore ----------
@app.get("/api/backup")
def backup_database(
    db: Session = Depends(get_db),
    user: str = Depends(require_roles("admin")),
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
        "scale_tickets": [
            _scale_ticket_backup(ticket)
            for ticket in db.query(ScaleTicket).order_by(ScaleTicket.id).all()
        ],
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
    user: str = Depends(require_roles("admin")),
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
    scale_tickets = data.get("scale_tickets", []) if isinstance(data, dict) else []
    if not isinstance(scale_tickets, list):
        raise HTTPException(400, "Backup scale_tickets must be an array")

    mode = (mode or "replace").lower().strip()
    if mode not in ("replace", "merge"):
        raise HTTPException(400, "mode must be replace or merge")

    deleted = 0
    if mode == "replace":
        deleted = db.query(OrderItem).delete()
        db.query(ScaleTicket).delete()
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

    db.flush()
    restored_tickets = 0
    for row in scale_tickets:
        if not isinstance(row, dict):
            errors += 1
            continue
        try:
            ticket_mode = row.get("mode", "hypothetical")
            if ticket_mode == "actual":
                item_code = str(row.get("item_code") or "").strip()
                ticket_order_number = row.get("order_number")
                ticket_data = ScaleTicketUIInput(
                    mode="actual",
                    spout_count=row.get("spout_count"),
                    cup_min=row.get("cup_min", row.get("cup_min_g")),
                    cup_max=row.get("cup_max", row.get("cup_max_g")),
                    pour1_min=row.get("pour1_min", row.get("pour1_min_g")),
                    pour1_max=row.get("pour1_max", row.get("pour1_max_g")),
                    pour2_min=row.get("pour2_min", row.get("pour2_min_g")),
                    pour2_max=row.get("pour2_max", row.get("pour2_max_g")),
                    readings=[
                        {
                            "spout": reading["spout"],
                            "cup": reading.get("cup", reading.get("cup_weight_g")),
                            "pour1": reading.get("pour1", reading.get("pour1_weight_g")),
                            "pour2": reading.get("pour2", reading.get("pour2_weight_g")),
                        }
                        for reading in row.get("readings", [])
                    ],
                    notes=row.get("notes"),
                )
            else:
                ticket_data = ScaleTicketCreate(**row)
                item_code = ticket_data.item_code
                ticket_order_number = ticket_data.order_number
                average_weight = _validate_scale_ticket(ticket_data)
            item_query = db.query(OrderItem).filter(OrderItem.item_code == item_code)
            if ticket_order_number:
                item_query = item_query.filter(OrderItem.order_number == ticket_order_number)
            item = item_query.first()
            if not item:
                errors += 1
                continue
            if ticket_mode == "actual":
                ticket = _save_scale_ticket_ui(item, ticket_data)
            else:
                ticket = ScaleTicket(
                    item_code=item.item_code,
                    order_number=row.get("order_number"),
                    customer_name=row.get("customer_name"),
                    description=row.get("description"),
                    spout_count=ticket_data.spout_count,
                    cup_min_g=ticket_data.cup_min_g,
                    cup_max_g=ticket_data.cup_max_g,
                    pour1_min_g=ticket_data.pour1_min_g,
                    pour1_max_g=ticket_data.pour1_max_g,
                    pour2_min_g=ticket_data.pour2_min_g,
                    pour2_max_g=ticket_data.pour2_max_g,
                    readings_json=json.dumps(
                        [reading.model_dump() if hasattr(reading, "model_dump") else reading.dict() for reading in ticket_data.readings],
                        ensure_ascii=False,
                    ),
                    average_wax_weight_g=average_weight,
                    notes=ticket_data.notes,
                ),
            ticket.customer_name = row.get("customer_name") or item.customer_name
            ticket.description = row.get("description") or item.description
            ticket.created_at = _dt(row.get("created_at")) or datetime.utcnow()
            db.add(ticket)
            restored_tickets += 1
        except (HTTPException, ValidationError, ValueError, TypeError, KeyError):
            errors += 1

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
        "restored_scale_tickets": restored_tickets,
        "errors": errors,
        "message": f"Restore OK ({mode}): {inserted} new, {updated} updated" + (f", cleared {deleted}" if deleted else ""),
    }



@app.get("/api/storage/status")
def storage_status(user: str = Depends(require_roles("admin"))):
    """Check whether Supabase Storage is configured (no secrets returned)."""
    from storage import storage_configured, _bucket
    import os
    url = (os.environ.get("SUPABASE_URL") or "").strip()
    return {
        "configured": storage_configured(),
        "bucket": _bucket() if storage_configured() else None,
        "supabase_url_set": bool(url),
        "hint": None if storage_configured() else "Set SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, SUPABASE_BUCKET on the server",
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
            .all()
        )

        grouped = {}
        for x in q:
            order_number = x.order_number or "UNKNOWN"
            group = grouped.setdefault(
                order_number,
                {
                    "order_number": order_number,
                    "customer_name": x.customer_name,
                    "item_count": 0,
                    "total_qty": 0,
                },
            )
            group["item_count"] += 1
            group["total_qty"] += int(x.order_qty or 0)
            if not group["customer_name"] and x.customer_name:
                group["customer_name"] = x.customer_name

        newest = [
            {
                "order_number": v["order_number"],
                "customer_name": v["customer_name"],
                "item_count": v["item_count"],
                "total_qty": v["total_qty"],
            }
            for v in grouped.values()
        ]
        newest.sort(key=lambda x: x["order_number"], reverse=True)
        new_count = len(newest)
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
