from sqlalchemy import create_engine, Column, Integer, String, Float, Text, DateTime, Date
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime
import os

Base = declarative_base()

# DATABASE_URL:
#   - Local default: SQLite file data/orders.db
#   - Production (Supabase): postgresql+psycopg2://...
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

if not DATABASE_URL:
    _base = os.path.dirname(os.path.abspath(__file__))
    _db_path = os.path.join(_base, "data", "orders.db")
    os.makedirs(os.path.dirname(_db_path), exist_ok=True)
    DATABASE_URL = f"sqlite:///{_db_path}"

connect_args = {}
engine_kwargs = {}

if DATABASE_URL.startswith("sqlite"):
    connect_args = {"check_same_thread": False}
elif DATABASE_URL.startswith("postgres"):
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg2://", 1)
    elif DATABASE_URL.startswith("postgresql://") and "+psycopg2" not in DATABASE_URL:
        DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)
    if "sslmode" not in DATABASE_URL:
        sep = "&" if "?" in DATABASE_URL else "?"
        DATABASE_URL = f"{DATABASE_URL}{sep}sslmode=require"
    if ":6543" in DATABASE_URL:
        from sqlalchemy.pool import NullPool
        engine_kwargs["poolclass"] = NullPool
    else:
        engine_kwargs["pool_pre_ping"] = True
        engine_kwargs["pool_size"] = 5
        engine_kwargs["max_overflow"] = 10

engine = create_engine(DATABASE_URL, connect_args=connect_args, **engine_kwargs)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class OrderItem(Base):
    __tablename__ = "order_items"

    id = Column(Integer, primary_key=True, index=True)
    customer_name = Column(String(200))
    order_number = Column(String(100), index=True)
    order_date = Column(Date, nullable=True)
    delivery_date = Column(Date, nullable=True)
    customer_order_number = Column(String(100))
    product_image = Column(String(500), nullable=True)
    item_code = Column(String(100), index=True)
    customer_item_code = Column(String(100))
    sub_item_code = Column(String(100))
    main_category = Column(String(100))
    sub_category = Column(String(100))
    description = Column(Text)
    container_size = Column(String(100))
    container_process = Column(String(100))
    container_color = Column(String(200))
    order_qty = Column(Integer)
    retail_pack_rate = Column(Integer)
    unit_qty = Column(Integer)
    unit_wax_weight_g = Column(Float)
    fragrance_net_content_ml = Column(String(50))
    total_wax_weight_kg = Column(String(50))
    wax_material = Column(String(100))
    solid_or_bubble_wax = Column(String(50))
    wick_count = Column(Integer)
    wax_color = Column(String(50))
    lid_process = Column(String(100))
    fragrance_name = Column(String(200))
    fragrance_code = Column(String(200))
    fragrance_company = Column(String(100))
    fragrance_ratio = Column(Text)
    quality_requirement = Column(String(50))
    inspection_type = Column(String(100))
    inspection_requirement = Column(String(200))
    test_requirement = Column(String(200))
    sample_requirement = Column(String(200))
    salesperson = Column(String(100))
    remarks = Column(Text)
    packaging_detail = Column(Text)
    merchandiser = Column(String(100))
    outer_box_barcode = Column(String(100))
    outer_box_pack_rate = Column(Integer)
    inner_box_barcode = Column(String(100))
    retail_barcode = Column(String(100))
    imported_at = Column(DateTime, default=datetime.utcnow)
    source_file = Column(String(300), nullable=True)


def init_db():
    Base.metadata.create_all(bind=engine)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
