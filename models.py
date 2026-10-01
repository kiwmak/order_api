from sqlalchemy import create_engine, Column, Integer, String, Float, Text, DateTime, Date
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime
import os

Base = declarative_base()

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "orders.db")
engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class OrderItem(Base):
    __tablename__ = "order_items"

    id = Column(Integer, primary_key=True, index=True)
    # Customer / Order info
    customer_name = Column(String(200))          # 客户名称
    order_number = Column(String(100), index=True)  # 订单号码
    order_date = Column(Date, nullable=True)     # 下单日期
    delivery_date = Column(Date, nullable=True)  # 交货日期
    customer_order_number = Column(String(100))  # 客户订单号码

    # Product info
    product_image = Column(String(500), nullable=True)  # 产品图片
    item_code = Column(String(100), index=True)  # 货号
    customer_item_code = Column(String(100))     # 客号
    sub_item_code = Column(String(100))          # 子货号
    main_category = Column(String(100))          # 主类别
    sub_category = Column(String(100))           # 分类别
    description = Column(Text)                   # 描述
    container_size = Column(String(100))         # 容器尺寸(CM)
    container_process = Column(String(100))      # 容器工艺
    container_color = Column(String(200))        # 容器颜色及色卡号

    # Quantity / Specs
    order_qty = Column(Integer)                  # 订单数量
    retail_pack_rate = Column(Integer)           # 零售包装率
    unit_qty = Column(Integer)                   # 单个数量
    unit_wax_weight_g = Column(Float)            # 单个蜡重g
    fragrance_net_content_ml = Column(String(50))  # 香薰净含量ml
    total_wax_weight_kg = Column(String(50))     # 订单蜡总重量（kg）
    wax_material = Column(String(100))           # 蜡材
    solid_or_bubble_wax = Column(String(50))     # 实/泡蜡
    wick_count = Column(Integer)                 # 灯芯数量
    wax_color = Column(String(50))               # 蜡颜色
    lid_process = Column(String(100))            # 盖子工艺

    # Fragrance
    fragrance_name = Column(String(200))         # 香精名称
    fragrance_code = Column(String(200))         # 香精编号
    fragrance_company = Column(String(100))      # 香精公司
    fragrance_ratio = Column(Text)               # 香精比例

    # Quality / Inspection
    quality_requirement = Column(String(50))     # 质量要求
    inspection_type = Column(String(100))        # 验货类型
    inspection_requirement = Column(String(200)) # 验货要求
    test_requirement = Column(String(200))       # 测试要求
    sample_requirement = Column(String(200))     # 留样要求

    # People / Notes
    salesperson = Column(String(100))            # 业务员
    remarks = Column(Text)                       # 备注
    packaging_detail = Column(Text)              # 包装详细说明
    merchandiser = Column(String(100))           # 跟单

    # Barcodes / Packing
    outer_box_barcode = Column(String(100))      # 外箱条形码
    outer_box_pack_rate = Column(Integer)        # 外箱包装率
    inner_box_barcode = Column(String(100))      # 内盒条形码
    retail_barcode = Column(String(100))         # 零售条形码

    # Meta
    imported_at = Column(DateTime, default=datetime.utcnow)
    source_file = Column(String(300), nullable=True)


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    Base.metadata.create_all(bind=engine)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
