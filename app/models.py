from datetime import datetime
from decimal import Decimal

from flask_sqlalchemy import SQLAlchemy


db = SQLAlchemy()


class TimestampMixin:
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )


class Company(TimestampMixin, db.Model):
    __tablename__ = "companies"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False, index=True)
    address = db.Column(db.Text, default="", nullable=False)
    phone = db.Column(db.String(100), default="", nullable=False)
    tax_id = db.Column(db.String(50), default="", nullable=False)
    bank_name = db.Column(db.String(255), default="", nullable=False)
    bank_branch = db.Column(db.String(255), default="", nullable=False)
    account_no = db.Column(db.String(100), default="", nullable=False)
    account_name = db.Column(db.String(255), default="", nullable=False)
    business_type = db.Column(
        db.String(255),
        default="ขายส่ง,ขายปลีก,ให้บริการ",
        nullable=False,
    )
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)


class Unit(TimestampMixin, db.Model):
    __tablename__ = "units"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True, index=True)
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)


class Item(TimestampMixin, db.Model):
    __tablename__ = "items"

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(100), default="", nullable=False, index=True)
    name = db.Column(db.String(500), nullable=False, index=True)
    specification = db.Column(db.Text, default="", nullable=False)
    default_price = db.Column(
        db.Numeric(14, 2),
        default=Decimal("0.00"),
        nullable=False,
    )
    unit_id = db.Column(db.Integer, db.ForeignKey("units.id"), nullable=False)
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)

    unit = db.relationship("Unit")


class GovernmentProfile(TimestampMixin, db.Model):
    __tablename__ = "government_profiles"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(
        db.String(255),
        default="ข้อมูลโรงพยาบาลสิงห์บุรี",
        nullable=False,
        index=True,
    )
    department = db.Column(
        db.Text,
        default=(
            "โรงพยาบาลสิงห์บุรี กลุ่มงานเภสัชกรรม "
            "โทร. ๐ ๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙"
        ),
        nullable=False,
    )
    letter_prefix = db.Column(
        db.String(255),
        default="สห ๐๐๓๓.๒๐๕.๑๒/",
        nullable=False,
    )
    recipient = db.Column(
        db.String(255),
        default="ผู้ว่าราชการจังหวัดสิงห์บุรี",
        nullable=False,
    )
    officer_name = db.Column(
        db.String(255),
        default="นางพิณนภา ศริพันธุ์",
        nullable=False,
    )
    officer_position = db.Column(
        db.String(255),
        default="เจ้าหน้าที่",
        nullable=False,
    )
    chief_name = db.Column(
        db.String(255),
        default="นายชัชวาลย์ บุญญฤทธิ์",
        nullable=False,
    )
    chief_position = db.Column(
        db.String(255),
        default="เภสัชกรชำนาญการพิเศษ",
        nullable=False,
    )
    approver_name = db.Column(
        db.String(255),
        default="นายพิรุณ ปิตะหงษ์นันท์",
        nullable=False,
    )
    approver_position = db.Column(
        db.Text,
        default=(
            "ผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติราชการแทน "
            "ผู้ว่าราชการจังหวัดสิงห์บุรี"
        ),
        nullable=False,
    )
    inspector1_name = db.Column(
        db.String(255),
        default="นางสาวกัญญพัชร ธนกิจการค้า",
        nullable=False,
    )
    inspector1_position = db.Column(
        db.String(255),
        default="เภสัชกร",
        nullable=False,
    )
    inspector2_name = db.Column(
        db.String(255),
        default="นางสาวชุลีพร สุขมี",
        nullable=False,
    )
    inspector2_position = db.Column(
        db.String(255),
        default="เจ้าพนักงานเภสัชกรรมชำนาญงาน",
        nullable=False,
    )
    inspector3_name = db.Column(
        db.String(255),
        default="นางสาวกัญญาพัชร เลิศอนันตกูล",
        nullable=False,
    )
    inspector3_position = db.Column(
        db.String(255),
        default="เจ้าพนักงานเภสัชกรรม",
        nullable=False,
    )
    specifier_name = db.Column(
        db.String(255),
        default="นางสาวนลินี เครือทิวา",
        nullable=False,
    )
    specifier_position = db.Column(
        db.String(255),
        default="เภสัชกรชำนาญการ",
        nullable=False,
    )
    receipt1_name = db.Column(
        db.String(255),
        default="นางสาวธัญรดา ใจเสน",
        nullable=False,
    )
    receipt2_name = db.Column(
        db.String(255),
        default="นางสาวจุฑามาศ อาคมสรรเสริญ",
        nullable=False,
    )
    receipt3_name = db.Column(
        db.String(255),
        default="นางสาวนริศรา ม่วงงาม",
        nullable=False,
    )
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)


class DropdownOption(TimestampMixin, db.Model):
    __tablename__ = "dropdown_options"

    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(100), nullable=False, index=True)
    label = db.Column(db.String(500), nullable=False, default="")
    value = db.Column(db.String(500), nullable=False)
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)

    def __init__(self, **kwargs):
        if not kwargs.get("label") and kwargs.get("value"):
            kwargs["label"] = kwargs["value"]
        super().__init__(**kwargs)

    def __repr__(self):
        return f"<DropdownOption {self.category}: {self.label or self.value}>"


class Purchase(TimestampMixin, db.Model):
    __tablename__ = "purchases"

    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(100), nullable=False, unique=True, index=True)
    document_date = db.Column(db.Date, nullable=False, index=True)
    company_id = db.Column(db.Integer, db.ForeignKey("companies.id"), nullable=False)
    government_profile_id = db.Column(
        db.Integer,
        db.ForeignKey("government_profiles.id"),
        nullable=True,
    )
    project_number = db.Column(db.String(100), default="", nullable=False)
    contract_control_number = db.Column(db.String(100), default="", nullable=False)
    delivery_days = db.Column(db.Integer, default=30, nullable=False)
    delivery_place = db.Column(
        db.String(500),
        default="โรงพยาบาลสิงห์บุรี ๙๑๗/๓",
        nullable=False,
    )
    budget_source = db.Column(
        db.String(500),
        default=(
            "เงินนอกงบประมาณจาก "
            "เงินบำรุงโรงพยาบาลสิงห์บุรี ปี ๒๕๖๙"
        ),
        nullable=False,
    )
    procurement_type = db.Column(
        db.String(255),
        default="เวชภัณฑ์มิใช่ยา",
        nullable=False,
    )
    necessity_reason = db.Column(
        db.Text,
        default="ใช้ในการรักษาผู้ป่วย",
        nullable=False,
    )
    budget_allocated = db.Column(
        db.Numeric(16, 2),
        default=Decimal("0.00"),
        nullable=False,
    )
    budget_previously_used = db.Column(
        db.Numeric(16, 2),
        default=Decimal("0.00"),
        nullable=False,
    )
    note = db.Column(db.Text, default="", nullable=False)
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)

    company = db.relationship("Company")
    government_profile = db.relationship("GovernmentProfile")
    lines = db.relationship(
        "PurchaseLine",
        cascade="all, delete-orphan",
        back_populates="purchase",
        order_by="PurchaseLine.line_no",
    )

    @property
    def total_amount(self):
        return sum((line.amount for line in self.lines), Decimal("0.00"))

    @property
    def budget_this_time(self):
        return self.total_amount

    @property
    def budget_remaining(self):
        return (
            Decimal(self.budget_allocated or 0)
            - Decimal(self.budget_previously_used or 0)
            - Decimal(self.total_amount or 0)
        )


class PurchaseLine(db.Model):
    __tablename__ = "purchase_lines"

    id = db.Column(db.Integer, primary_key=True)
    purchase_id = db.Column(
        db.Integer,
        db.ForeignKey("purchases.id"),
        nullable=False,
        index=True,
    )
    line_no = db.Column(db.Integer, nullable=False)
    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False)
    description = db.Column(db.String(500), nullable=False)
    specification = db.Column(db.Text, default="", nullable=False)
    quantity = db.Column(db.Numeric(14, 2), nullable=False)
    unit_id = db.Column(db.Integer, db.ForeignKey("units.id"), nullable=False)
    unit_price = db.Column(db.Numeric(14, 2), nullable=False)
    amount = db.Column(db.Numeric(14, 2), nullable=False)

    purchase = db.relationship("Purchase", back_populates="lines")
    item = db.relationship("Item")
    unit = db.relationship("Unit")
