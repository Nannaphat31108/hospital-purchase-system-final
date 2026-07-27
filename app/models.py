from datetime import datetime
from decimal import Decimal
from flask_sqlalchemy import SQLAlchemy


db = SQLAlchemy()


class TimestampMixin:
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)


class Company(TimestampMixin, db.Model):
    __tablename__ = "companies"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False, index=True)
    address = db.Column(db.Text, default="")
    phone = db.Column(db.String(100), default="")
    tax_id = db.Column(db.String(50), default="")
    bank_name = db.Column(db.String(255), default="")
    bank_branch = db.Column(db.String(255), default="")
    account_no = db.Column(db.String(100), default="")
    account_name = db.Column(db.String(255), default="")
    business_type = db.Column(db.String(255), default="ขายส่ง,ขายปลีก,ให้บริการ")
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)


class Unit(TimestampMixin, db.Model):
    __tablename__ = "units"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True, index=True)
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)


class Item(TimestampMixin, db.Model):
    __tablename__ = "items"
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(100), default="", index=True)
    name = db.Column(db.String(500), nullable=False, index=True)
    specification = db.Column(db.Text, default="")
    default_price = db.Column(db.Numeric(14, 2), default=Decimal("0.00"), nullable=False)
    unit_id = db.Column(db.Integer, db.ForeignKey("units.id"), nullable=False)
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)
    unit = db.relationship("Unit")


class GovernmentProfile(TimestampMixin, db.Model):
    __tablename__ = "government_profiles"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False, default="ข้อมูลโรงพยาบาลสิงห์บุรี")
    department = db.Column(db.Text, nullable=False, default="โรงพยาบาลสิงห์บุรี กลุ่มงานเภสัชกรรม โทร. ๐ ๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙")
    letter_prefix = db.Column(db.String(255), default="สห ๐๐๓๓.๒๐๕.๑๒/")
    recipient = db.Column(db.String(255), default="ผู้ว่าราชการจังหวัดสิงห์บุรี")
    officer_name = db.Column(db.String(255), default="นางพิณนภา ศริพันธุ์")
    officer_position = db.Column(db.String(255), default="เจ้าหน้าที่")
    chief_name = db.Column(db.String(255), default="นายชัชวาลย์ บุญญฤทธิ์")
    chief_position = db.Column(db.String(255), default="เภสัชกรชำนาญการพิเศษ")
    approver_name = db.Column(db.String(255), default="นายพิรุณ ปิตะหงษ์นันท์")
    approver_position = db.Column(db.Text, default="ผู้อำนวยการโรงพยาบาลสิงห์บุรี ปฏิบัติราชการแทน ผู้ว่าราชการจังหวัดสิงห์บุรี")
    inspector1_name = db.Column(db.String(255), default="นางสาวกัญญพัชร ธนกิจการค้า")
    inspector1_position = db.Column(db.String(255), default="เภสัชกร")
    inspector2_name = db.Column(db.String(255), default="นางสาวชุลีพร สุขมี")
    inspector2_position = db.Column(db.String(255), default="เจ้าพนักงานเภสัชกรรมชำนาญงาน")
    inspector3_name = db.Column(db.String(255), default="นางสาวกัญญาพัชร เลิศอนันตกูล")
    inspector3_position = db.Column(db.String(255), default="เจ้าพนักงานเภสัชกรรม")
    specifier_name = db.Column(db.String(255), default="นางสาวนลินี เครือทิวา")
    specifier_position = db.Column(db.String(255), default="เภสัชกรชำนาญการ")
    receipt1_name = db.Column(db.String(255), default="นางสาวธัญรดา ใจเสน")
    receipt2_name = db.Column(db.String(255), default="นางสาวจุฑามาศ อาคมสรรเสริญ")
    receipt3_name = db.Column(db.String(255), default="นางสาวนริศรา ม่วงงาม")
    active = db.Column(db.Boolean, default=True, nullable=False)


class DropdownOption(TimestampMixin, db.Model):
    __tablename__ = "dropdown_options"
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(100), nullable=False, index=True)
    value = db.Column(db.Text, nullable=False)
    active = db.Column(db.Boolean, default=True, nullable=False)


class Purchase(TimestampMixin, db.Model):
    __tablename__ = "purchases"
    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(100), nullable=False, unique=True, index=True)
    document_date = db.Column(db.Date, nullable=False, index=True)
    company_id = db.Column(db.Integer, db.ForeignKey("companies.id"), nullable=False)
    government_profile_id = db.Column(db.Integer, db.ForeignKey("government_profiles.id"), nullable=True)
    procurement_type = db.Column(db.String(255), default="เวชภัณฑ์มิใช่ยา")
    necessity_reason = db.Column(db.Text, default="ใช้ในการรักษาผู้ป่วย")
    project_number = db.Column(db.String(100), default="")
    contract_control_number = db.Column(db.String(100), default="")
    delivery_days = db.Column(db.Integer, default=30)
    delivery_place = db.Column(db.String(500), default="โรงพยาบาลสิงห์บุรี ๙๑๗/๓")
    budget_source = db.Column(db.String(500), default="เงินนอกงบประมาณจาก เงินบำรุงโรงพยาบาลสิงห์บุรี ปี ๒๕๖๙")
    budget_allocated = db.Column(db.Numeric(16, 2), default=Decimal("0.00"), nullable=False)
    budget_previously_used = db.Column(db.Numeric(16, 2), default=Decimal("0.00"), nullable=False)
    note = db.Column(db.Text, default="")
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)

    company = db.relationship("Company")
    government_profile = db.relationship("GovernmentProfile")
    lines = db.relationship("PurchaseLine", cascade="all, delete-orphan", back_populates="purchase", order_by="PurchaseLine.line_no")

    @property
    def total_amount(self):
        return sum((line.amount for line in self.lines), Decimal("0.00"))

    @property
    def budget_remaining(self):
        return Decimal(self.budget_allocated or 0) - Decimal(self.budget_previously_used or 0) - self.total_amount

    @property
    def procurement_label(self):
        if len(self.lines) == 1:
            line = self.lines[0]
            return f"{line.description} จำนวน {line.quantity:g} {line.unit.name}"
        return f"{self.procurement_type} จำนวน {len(self.lines)} รายการ"


class PurchaseLine(db.Model):
    __tablename__ = "purchase_lines"
    id = db.Column(db.Integer, primary_key=True)
    purchase_id = db.Column(db.Integer, db.ForeignKey("purchases.id"), nullable=False, index=True)
    line_no = db.Column(db.Integer, nullable=False)
    item_id = db.Column(db.Integer, db.ForeignKey("items.id"), nullable=False)
    description = db.Column(db.String(500), nullable=False)
    specification = db.Column(db.Text, default="")
    quantity = db.Column(db.Numeric(14, 2), nullable=False)
    unit_id = db.Column(db.Integer, db.ForeignKey("units.id"), nullable=False)
    unit_price = db.Column(db.Numeric(14, 2), nullable=False)
    amount = db.Column(db.Numeric(14, 2), nullable=False)
    purchase = db.relationship("Purchase", back_populates="lines")
    item = db.relationship("Item")
    unit = db.relationship("Unit")
