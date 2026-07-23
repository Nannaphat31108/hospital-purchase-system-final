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


class Purchase(TimestampMixin, db.Model):
    __tablename__ = "purchases"
    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(100), nullable=False, unique=True, index=True)
    document_date = db.Column(db.Date, nullable=False, index=True)
    company_id = db.Column(db.Integer, db.ForeignKey("companies.id"), nullable=False)
    project_number = db.Column(db.String(100), default="")
    contract_control_number = db.Column(db.String(100), default="")
    delivery_days = db.Column(db.Integer, default=30)
    delivery_place = db.Column(db.String(500), default="โรงพยาบาลสิงห์บุรี")
    budget_source = db.Column(db.String(500), default="เงินบำรุงโรงพยาบาลสิงห์บุรี")
    note = db.Column(db.Text, default="")
    active = db.Column(db.Boolean, default=True, nullable=False, index=True)

    company = db.relationship("Company")
    lines = db.relationship(
        "PurchaseLine",
        cascade="all, delete-orphan",
        back_populates="purchase",
        order_by="PurchaseLine.line_no",
    )

    @property
    def total_amount(self):
        return sum((line.amount for line in self.lines), Decimal("0.00"))


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
