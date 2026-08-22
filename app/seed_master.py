from decimal import Decimal

from .master_data import COMPANIES, ITEM_NAMES
from .models import Company, Item, Unit, db


def seed_user_master_data():
    unit = Unit.query.filter_by(name='ชิ้น').first()
    if unit is None:
        unit = Unit(name='ชิ้น', active=True)
        db.session.add(unit)
        db.session.flush()

    existing_companies = {c.name.strip(): c for c in Company.query.all()}
    for name, address, phone, tax_id, bank_name, bank_branch, account_no, account_name in COMPANIES:
        company = existing_companies.get(name)
        if company is None:
            company = Company(name=name)
            db.session.add(company)
            existing_companies[name] = company
        company.address = address
        company.phone = phone
        company.tax_id = tax_id
        company.bank_name = bank_name
        company.bank_branch = bank_branch
        company.account_no = account_no
        company.account_name = account_name
        company.active = True

    sample_company = Company.query.filter_by(name='บริษัทตัวอย่าง จำกัด').first()
    if sample_company:
        sample_company.active = False

    existing_items = {i.name.strip(): i for i in Item.query.all()}
    for idx, name in enumerate(ITEM_NAMES, start=1):
        item = existing_items.get(name)
        if item is None:
            item = Item(name=name, unit_id=unit.id)
            db.session.add(item)
            existing_items[name] = item
        if not item.code:
            item.code = f'MED-{idx:04d}'
        item.unit_id = item.unit_id or unit.id
        if item.default_price is None:
            item.default_price = Decimal('0.00')
        item.active = True

    sample_item = Item.query.filter_by(name='เวชภัณฑ์ตัวอย่าง').first()
    if sample_item:
        sample_item.active = False

    db.session.commit()
