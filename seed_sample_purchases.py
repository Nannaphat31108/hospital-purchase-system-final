from datetime import date
from decimal import Decimal

from app import create_app
from app.models import Company, Item, Purchase, PurchaseLine, Unit, db

app = create_app()

with app.app_context():
    company = Company.query.first()
    unit = Unit.query.first()
    item = Item.query.first()

    samples = [
        {
            "po_number": "PO-SAMPLE-001",
            "lines": [
                {"description": "เวชภัณฑ์ตัวอย่าง A", "quantity": 10, "price": 100},
            ],
        },
        {
            "po_number": "PO-SAMPLE-002",
            "lines": [
                {"description": "เวชภัณฑ์ตัวอย่าง A", "quantity": 5, "price": 100},
                {"description": "เวชภัณฑ์ตัวอย่าง B", "quantity": 20, "price": 50},
            ],
        },
        {
            "po_number": "PO-SAMPLE-003",
            "lines": [
                {"description": "เวชภัณฑ์ตัวอย่าง A", "quantity": 3, "price": 100},
                {"description": "เวชภัณฑ์ตัวอย่าง B", "quantity": 8, "price": 50},
                {"description": "เวชภัณฑ์ตัวอย่าง C", "quantity": 15, "price": 30},
            ],
        },
    ]

    for sample in samples:
        if Purchase.query.filter_by(po_number=sample["po_number"]).first():
            print(f"ข้าม {sample['po_number']} (มีอยู่แล้ว)")
            continue

        purchase = Purchase(
            po_number=sample["po_number"],
            document_date=date.today(),
            company_id=company.id,
            budget_allocated=Decimal("1000000.00"),
        )
        db.session.add(purchase)
        db.session.flush()

        for i, line in enumerate(sample["lines"], start=1):
            qty = Decimal(str(line["quantity"]))
            price = Decimal(str(line["price"]))
            db.session.add(
                PurchaseLine(
                    purchase_id=purchase.id,
                    line_no=i,
                    item_id=item.id,
                    description=line["description"],
                    specification=item.specification,
                    quantity=qty,
                    unit_id=unit.id,
                    unit_price=price,
                    amount=qty * price,
                )
            )

        db.session.commit()
        print(f"สร้าง {sample['po_number']} เรียบร้อย ({len(sample['lines'])} รายการ)")
