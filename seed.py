from app import create_app
from app.models import Company, Item, Unit, db

app = create_app()
with app.app_context():
    if Unit.query.count() == 0:
        units = [Unit(name="SET"), Unit(name="ชิ้น"), Unit(name="กล่อง"), Unit(name="ม้วน")]
        db.session.add_all(units)
        db.session.flush()
        db.session.add(Company(name="บริษัทตัวอย่าง จำกัด", address="กรุงเทพมหานคร", phone="02-000-0000", tax_id="0100000000000"))
        db.session.add(Item(code="MED-001", name="เวชภัณฑ์ตัวอย่าง", default_price=100, unit_id=units[0].id, specification="คุณลักษณะเฉพาะตัวอย่าง สามารถแก้ไขได้"))
        db.session.commit()
        print("เพิ่มข้อมูลตัวอย่างเรียบร้อย")
    else:
        print("ฐานข้อมูลมีข้อมูลอยู่แล้ว")
