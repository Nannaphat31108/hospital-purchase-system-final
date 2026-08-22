# Hospital Purchase System

ระบบจัดซื้อโรงพยาบาลด้วย Flask รองรับบริษัท หน่วย เวชภัณฑ์ ใบสั่งซื้อ การค้นหา แก้ไขย้อนหลัง และส่งออก Word ตามแบบราชการ

## ข้อมูลไม่หายบน Render
- ใช้ PostgreSQL ผ่าน `DATABASE_URL` ที่ Render Blueprint สร้างให้จาก `render.yaml`
- `db.create_all()` สร้างเฉพาะตารางที่ยังไม่มี และไม่มี `db.drop_all()` ในการเริ่มแอป
- บริษัท หน่วย เวชภัณฑ์ ใบสั่งซื้อ และรายการย่อยจึงไม่ถูกล้างเมื่อ Deploy ใหม่ตามปกติ
- ในเครื่องที่ไม่มี `DATABASE_URL` จะใช้ SQLite ใน `instance/hospital_purchase.db` เท่านั้น

## ฟอร์ม Word 8 แบบ
เอกสาร Word ทุกแบบตั้งฟอนต์ **TH Sarabun New ขนาด 16 pt**

1. ใบสั่งซื้อ 1 รายการ
2. ใบสั่งซื้อหลายรายการ
3. แบบกำหนดรายละเอียดคุณลักษณะเฉพาะ (Spec)
4. ใบตรวจรับ 1 รายการ
5. ใบตรวจรับ 2 รายการ
6. ใบตรวจรับมากกว่า 2 รายการ
7. แบบแสดงความบริสุทธิ์ใจ
8. ชุดรายงานจัดซื้อ: รายงานขอซื้อ รายงานผลการพิจารณา และประกาศผู้ชนะ

มีปุ่ม **Word ทุกฟอร์มรวมไฟล์เดียว** เพิ่มเติม

## รันในเครื่อง
```powershell
pip install -r requirements.txt
python run.py
```
เปิด `http://127.0.0.1:5000`

## Deploy บน Render
1. Push โค้ดขึ้น GitHub
2. ใน Render เลือก **New > Blueprint**
3. เลือก repository นี้
4. Render จะสร้าง Web Service และ PostgreSQL พร้อมกำหนด `DATABASE_URL` อัตโนมัติ

หลัง Deploy ให้ตรวจใน Web Service > Environment ว่ามี `DATABASE_URL` และค่าเชื่อมต่อเป็น PostgreSQL
