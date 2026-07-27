import os

from flask import Flask
from sqlalchemy import inspect, text

from .models import DropdownOption, GovernmentProfile, db
from .routes import main_bp


def _add_missing_columns(table_name: str, definitions: dict[str, str]) -> None:
    """เพิ่มคอลัมน์ที่ขาด โดยไม่ลบข้อมูลเดิมในฐานข้อมูล"""
    inspector = inspect(db.engine)

    if table_name not in inspector.get_table_names():
        return

    existing_columns = {
        column["name"]
        for column in inspector.get_columns(table_name)
    }

    for column_name, column_definition in definitions.items():
        if column_name in existing_columns:
            continue

        db.session.execute(
            text(
                f'ALTER TABLE "{table_name}" '
                f'ADD COLUMN "{column_name}" {column_definition}'
            )
        )


def create_app():
    app = Flask(__name__, instance_relative_config=True)

    os.makedirs(app.instance_path, exist_ok=True)

    sqlite_path = os.path.join(
        app.instance_path,
        "hospital_purchase.db",
    )
    default_database_url = f"sqlite:///{sqlite_path}"

    database_url = os.getenv(
        "DATABASE_URL",
        default_database_url,
    ).strip()

    # รองรับ URL รูปแบบเก่าของ PostgreSQL
    if database_url.startswith("postgres://"):
        database_url = database_url.replace(
            "postgres://",
            "postgresql://",
            1,
        )

    app.config.update(
        SECRET_KEY=os.getenv(
            "SECRET_KEY",
            "change-this-secret-key",
        ),
        SQLALCHEMY_DATABASE_URI=database_url,
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        SQLALCHEMY_ENGINE_OPTIONS={
            "pool_pre_ping": True,
            "pool_recycle": 280,
        },
    )

    db.init_app(app)
    app.register_blueprint(main_bp)

    with app.app_context():
        # สร้างเฉพาะตารางที่ยังไม่มี
        db.create_all()

        # เพิ่มคอลัมน์ใหม่ในตารางบริษัท
        _add_missing_columns(
            "companies",
            {
                "business_type": (
                    "VARCHAR(255) "
                    "DEFAULT 'ขายส่ง,ขายปลีก,ให้บริการ'"
                ),
            },
        )

        # เพิ่มคอลัมน์ใหม่ในตารางใบสั่งซื้อ
        _add_missing_columns(
            "purchases",
            {
                "government_profile_id": "INTEGER",
                "procurement_type": (
                    "VARCHAR(255) "
                    "DEFAULT 'เวชภัณฑ์มิใช่ยา'"
                ),
                "necessity_reason": (
                    "TEXT DEFAULT 'ใช้ในการรักษาผู้ป่วย'"
                ),
                "budget_allocated": (
                    "NUMERIC(16,2) DEFAULT 0"
                ),
                "budget_previously_used": (
                    "NUMERIC(16,2) DEFAULT 0"
                ),
            },
        )

        # เพิ่มคอลัมน์ข้อมูลราชการทั้งหมดที่ฐานข้อมูลเดิมยังไม่มี
        _add_missing_columns(
            "government_profiles",
            {
                "name": (
                    "VARCHAR(255) "
                    "DEFAULT 'ข้อมูลโรงพยาบาลสิงห์บุรี'"
                ),
                "department": (
                    "TEXT DEFAULT "
                    "'โรงพยาบาลสิงห์บุรี "
                    "กลุ่มงานเภสัชกรรม "
                    "โทร. ๐ ๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙'"
                ),
                "letter_prefix": (
                    "VARCHAR(255) "
                    "DEFAULT 'สห ๐๐๓๓.๒๐๕.๑๒/'"
                ),
                "recipient": (
                    "VARCHAR(255) "
                    "DEFAULT 'ผู้ว่าราชการจังหวัดสิงห์บุรี'"
                ),
                "officer_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางพิณนภา ศริพันธุ์'"
                ),
                "officer_position": (
                    "VARCHAR(255) DEFAULT 'เจ้าหน้าที่'"
                ),
                "chief_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นายชัชวาลย์ บุญญฤทธิ์'"
                ),
                "chief_position": (
                    "VARCHAR(255) "
                    "DEFAULT 'เภสัชกรชำนาญการพิเศษ'"
                ),
                "approver_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นายพิรุณ ปิตะหงษ์นันท์'"
                ),
                "approver_position": (
                    "TEXT DEFAULT "
                    "'ผู้อำนวยการโรงพยาบาลสิงห์บุรี "
                    "ปฏิบัติราชการแทน "
                    "ผู้ว่าราชการจังหวัดสิงห์บุรี'"
                ),
                "inspector1_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางสาวกัญญพัชร ธนกิจการค้า'"
                ),
                "inspector1_position": (
                    "VARCHAR(255) DEFAULT 'เภสัชกร'"
                ),
                "inspector2_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางสาวชุลีพร สุขมี'"
                ),
                "inspector2_position": (
                    "VARCHAR(255) "
                    "DEFAULT 'เจ้าพนักงานเภสัชกรรมชำนาญงาน'"
                ),
                "inspector3_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางสาวกัญญาพัชร เลิศอนันตกูล'"
                ),
                "inspector3_position": (
                    "VARCHAR(255) "
                    "DEFAULT 'เจ้าพนักงานเภสัชกรรม'"
                ),
                "specifier_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางสาวนลินี เครือทิวา'"
                ),
                "specifier_position": (
                    "VARCHAR(255) "
                    "DEFAULT 'เภสัชกรชำนาญการ'"
                ),
                "receipt1_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางสาวธัญรดา ใจเสน'"
                ),
                "receipt2_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางสาวจุฑามาศ อาคมสรรเสริญ'"
                ),
                "receipt3_name": (
                    "VARCHAR(255) "
                    "DEFAULT 'นางสาวนริศรา ม่วงงาม'"
                ),
                "active": "BOOLEAN DEFAULT TRUE",
                "created_at": (
                    "TIMESTAMP DEFAULT CURRENT_TIMESTAMP"
                ),
                "updated_at": (
                    "TIMESTAMP DEFAULT CURRENT_TIMESTAMP"
                ),
            },
        )

        db.session.commit()

        # สร้างข้อมูลราชการเริ่มต้น หากยังไม่มีข้อมูล
        if GovernmentProfile.query.count() == 0:
            db.session.add(GovernmentProfile())

        default_dropdowns = {
            "procurement_type": [
                "เวชภัณฑ์มิใช่ยา",
                "เวชภัณฑ์ยา",
                "วัสดุการแพทย์",
                "ครุภัณฑ์การแพทย์",
            ],
            "necessity_reason": [
                "ใช้ในการรักษาผู้ป่วย",
            ],
            "budget_source": [
                (
                    "เงินนอกงบประมาณจาก "
                    "เงินบำรุงโรงพยาบาลสิงห์บุรี ปี ๒๕๖๙"
                ),
            ],
            "delivery_place": [
                "โรงพยาบาลสิงห์บุรี ๙๑๗/๓",
            ],
        }

        for category, values in default_dropdowns.items():
            for value in values:
                exists = DropdownOption.query.filter_by(
                    category=category,
                    value=value,
                ).first()

                if not exists:
                    db.session.add(
                        DropdownOption(
                            category=category,
                            value=value,
                        )
                    )

        db.session.commit()

    return app