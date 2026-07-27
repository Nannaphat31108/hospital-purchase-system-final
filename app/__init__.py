import os

from flask import Flask
from sqlalchemy import inspect, text

from .models import DropdownOption, GovernmentProfile, db
from .routes import main_bp


def create_app():
    app = Flask(__name__, instance_relative_config=True)

    # สร้างโฟลเดอร์ instance สำหรับ SQLite ตอนรันในเครื่อง
    os.makedirs(app.instance_path, exist_ok=True)

    default_db = "sqlite:///" + os.path.join(
        app.instance_path,
        "hospital_purchase.db",
    )

    database_url = os.getenv("DATABASE_URL", default_db).strip()

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
        # คำสั่งนี้ไม่ลบข้อมูลเดิม
        db.create_all()

        try:
            inspector = inspect(db.engine)
            table_names = set(inspector.get_table_names())

            # ---------------------------------------------------------
            # Migration: government_profiles
            # ต้องเพิ่ม name ก่อน GovernmentProfile.query.count()
            # ---------------------------------------------------------
            if "government_profiles" in table_names:
                government_columns = {
                    column["name"]
                    for column in inspector.get_columns(
                        "government_profiles"
                    )
                }

                if "name" not in government_columns:
                    db.session.execute(
                        text(
                            """
                            ALTER TABLE government_profiles
                            ADD COLUMN name VARCHAR(255)
                            """
                        )
                    )

                    db.session.execute(
                        text(
                            """
                            UPDATE government_profiles
                            SET name = 'ข้อมูลราชการหลัก'
                            WHERE name IS NULL OR name = ''
                            """
                        )
                    )

            # ---------------------------------------------------------
            # Migration: purchases
            # ---------------------------------------------------------
            if "purchases" in table_names:
                purchase_columns = {
                    column["name"]
                    for column in inspector.get_columns("purchases")
                }

                purchase_migrations = {
                    "government_profile_id": "INTEGER",
                    "procurement_type": (
                        "VARCHAR(255) DEFAULT 'เวชภัณฑ์มิใช่ยา'"
                    ),
                    "necessity_reason": (
                        "TEXT DEFAULT 'ใช้ในการรักษาผู้ป่วย'"
                    ),
                    "budget_allocated": (
                        "NUMERIC(16,2) DEFAULT 0 NOT NULL"
                    ),
                    "budget_previously_used": (
                        "NUMERIC(16,2) DEFAULT 0 NOT NULL"
                    ),
                }

                for column_name, column_definition in (
                    purchase_migrations.items()
                ):
                    if column_name not in purchase_columns:
                        db.session.execute(
                            text(
                                f"""
                                ALTER TABLE purchases
                                ADD COLUMN {column_name}
                                {column_definition}
                                """
                            )
                        )

            # ---------------------------------------------------------
            # Migration: companies
            # ---------------------------------------------------------
            if "companies" in table_names:
                company_columns = {
                    column["name"]
                    for column in inspector.get_columns("companies")
                }

                if "business_type" not in company_columns:
                    db.session.execute(
                        text(
                            """
                            ALTER TABLE companies
                            ADD COLUMN business_type VARCHAR(255)
                            DEFAULT 'ขายส่ง,ขายปลีก,ให้บริการ'
                            """
                        )
                    )

            db.session.commit()

        except Exception:
            db.session.rollback()
            app.logger.exception(
                "Database migration failed during application startup"
            )
            raise

        # -------------------------------------------------------------
        # เพิ่มข้อมูลราชการเริ่มต้น
        # -------------------------------------------------------------
        if GovernmentProfile.query.count() == 0:
            profile = GovernmentProfile()

            # ตั้งชื่อเฉพาะเมื่อ Model มี attribute name
            if hasattr(profile, "name"):
                profile.name = "ข้อมูลราชการหลัก"

            db.session.add(profile)
            db.session.commit()

        # -------------------------------------------------------------
        # เพิ่มตัวเลือก Dropdown เริ่มต้น
        # -------------------------------------------------------------
        defaults = {
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

        for category, values in defaults.items():
            for value in values:
                existing_option = DropdownOption.query.filter_by(
                    category=category,
                    value=value,
                ).first()

                if existing_option is None:
                    db.session.add(
                        DropdownOption(
                            category=category,
                            value=value,
                        )
                    )

        db.session.commit()

    return app