import os

from flask import Flask
from sqlalchemy import inspect, text

from .models import DropdownOption, GovernmentProfile, db
from .routes import main_bp


def _quote_identifier(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def _add_missing_columns(
    table_name: str,
    definitions: dict[str, str],
) -> None:
    inspector = inspect(db.engine)

    if table_name not in inspector.get_table_names():
        return

    existing_columns = {
        column["name"]
        for column in inspector.get_columns(table_name)
    }

    quoted_table = _quote_identifier(table_name)

    for column_name, column_definition in definitions.items():
        if column_name in existing_columns:
            continue

        quoted_column = _quote_identifier(column_name)
        db.session.execute(
            text(
                f"ALTER TABLE {quoted_table} "
                f"ADD COLUMN {quoted_column} {column_definition}"
            )
        )


def _migrate_existing_database() -> None:
    _add_missing_columns(
        "companies",
        {
            "business_type": (
                "VARCHAR(255) DEFAULT 'ขายส่ง,ขายปลีก,ให้บริการ'"
            ),
        },
    )

    _add_missing_columns(
        "purchases",
        {
            "government_profile_id": "INTEGER",
            "procurement_type": (
                "VARCHAR(255) DEFAULT 'เวชภัณฑ์มิใช่ยา'"
            ),
            "necessity_reason": (
                "TEXT DEFAULT 'ใช้ในการรักษาผู้ป่วย'"
            ),
            "budget_allocated": "NUMERIC(16,2) DEFAULT 0",
            "budget_previously_used": "NUMERIC(16,2) DEFAULT 0",
        },
    )

    _add_missing_columns(
        "government_profiles",
        {
            "name": (
                "VARCHAR(255) DEFAULT 'ข้อมูลโรงพยาบาลสิงห์บุรี'"
            ),
            "department": (
                "TEXT DEFAULT 'โรงพยาบาลสิงห์บุรี "
                "กลุ่มงานเภสัชกรรม "
                "โทร. ๐ ๓๖๕๒ ๒๕๐๘ ต่อ ๑๑๒๙'"
            ),
            "letter_prefix": (
                "VARCHAR(255) DEFAULT 'สห ๐๐๓๓.๒๐๕.๑๒/'"
            ),
            "recipient": (
                "VARCHAR(255) DEFAULT 'ผู้ว่าราชการจังหวัดสิงห์บุรี'"
            ),
            "officer_name": (
                "VARCHAR(255) DEFAULT 'นางพิณนภา ศริพันธุ์'"
            ),
            "officer_position": (
                "VARCHAR(255) DEFAULT 'เจ้าหน้าที่'"
            ),
            "chief_name": (
                "VARCHAR(255) DEFAULT 'นายชัชวาลย์ บุญญฤทธิ์'"
            ),
            "chief_position": (
                "VARCHAR(255) DEFAULT 'เภสัชกรชำนาญการพิเศษ'"
            ),
            "approver_name": (
                "VARCHAR(255) DEFAULT 'นายพิรุณ ปิตะหงษ์นันท์'"
            ),
            "approver_position": (
                "TEXT DEFAULT 'ผู้อำนวยการโรงพยาบาลสิงห์บุรี "
                "ปฏิบัติราชการแทน ผู้ว่าราชการจังหวัดสิงห์บุรี'"
            ),
            "inspector1_name": (
                "VARCHAR(255) DEFAULT 'นางสาวกัญญพัชร ธนกิจการค้า'"
            ),
            "inspector1_position": (
                "VARCHAR(255) DEFAULT 'เภสัชกร'"
            ),
            "inspector2_name": (
                "VARCHAR(255) DEFAULT 'นางสาวชุลีพร สุขมี'"
            ),
            "inspector2_position": (
                "VARCHAR(255) DEFAULT 'เจ้าพนักงานเภสัชกรรมชำนาญงาน'"
            ),
            "inspector3_name": (
                "VARCHAR(255) DEFAULT 'นางสาวกัญญาพัชร เลิศอนันตกูล'"
            ),
            "inspector3_position": (
                "VARCHAR(255) DEFAULT 'เจ้าพนักงานเภสัชกรรม'"
            ),
            "specifier_name": (
                "VARCHAR(255) DEFAULT 'นางสาวนลินี เครือทิวา'"
            ),
            "specifier_position": (
                "VARCHAR(255) DEFAULT 'เภสัชกรชำนาญการ'"
            ),
            "receipt1_name": (
                "VARCHAR(255) DEFAULT 'นางสาวธัญรดา ใจเสน'"
            ),
            "receipt2_name": (
                "VARCHAR(255) DEFAULT 'นางสาวจุฑามาศ อาคมสรรเสริญ'"
            ),
            "receipt3_name": (
                "VARCHAR(255) DEFAULT 'นางสาวนริศรา ม่วงงาม'"
            ),
            "active": "BOOLEAN DEFAULT TRUE",
            "created_at": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
            "updated_at": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        },
    )

    _add_missing_columns(
        "dropdown_options",
        {
            "label": "VARCHAR(500) DEFAULT ''",
            "value": "VARCHAR(500) DEFAULT ''",
            "active": "BOOLEAN DEFAULT TRUE",
            "created_at": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
            "updated_at": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        },
    )

    db.session.commit()

    db.session.execute(
        text(
            """
            UPDATE dropdown_options
               SET label = COALESCE(NULLIF(label, ''), value, '')
             WHERE label IS NULL OR label = ''
            """
        )
    )
    db.session.execute(
        text(
            """
            UPDATE dropdown_options
               SET value = COALESCE(NULLIF(value, ''), label, '')
             WHERE value IS NULL OR value = ''
            """
        )
    )
    db.session.commit()


def _create_default_master_data() -> None:
    if GovernmentProfile.query.count() == 0:
        db.session.add(
            GovernmentProfile(name="ข้อมูลโรงพยาบาลสิงห์บุรี")
        )
        db.session.commit()

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

    with db.session.no_autoflush:
        existing_pairs = {
            (option.category, option.value)
            for option in DropdownOption.query.all()
        }

    for category, values in default_dropdowns.items():
        for value in values:
            if (category, value) in existing_pairs:
                continue

            db.session.add(
                DropdownOption(
                    category=category,
                    label=value,
                    value=value,
                    active=True,
                )
            )
            existing_pairs.add((category, value))

    db.session.commit()


def create_app():
    app = Flask(__name__, instance_relative_config=True)
    os.makedirs(app.instance_path, exist_ok=True)

    default_db = "sqlite:///" + os.path.join(
        app.instance_path,
        "hospital_purchase.db",
    )
    database_url = os.getenv("DATABASE_URL", default_db).strip()

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
        db.create_all()
        _migrate_existing_database()
        _create_default_master_data()

    return app
