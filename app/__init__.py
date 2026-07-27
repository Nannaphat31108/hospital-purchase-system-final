import os
from flask import Flask
from sqlalchemy import inspect, text
from .models import DropdownOption, GovernmentProfile, db
from .routes import main_bp


def create_app():
    app = Flask(__name__, instance_relative_config=True)
    os.makedirs(app.instance_path, exist_ok=True)
    default_db = "sqlite:///" + os.path.join(app.instance_path, "hospital_purchase.db")
    database_url = os.getenv("DATABASE_URL", default_db).strip()
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)
    app.config.update(SECRET_KEY=os.getenv("SECRET_KEY", "change-this-secret-key"), SQLALCHEMY_DATABASE_URI=database_url, SQLALCHEMY_TRACK_MODIFICATIONS=False, SQLALCHEMY_ENGINE_OPTIONS={"pool_pre_ping": True, "pool_recycle": 280})
    db.init_app(app)
    app.register_blueprint(main_bp)
    with app.app_context():
        db.create_all()
        # Add new dropdown fields to an existing database without deleting data.
        inspector = inspect(db.engine)
        purchase_columns = {c["name"] for c in inspector.get_columns("purchases")}
        migrations = {
            "government_profile_id": "INTEGER",
            "procurement_type": "VARCHAR(255) DEFAULT 'เวชภัณฑ์มิใช่ยา'",
            "necessity_reason": "TEXT DEFAULT 'ใช้ในการรักษาผู้ป่วย'",
            "budget_allocated": "NUMERIC(16,2) DEFAULT 0 NOT NULL",
            "budget_previously_used": "NUMERIC(16,2) DEFAULT 0 NOT NULL",
        }
        for column, definition in migrations.items():
            if column not in purchase_columns:
                db.session.execute(text(f"ALTER TABLE purchases ADD COLUMN {column} {definition}"))
        company_columns = {c["name"] for c in inspector.get_columns("companies")}
        if "business_type" not in company_columns:
            db.session.execute(text("ALTER TABLE companies ADD COLUMN business_type VARCHAR(255) DEFAULT 'ขายส่ง,ขายปลีก,ให้บริการ'"))
        db.session.commit()
        if GovernmentProfile.query.count() == 0:
            db.session.add(GovernmentProfile())
        defaults = {
            "procurement_type": ["เวชภัณฑ์มิใช่ยา", "เวชภัณฑ์ยา", "วัสดุการแพทย์", "ครุภัณฑ์การแพทย์"],
            "necessity_reason": ["ใช้ในการรักษาผู้ป่วย"],
            "budget_source": ["เงินนอกงบประมาณจาก เงินบำรุงโรงพยาบาลสิงห์บุรี ปี ๒๕๖๙"],
            "delivery_place": ["โรงพยาบาลสิงห์บุรี ๙๑๗/๓"],
        }
        for category, values in defaults.items():
            for value in values:
                if not DropdownOption.query.filter_by(category=category, value=value).first():
                    db.session.add(DropdownOption(category=category, value=value))
        db.session.commit()
    return app
