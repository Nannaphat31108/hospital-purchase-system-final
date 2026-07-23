import os
from flask import Flask
from .models import db
from .routes import main_bp


def create_app():
    app = Flask(__name__, instance_relative_config=True)
    os.makedirs(app.instance_path, exist_ok=True)

    # Local development uses SQLite. Render uses the persistent PostgreSQL
    # connection supplied through DATABASE_URL by render.yaml.
    default_db = "sqlite:///" + os.path.join(app.instance_path, "hospital_purchase.db")
    database_url = os.getenv("DATABASE_URL", default_db).strip()
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)

    app.config.update(
        SECRET_KEY=os.getenv("SECRET_KEY", "change-this-secret-key"),
        SQLALCHEMY_DATABASE_URI=database_url,
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        SQLALCHEMY_ENGINE_OPTIONS={
            "pool_pre_ping": True,
            "pool_recycle": 280,
        },
    )

    db.init_app(app)
    app.register_blueprint(main_bp)

    # Creates missing tables only. It never drops or clears existing data.
    with app.app_context():
        db.create_all()

    return app
