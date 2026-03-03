from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

def init_db(app):
    app.config.setdefault('SQLALCHEMY_DATABASE_URI', 'sqlite:///app.db')
    app.config.setdefault('SQLALCHEMY_TRACK_MODIFICATIONS', False)
    db.init_app(app)
    with app.app_context():
        from models import db_models  # noqa: F401 — ensure models are registered
        db.create_all()
        print("資料庫初始化完成")
