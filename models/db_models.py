from datetime import datetime
from .database import db


class User(db.Model):
    __tablename__ = 'users'

    id         = db.Column(db.Integer, primary_key=True)
    email      = db.Column(db.String(255), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    cases = db.relationship('Case', backref='user', lazy=True,
                            cascade='all, delete-orphan')


class Case(db.Model):
    __tablename__ = 'cases'

    id               = db.Column(db.Integer, primary_key=True)
    user_id          = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    name             = db.Column(db.String(255), nullable=False)
    site_type        = db.Column(db.String(50), default='single')
    excel_filename   = db.Column(db.String(500))   # 原始檔名，例如 公版.xlsx
    created_at       = db.Column(db.DateTime, default=datetime.utcnow)

    messages = db.relationship('ChatMessage', backref='case', lazy=True,
                               cascade='all, delete-orphan')


class ChatMessage(db.Model):
    __tablename__ = 'chat_messages'

    id         = db.Column(db.Integer, primary_key=True)
    case_id    = db.Column(db.Integer, db.ForeignKey('cases.id'), nullable=False)
    role       = db.Column(db.String(20), nullable=False)   # 'user' 或 'bot'
    content    = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
