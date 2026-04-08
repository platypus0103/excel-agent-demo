from flask import Blueprint, request, jsonify, session
from models.database import db
from models.db_models import User
from utils.app_logger import log_action

auth_bp = Blueprint('auth_bp', __name__)


@auth_bp.route('/auth/login', methods=['POST'])
def login():
    data  = request.json or {}
    email = data.get('email', '').strip().lower()

    if not email or '@' not in email:
        return jsonify({'status': 'error', 'error': '請輸入有效的 Email'}), 400

    user = User.query.filter_by(email=email).first()
    if not user:
        user = User(email=email)
        db.session.add(user)
        db.session.commit()
        print(f"[Auth] 新使用者建立: {email}")
        log_action(email, 'register', '新帳號建立')
    else:
        print(f"[Auth] 使用者登入: {email}")
        log_action(email, 'login', '登入成功')

    session['user_id']    = user.id
    session['user_email'] = user.email
    return jsonify({'status': 'success', 'email': user.email, 'user_id': user.id})


@auth_bp.route('/auth/logout', methods=['POST'])
def logout():
    email = session.get('user_email', 'anonymous')
    log_action(email, 'logout', '登出')
    session.clear()
    return jsonify({'status': 'success'})


@auth_bp.route('/auth/me', methods=['GET'])
def me():
    if 'user_id' not in session:
        return jsonify({'status': 'not_logged_in'}), 401
    return jsonify({
        'status': 'ok',
        'email':   session['user_email'],
        'user_id': session['user_id']
    })
