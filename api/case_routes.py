from flask import Blueprint, request, jsonify, session
from models.database import db
from models.db_models import Case, ChatMessage

case_bp = Blueprint('case_bp', __name__)


def _current_user_id():
    return session.get('user_id')


def _require_login():
    if not _current_user_id():
        return jsonify({'status': 'error', 'error': '請先登入'}), 401
    return None


# ── 案場 CRUD ──────────────────────────────────────────────

@case_bp.route('/cases', methods=['GET'])
def get_cases():
    err = _require_login()
    if err: return err

    cases = (Case.query
             .filter_by(user_id=_current_user_id())
             .order_by(Case.created_at)
             .all())
    return jsonify({'status': 'success', 'cases': [
        {
            'id':             c.id,
            'name':           c.name,
            'site_type':      c.site_type,
            'excel_filename': c.excel_filename,
        } for c in cases
    ]})


@case_bp.route('/cases', methods=['POST'])
def create_case():
    err = _require_login()
    if err: return err

    data = request.json or {}
    case = Case(
        user_id        = _current_user_id(),
        name           = data.get('name', '新試算表'),
        site_type      = data.get('site_type', 'single'),
        excel_filename = data.get('excel_filename'),
    )
    db.session.add(case)
    db.session.commit()
    return jsonify({'status': 'success', 'case': {
        'id':             case.id,
        'name':           case.name,
        'site_type':      case.site_type,
        'excel_filename': case.excel_filename,
    }})


@case_bp.route('/cases/<int:case_id>', methods=['PUT'])
def update_case(case_id):
    err = _require_login()
    if err: return err

    case = Case.query.filter_by(id=case_id, user_id=_current_user_id()).first()
    if not case:
        return jsonify({'status': 'error', 'error': '案場不存在'}), 404

    data = request.json or {}
    if 'name'           in data: case.name           = data['name']
    if 'excel_filename' in data: case.excel_filename = data['excel_filename']
    if 'site_type'      in data: case.site_type      = data['site_type']
    db.session.commit()
    return jsonify({'status': 'success'})


@case_bp.route('/cases/<int:case_id>', methods=['DELETE'])
def delete_case(case_id):
    err = _require_login()
    if err: return err

    case = Case.query.filter_by(id=case_id, user_id=_current_user_id()).first()
    if not case:
        return jsonify({'status': 'error', 'error': '案場不存在'}), 404

    db.session.delete(case)
    db.session.commit()
    return jsonify({'status': 'success'})


# ── 聊天紀錄 ───────────────────────────────────────────────

@case_bp.route('/cases/<int:case_id>/messages', methods=['GET'])
def get_messages(case_id):
    err = _require_login()
    if err: return err

    case = Case.query.filter_by(id=case_id, user_id=_current_user_id()).first()
    if not case:
        return jsonify({'status': 'error', 'error': '案場不存在'}), 404

    msgs = (ChatMessage.query
            .filter_by(case_id=case_id)
            .order_by(ChatMessage.created_at)
            .all())
    return jsonify({'status': 'success', 'messages': [
        {'role': m.role, 'content': m.content} for m in msgs
    ]})


@case_bp.route('/cases/<int:case_id>/messages', methods=['POST'])
def save_messages(case_id):
    err = _require_login()
    if err: return err

    case = Case.query.filter_by(id=case_id, user_id=_current_user_id()).first()
    if not case:
        return jsonify({'status': 'error', 'error': '案場不存在'}), 404

    for msg in (request.json or {}).get('messages', []):
        db.session.add(ChatMessage(
            case_id = case_id,
            role    = msg['role'],
            content = msg['content'],
        ))
    db.session.commit()
    return jsonify({'status': 'success'})
