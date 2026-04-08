"""
Flask 主應用程式
AI Agent Web 介面
"""
import os
import sys
import time
from flask import Flask, render_template, jsonify, request, session, g
from flask_cors import CORS
from utils.app_logger import get_logger, log_action, log_error

# 將當前目錄添加到 Python 路徑
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

def create_app():
    """工廠函數創建 Flask 應用"""
    app = Flask(__name__,
                static_folder='static',
                static_url_path='/static',
                template_folder='templates')

    # 配置 CORS
    CORS(app, resources={
        r"/api/*": {
            "origins": "*",
            "methods": ["GET", "POST", "PUT", "DELETE"],
            "allow_headers": ["Content-Type", "Authorization"]
        }
    })

    # 基本配置
    app.config['DEBUG'] = True
    app.config['SECRET_KEY'] = 'ai-agent-secret-key-2025'
    app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

    # 初始化資料庫
    from models.database import init_db
    init_db(app)

    # 註冊 AI Agent API 藍圖
    try:
        from api.agent_routes import agent_bp
        app.register_blueprint(agent_bp, url_prefix='/api')
        print("AI Agent API 路由註冊成功")
    except ImportError as e:
        print(f"警告: 無法匯入 api.agent_routes: {e}")

    # 註冊 Auth 藍圖
    try:
        from api.auth_routes import auth_bp
        app.register_blueprint(auth_bp, url_prefix='/api')
        print("Auth 路由註冊成功")
    except ImportError as e:
        print(f"警告: 無法匯入 api.auth_routes: {e}")

    # 註冊 Case 藍圖
    try:
        from api.case_routes import case_bp
        app.register_blueprint(case_bp, url_prefix='/api')
        print("Case 路由註冊成功")
    except ImportError as e:
        print(f"警告: 無法匯入 api.case_routes: {e}")

    # ── 請求日誌 ──
    logger = get_logger()
    logger.info("=" * 60)
    logger.info("財模助手 啟動")
    logger.info("=" * 60)

    @app.before_request
    def _before():
        g._req_start = time.time()

    @app.after_request
    def _after(response):
        user = session.get('user_email', 'anonymous')
        elapsed = int((time.time() - getattr(g, '_req_start', time.time())) * 1000)
        status = 'ERROR' if response.status_code >= 400 else 'OK'

        # 跳過靜態資源，避免日誌過多
        if not request.path.startswith('/static'):
            detail = f"method={request.method} path={request.path} status={response.status_code} time={elapsed}ms"
            # agent_chat 額外記錄查詢內容
            if request.path == '/api/agent_chat' and request.is_json:
                try:
                    query = (request.get_json(silent=True) or {}).get('query', '')
                    if query:
                        detail += f" | query={query[:120]}"
                except Exception:
                    pass
            log_action(user, 'http_request', detail, status)

        return response

    @app.teardown_request
    def _teardown(exc):
        if exc is not None:
            user = session.get('user_email', 'anonymous') if session else 'anonymous'
            log_error(user, 'request_teardown', exc)

    # 主頁路由
    @app.route('/')
    def index():
        return render_template('LLMweb.html')

    # 健康檢查
    @app.route('/api/health')
    def health_check():
        return jsonify({'status': 'healthy', 'service': 'AI Agent'})

    return app

# 創建應用實例
app = create_app()

if __name__ == '__main__':
    print("啟動 AI Agent Web 服務...")
    print("訪問: http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
