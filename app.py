"""
Flask 主應用程式
AI Agent Web 介面
"""
import os
import sys
from flask import Flask, render_template, jsonify
from flask_cors import CORS

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
    app.config['SECRET_KEY'] = 'ai-agent-secret-key'

    # 註冊 AI Agent API 藍圖
    try:
        from api.agent_routes import agent_bp
        app.register_blueprint(agent_bp, url_prefix='/api')
        print("AI Agent API 路由註冊成功")
    except ImportError as e:
        print(f"警告: 無法匯入 api.agent_routes: {e}")

    # 主頁路由
    @app.route('/')
    def index():
        """主頁面 - AI Agent 聊天介面"""
        return render_template('LLMweb.html')

    # API 根路由
    @app.route('/api')
    def api_info():
        """API 資訊端點"""
        return jsonify({
            'name': 'AI Agent API',
            'version': '1.0.0',
            'description': 'AI Agent with Excel Tool - Web Interface',
            'endpoints': {
                'agent_chat': '/api/agent_chat',
                'health_check': '/api/health'
            }
        })

    # 健康檢查
    @app.route('/api/health')
    def health_check():
        """健康檢查端點"""
        return jsonify({
            'status': 'healthy',
            'service': 'AI Agent',
            'version': '1.0.0'
        })

    return app

# 創建應用實例
app = create_app()

if __name__ == '__main__':
    print("啟動 AI Agent Web 服務...")
    print("訪問: http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
