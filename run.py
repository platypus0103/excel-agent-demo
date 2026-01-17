"""
應用啟動腳本
"""
from app import app

if __name__ == '__main__':
    print("=" * 60)
    print("AI Agent Web 服務啟動中...")
    print("=" * 60)
    print("URL: http://localhost:5000")
    print("API: http://localhost:5000/api")
    print("按 Ctrl+C 停止服務")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
