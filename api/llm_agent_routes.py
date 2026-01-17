"""
LLM Agent 藍圖 - 專用於處理結合 Excel 數據分析的 Chat 請求
"""
from flask import Blueprint, request, jsonify
import sys
import os

# 調整匯入路徑
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

try:
    from services.llm_service import process_user_query
except ImportError:
    print("嚴重警告: 無法匯入 services.llm_service，Agent 功能將無法運作。")
    def process_user_query(query):
        return f"錯誤: 核心 Agent 服務 (llm_service) 載入失敗。"

llm_agent_bp = Blueprint('llm_agent_bp', __name__)

@llm_agent_bp.route('/agent_chat', methods=['POST'])
def agent_chat():
    """
    處理來自前端的 Agent 聊天請求。
    支援兩種模式：
    1. 一般聊天模式：用戶輸入自然語言查詢
    2. 滾算模式：網頁提供明確的滾算模式和參數
    """
    if not request.is_json:
        return jsonify({"error": "請求必須是 JSON 格式"}), 400

    data = request.json
    user_query = data.get('query')

    # 獲取敏感性分析參數 (每期調降金額)
    simulation_amount = data.get('simulation_amount', 0)

    # 獲取使用者上傳的 Excel 檔案路徑
    excel_path = data.get('excel_path')
    
    # 🆕 獲取網頁指定的滾算模式和參數（如果有）
    rolling_mode = data.get('rolling_mode')  # CashMode, RatioMode, ConditionalMode, CustomizeMode
    rolling_params = data.get('rolling_params')  # 包含所有滾算參數的字典

    if not user_query:
        return jsonify({"error": "請提供 'query' 參數"}), 400

    # 關鍵的偵錯輸出
    print(f"\n--- API 請求: 執行 Agent 邏輯 ---")
    print(f"收到查詢: {user_query}")
    print(f"敏感性分析調降金額: {simulation_amount}")
    print(f"使用 Excel 檔案: {excel_path if excel_path else '預設檔案'}")
    print(f"網頁指定滾算模式: {rolling_mode if rolling_mode else '未指定'}")
    print(f"網頁提供滾算參數: {rolling_params if rolling_params else '未提供'}")

    try:
        # 核心 Agent 執行點 - 傳遞所有參數
        agent_response = process_user_query(
            user_query,
            simulation_amount=simulation_amount,
            excel_path=excel_path,
            rolling_mode=rolling_mode,
            rolling_params=rolling_params
        )
        
        print(f"Agent 回應成功。")
        return jsonify({
            "query": user_query,
            "response": agent_response
        })

    except Exception as e:
        import traceback
        print(f"🚨 LLM Agent 處理錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({"error": f"Agent 處理請求時發生錯誤: {str(e)}"}), 500