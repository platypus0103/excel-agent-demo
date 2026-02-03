"""
AI Agent 藍圖 - 處理聊天請求
"""
from flask import Blueprint, request, jsonify, send_file
from werkzeug.utils import secure_filename
import sys
import os

# 調整匯入路徑
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

try:
    from core.agent import AIAgent
    from config.settings import DEFAULT_CONFIG
except ImportError as e:
    print(f"嚴重警告: 無法匯入 AI Agent 核心模組: {e}")
    AIAgent = None
    DEFAULT_CONFIG = None

agent_bp = Blueprint('agent_bp', __name__)

# 全域 Agent 實例
_agent_instance = None

def get_agent():
    """獲取或創建 AI Agent 實例"""
    global _agent_instance
    if _agent_instance is None:
        if AIAgent is None or DEFAULT_CONFIG is None:
            raise RuntimeError("AI Agent 核心模組未載入")

        print("初始化 AI Agent...")
        _agent_instance = AIAgent(DEFAULT_CONFIG)
        print("AI Agent 初始化完成")

    return _agent_instance

@agent_bp.route('/agent_chat', methods=['POST'])
def agent_chat():
    """
    處理來自前端的 Agent 聊天請求
    """
    if not request.is_json:
        return jsonify({"error": "請求必須是 JSON 格式"}), 400

    data = request.json
    user_query = data.get('query')

    # 獲取參數（為了與 Xinbon_project 兼容，但這裡不會使用）
    equipment_cost_adj = float(data.get('equipment_cost_adj', 1.0))
    rent_adj = float(data.get('rent_adj', 1.0))
    loan_amount = data.get('loan_amount')

    # 獲取案場資訊
    case_name = data.get('case_name', '')
    original_filename = data.get('original_filename', '')
    sheet_name = data.get('sheet_name', None)  # 工作表名稱，可選

    if not user_query:
        return jsonify({"error": "請提供 'query' 參數"}), 400

    print(f"\n--- API 請求: 執行 Agent 對話 ---")
    print(f"收到查詢: {user_query}")
    print(f"案場資訊: case_name={case_name}, original_filename={original_filename}, sheet_name={sheet_name}")

    try:
        # 獲取 Agent 實例
        agent = get_agent()

        # 如果有案場的 Excel 檔案資訊，設定財務工具的檔案路徑
        if case_name and original_filename:
            import os
            parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            excel_filename = f"{case_name}_{original_filename}"
            excel_path = os.path.join(parent_dir, 'Excel', excel_filename)

            # 檢查檔案是否存在
            if os.path.exists(excel_path):
                agent.tool_manager.set_finance_excel_file(excel_path, sheet_name)
                print(f"已設定財務工具 Excel 檔案: {excel_path}")
            else:
                print(f"警告: Excel 檔案不存在: {excel_path}")
                # 嘗試尋找 Excel 資料夾中的任何檔案
                excel_dir = os.path.join(parent_dir, 'Excel')
                if os.path.exists(excel_dir):
                    excel_files = [f for f in os.listdir(excel_dir) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
                    if excel_files:
                        fallback_file = os.path.join(excel_dir, excel_files[0])
                        agent.tool_manager.set_finance_excel_file(fallback_file, sheet_name)
                        print(f"使用備用 Excel 檔案: {fallback_file}")

        # 檢查是否為滾算相關請求，直接調用工具處理
        import re
        if re.search(r'滾算|價金|price.*rolling|IRR|設備成本|計算|分析|模擬|cashmode|ratiomode|conditional|customize|執行', user_query, re.IGNORECASE):
            print("檢測到滾算相關請求，使用 llm_service 處理...")
            from services.llm_service import process_user_query

            # 構建當前聊天室的 Excel 路徑
            current_excel_path = None
            if case_name and original_filename:
                current_excel_path = os.path.join(parent_dir, 'Excel', f"{case_name}_{original_filename}")
                if not os.path.exists(current_excel_path):
                    # 如果檔案不存在，嘗試尋找備用檔案
                    excel_dir = os.path.join(parent_dir, 'Excel')
                    if os.path.exists(excel_dir):
                        excel_files = [f for f in os.listdir(excel_dir) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
                        if excel_files:
                            current_excel_path = os.path.join(excel_dir, excel_files[0])
                            print(f"原檔案不存在，使用備用檔案: {current_excel_path}")

            # 傳遞 excel_path 給 llm_service
            agent_response = process_user_query(user_query, excel_path=current_excel_path, sheet_name=sheet_name)
            excel_modified = True  # 滾算會產生 Excel 記錄
        else:
            # 其他請求調用 AI Agent 的 chat 方法
            agent_response = agent.chat(user_query)
            # 檢查是否有使用 Excel 工具
            excel_tools = ['write_excel_cell', 'delete_excel_cell', 'read_excel_cell',
                          'add_excel_sheet', 'delete_excel_sheet']
            excel_modified = any(tool in agent.last_used_tools for tool in excel_tools)

        print(f"Agent 回應成功 (Excel modified: {excel_modified})")
        return jsonify({
            "query": user_query,
            "response": agent_response,
            "excel_modified": excel_modified
        })

    except Exception as e:
        import traceback
        print(f"🚨 AI Agent 處理錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({"error": f"Agent 處理請求時發生錯誤: {str(e)}"}), 500

@agent_bp.route('/reset', methods=['POST'])
def reset_agent():
    """重置 Agent 對話歷史"""
    try:
        agent = get_agent()
        agent.reset(keep_system=True)
        return jsonify({
            "status": "success",
            "message": "對話已重置"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@agent_bp.route('/history', methods=['GET'])
def get_history():
    """獲取對話歷史"""
    try:
        agent = get_agent()
        messages = agent.conversation.get_messages()
        return jsonify({
            "status": "success",
            "history": messages
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@agent_bp.route('/get_excel_defaults', methods=['GET'])
def get_excel_defaults():
    """獲取 Excel 輸入公版的預設值"""
    try:
        from tool.equipment_cost_tool import EquipmentCostTool
        
        tool = EquipmentCostTool()
        excel_file = tool._find_excel_file(tool.excel_folder)
        
        # 先設定基本預設值（備用值）
        defaults = {
            "equipment_cost": 30000,
            "profit_rate": 0.2,
            "development_fee": 0,  # 預設為0，實際值從Excel讀取
            # 其他可能用到的預設值
            "boundary": 20000,  # 邊界價格的合理預設值
            "cash_step": 2000,  # 現金模式步幅預設值
            "ratio_step": 0.05,  # 比率模式步幅預設值
            # 條件模式預設值
            "max_value": 50000,
            "min_value": 30000,
            "cond_step_1": 2000,
            "cond_step_2": 1000,
            "cond_step_3": 500,
            # 自訂模式預設值
            "adjust_times": 5
        }
        
        # 如果有 Excel 檔案，則用 Excel 中的實際值覆寫預設值
        if excel_file:
            try:
                data = tool._read_excel_data(excel_file)
                # 用Excel中的實際數值覆寫預設值
                defaults.update({
                    "equipment_cost": data.get("equipment_cost", 30000),
                    "profit_rate": data.get("profit_rate", 0.2),
                    "development_fee": data.get("development_fee", 0)  # 從Excel C18讀取開發費
                })
                print(f"已從Excel讀取預設值：equipment_cost={defaults['equipment_cost']}, profit_rate={defaults['profit_rate']}, development_fee={defaults['development_fee']}")
            except Exception as e:
                print(f"警告：讀取 Excel 檔案失敗，使用備用預設值: {e}")
        
        return jsonify({
            "status": "success",
            "defaults": defaults,
            "excel_file": excel_file
        })
        
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"讀取 Excel 預設值時發生錯誤: {str(e)}"
        }), 500

@agent_bp.route('/upload_excel', methods=['POST'])
def upload_excel():
    """
    上傳 Excel 檔案到後端 Excel 資料夾
    """
    try:
        # 檢查是否有檔案
        if 'file' not in request.files:
            return jsonify({"error": "沒有檔案"}), 400

        file = request.files['file']

        # 檢查檔案名稱
        if file.filename == '':
            return jsonify({"error": "檔案名稱為空"}), 400

        # 獲取原始檔名（從前端傳來的參數）
        original_filename = request.form.get('original_filename', file.filename)

        # 檢查檔案類型
        allowed_extensions = {'.xlsx', '.xls'}
        file_ext = os.path.splitext(original_filename)[1].lower()
        if file_ext not in allowed_extensions:
            return jsonify({"error": f"不支援的檔案格式: {file_ext}"}), 400

        # 獲取案場 ID 和前端顯示名稱
        case_id = request.form.get('case_id', '')
        case_name = request.form.get('case_name', '')  # 前端顯示名稱（例如：表 1、財報）

        # 確保 Excel 目錄存在
        excel_dir = os.path.join(parent_dir, 'Excel')
        if not os.path.exists(excel_dir):
            os.makedirs(excel_dir)
            print(f"創建 Excel 目錄: {excel_dir}")

        # 使用前端名稱作為前綴（格式：前端名稱_原始檔名）
        filename = f"{case_name}_{original_filename}"  # 例如：表 1_公版.xlsx
        file_path = os.path.join(excel_dir, filename)
        file.save(file_path)
        print(f"Excel 檔案已儲存: {file_path}")
        print(f"檔案資訊: case_id={case_id}, case_name={case_name}, filename={filename}")

        return jsonify({
            "status": "success",
            "message": f"檔案已上傳: {original_filename}",
            "file_path": file_path,
            "filename": filename,
            "original_filename": original_filename  # 回傳真正的原始檔名
        })
    except Exception as e:
        import traceback
        print(f"檔案上傳錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({"error": f"檔案上傳失敗: {str(e)}"}), 500


@agent_bp.route('/read_excel/<case_id>', methods=['GET'])
def read_excel(case_id):
    """
    讀取指定案場的 Excel 檔案並轉換為前端格式
    """
    try:
        excel_dir = os.path.join(parent_dir, 'Excel')

        # 從查詢參數獲取前端名稱和原始檔名
        case_name = request.args.get('case_name', '')
        original_filename = request.args.get('original_filename', '')

        print(f"讀取 Excel 請求: case_id={case_id}, case_name={case_name}, original_filename={original_filename}")

        excel_file = None

        # 如果有前端名稱和原始檔名，直接組成完整檔名
        if case_name and original_filename:
            full_filename = f"{case_name}_{original_filename}"
            file_path = os.path.join(excel_dir, full_filename)
            if os.path.exists(file_path):
                excel_file = file_path
                print(f"找到檔案: {full_filename}")

        # 如果找不到，嘗試搜尋所有符合的檔案
        if not excel_file and os.path.exists(excel_dir):
            existing_files = os.listdir(excel_dir)
            print(f"搜尋 Excel 目錄中的檔案: {existing_files}")
            for filename in existing_files:
                if filename.endswith(('.xlsx', '.xls')):
                    excel_file = os.path.join(excel_dir, filename)
                    print(f"使用找到的第一個 Excel 檔案: {filename}")
                    break

        if not excel_file or not os.path.exists(excel_file):
            return jsonify({"error": "找不到指定案場的 Excel 檔案"}), 404


        # 使用 openpyxl 讀取 Excel
        import openpyxl
        from openpyxl.utils import get_column_letter

        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheets_data = []

        for sheet_index, sheet_name in enumerate(workbook.sheetnames):
            sheet = workbook[sheet_name]
            celldata = []

            # 讀取所有有值的儲存格
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and cell.value != '':
                        cell_data = {
                            'r': cell.row - 1,  # 轉換為 0-based
                            'c': cell.column - 1,  # 轉換為 0-based
                            'v': {
                                'v': cell.value,
                                'm': str(cell.value),
                                'ct': {
                                    'fa': 'General',
                                    't': 'n' if isinstance(cell.value, (int, float)) else 's'
                                }
                            }
                        }
                        celldata.append(cell_data)

            # 計算表格大小
            max_row = sheet.max_row if sheet.max_row else 20
            max_col = sheet.max_column if sheet.max_column else 10

            sheet_data = {
                'name': sheet_name,
                'index': str(sheet_index),
                'status': '1' if sheet_index == 0 else '0',
                'order': str(sheet_index),
                'celldata': celldata,
                'row': max(max_row, 20),
                'column': max(max_col, 10),
                'config': {},
                'pivotTable': None,
                'isPivotTable': False,
                'luckysheet_select_save': [],
                'calcChain': [],
                'hyperlink': {},
                'dataVerification': {}
            }
            sheets_data.append(sheet_data)

        print(f"成功讀取 Excel: {excel_file}, 共 {len(sheets_data)} 個工作表")

        return jsonify({
            'status': 'success',
            'data': sheets_data,
            'filename': os.path.basename(excel_file)
        })

    except Exception as e:
        import traceback
        print(f"讀取 Excel 錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({"error": f"讀取 Excel 失敗: {str(e)}"}), 500

@agent_bp.route('/rename_excel', methods=['POST'])
def rename_excel():
    """
    重新命名後端 Excel 檔案
    """
    try:
        data = request.json
        case_id = data.get('case_id')
        old_filename = data.get('old_filename')  # 完整檔名（例如：表 1_公版.xlsx）
        new_filename = data.get('new_filename')  # 完整檔名（例如：財報_公版.xlsx）

        print(f"重新命名請求: case_id={case_id}, old={old_filename}, new={new_filename}")

        if not all([old_filename, new_filename]):
            return jsonify({"error": "缺少必要參數"}), 400

        excel_dir = os.path.join(parent_dir, 'Excel')

        # 直接使用完整檔名（不再加上 case_id 前綴）
        old_path = os.path.join(excel_dir, old_filename)
        new_path = os.path.join(excel_dir, new_filename)

        print(f"檢查檔案: {old_path}")
        print(f"目標路徑: {new_path}")

        # 檢查舊檔案是否存在
        if not os.path.exists(old_path):
            # 列出目錄中的所有檔案幫助除錯
            if os.path.exists(excel_dir):
                existing_files = os.listdir(excel_dir)
                print(f"找不到檔案，Excel 目錄中的檔案: {existing_files}")
            return jsonify({"error": f"找不到原始檔案: {old_filename}"}), 404

        # 重新命名檔案
        os.rename(old_path, new_path)

        print(f"檔案已重新命名: {old_filename} -> {new_filename}")

        return jsonify({
            "status": "success",
            "message": "檔案重新命名成功",
            "new_filename": new_filename
        })

    except Exception as e:
        import traceback
        print(f"🚨 重新命名錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({"error": f"重新命名失敗: {str(e)}"}), 500

@agent_bp.route('/delete_excel', methods=['POST'])
def delete_excel():
    """
    刪除後端 Excel 檔案
    """
    try:
        data = request.json
        case_id = data.get('case_id')
        filename = data.get('filename')  # 完整檔名（例如：財報_公版.xlsx）

        print(f"刪除 Excel 請求: case_id={case_id}, filename={filename}")

        if not filename:
            return jsonify({"error": "缺少檔案名稱參數"}), 400

        excel_dir = os.path.join(parent_dir, 'Excel')
        file_path = os.path.join(excel_dir, filename)

        print(f"檢查檔案: {file_path}")

        # 檢查檔案是否存在
        if not os.path.exists(file_path):
            # 列出目錄中的所有檔案幫助除錯
            if os.path.exists(excel_dir):
                existing_files = os.listdir(excel_dir)
                print(f"找不到檔案，Excel 目錄中的檔案: {existing_files}")
            return jsonify({"error": f"找不到檔案: {filename}"}), 404

        # 刪除檔案
        os.remove(file_path)

        print(f"檔案已刪除: {filename}")

        return jsonify({
            "status": "success",
            "message": "檔案刪除成功",
            "deleted_filename": filename
        })

    except Exception as e:
        import traceback
        print(f"🚨 刪除檔案錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({"error": f"刪除檔案失敗: {str(e)}"}), 500

def _format_price_rolling_result(tool_result, mode):
    """
    格式化價金滾算結果為 Markdown
    支援 execute_price_rolling（寫入Excel）和 calculate_price_rolling（純計算）兩種格式

    Args:
        tool_result: 工具返回的結果
        mode: 模式名稱

    Returns:
        格式化後的 Markdown 字符串
    """
    if not tool_result.get("success"):
        return None

    response_parts = []
    response_parts.append(f"## 價金滾算結果 ({mode})\n")

    # 格式化 IRR 加上 %
    def format_irr(val):
        if val is None or val == 'N/A':
            return 'N/A'
        return f"{val}%"

    # 檢查是 execute_price_rolling 還是 calculate_price_rolling 的結果
    if "result" in tool_result:
        # execute_price_rolling 格式 - 從 result 字段提取數據
        result_data = tool_result.get("result", {})
        summary = tool_result.get("summary", {})

        response_parts.append("### 滾算摘要")
        response_parts.append(f"- **初始價金**: {summary.get('initial_cost', 'N/A')} 元/kW")
        response_parts.append(f"- **最終價金**: {summary.get('final_cost', 'N/A')} 元/kW")
        response_parts.append(f"- **總降幅**: {summary.get('total_reduction', 'N/A')} 元/kW\n")

        # 顯示結果表格
        adjustment_record = result_data.get("adjustment_record", [])
        profit_record = result_data.get("profit_record", [])
        irr_results = result_data.get("irr_results", [])

        if adjustment_record:
            response_parts.append("### 滾算明細")
            response_parts.append("| 價金/kW | 信邦利潤/kW | 專案法 IRR | 成本法 IRR | 權益法 IRR |")
            response_parts.append("| --- | --- | --- | --- | --- |")

            for i, price in enumerate(adjustment_record):
                profit = profit_record[i] if i < len(profit_record) else 'N/A'
                irr_data = irr_results[i] if i < len(irr_results) else {}
                project_irr = format_irr(irr_data.get('project_irr'))
                cost_irr = format_irr(irr_data.get('cost_method_irr'))
                equity_irr = format_irr(irr_data.get('equity_method_irr'))

                response_parts.append(f"| {price} | {profit} | {project_irr} | {cost_irr} | {equity_irr} |")

        return "\n".join(response_parts)

    # calculate_price_rolling 格式 - 簡化顯示
    params = tool_result.get("used_parameters", {})
    results = tool_result.get("results_summary", {})
    data_rows = results.get("data", [])

    # 滾算摘要
    if data_rows:
        initial_cost = data_rows[0][0] if data_rows else 'N/A'
        final_cost = data_rows[-1][0] if data_rows else 'N/A'
        total_reduction = initial_cost - final_cost if isinstance(initial_cost, (int, float)) and isinstance(final_cost, (int, float)) else 'N/A'

        response_parts.append("### 滾算摘要")
        response_parts.append(f"- **初始價金**: {initial_cost} 元/kW")
        response_parts.append(f"- **最終價金**: {final_cost} 元/kW")
        response_parts.append(f"- **總降幅**: {total_reduction} 元/kW\n")

    # 顯示結果表格（移除次數列，只顯示需要的欄位）
    response_parts.append("### 滾算明細")
    response_parts.append("| 價金/kW | 信邦利潤/kW | 專案法 IRR | 成本法 IRR | 權益法 IRR |")
    response_parts.append("| --- | --- | --- | --- | --- |")

    for row in data_rows:
        # row 格式: [price_per_kw, profit_per_kw, final_price_per_kw, project_irr, cost_method_irr, equity_method_irr]
        price = row[0] if len(row) > 0 else 'N/A'
        profit = row[1] if len(row) > 1 else 'N/A'
        project_irr = format_irr(row[3]) if len(row) > 3 else 'N/A'
        cost_irr = format_irr(row[4]) if len(row) > 4 else 'N/A'
        equity_irr = format_irr(row[5]) if len(row) > 5 else 'N/A'

        response_parts.append(f"| {price} | {profit} | {project_irr} | {cost_irr} | {equity_irr} |")

    return "\n".join(response_parts)


@agent_bp.route('/calculate_price_rolling', methods=['POST'])
def calculate_price_rolling():
    """
    處理價金滾算請求
    所有模式都直接調用工具，提高速度和精準度
    """
    if not request.is_json:
        return jsonify({"error": "請求必須是 JSON 格式"}), 400

    data = request.json

    # 獲取基本參數（優先使用網頁傳來的數值）
    mode = data.get('mode')  # CashMode, RatioMode, ConditionalMode, CustomizeMode
    equipment_cost = data.get('equipment_cost')
    profit_rate = data.get('profit_rate')
    development_fee = data.get('development_fee')  # 開發費，None表示用戶未輸入
    boundary = data.get('boundary')

    # 如果用戶沒有輸入開發費（None），則嘗試從Excel讀取預設值
    if development_fee is None:
        try:
            from tool.equipment_cost_tool import EquipmentCostTool
            tool = EquipmentCostTool()
            excel_file = tool._find_excel_file(tool.excel_folder)
            if excel_file:
                excel_data = tool._read_excel_data(excel_file)
                development_fee = excel_data.get("development_fee", 0)
                print(f"用戶未輸入開發費，從Excel讀取: {development_fee}")
            else:
                development_fee = 0
                print("無Excel檔案，開發費設為0")
        except Exception as e:
            development_fee = 0
            print(f"讀取Excel開發費失敗，設為0: {e}")
    else:
        print(f"使用用戶輸入的開發費: {development_fee}")

    # 獲取案場資訊
    case_name = data.get('case_name', '')
    original_filename = data.get('original_filename', '')
    sheet_name = data.get('sheet_name', None)

    # 驗證必要參數
    if not mode or boundary is None:
        return jsonify({"error": "缺少必要參數：mode 和 boundary"}), 400

    print(f"\n--- API 請求: 價金滾算 (直接調用工具) ---")
    print(f"模式: {mode}")
    print(f"參數: equipment_cost={equipment_cost}, profit_rate={profit_rate}, development_fee={development_fee}, boundary={boundary}")
    print(f"案場資訊: case_name={case_name}, original_filename={original_filename}")

    try:
        # 獲取 Agent 實例
        agent = get_agent()

        # 設定財務工具的 Excel 檔案
        current_excel_path = None
        excel_dir = os.path.join(parent_dir, 'Excel')

        if case_name and original_filename:
            # 方法1: 使用完整的 case_name + original_filename
            excel_filename = f"{case_name}_{original_filename}"
            current_excel_path = os.path.join(excel_dir, excel_filename)

            if not os.path.exists(current_excel_path):
                print(f"警告: Excel 檔案不存在: {current_excel_path}")
                current_excel_path = None

        # 方法2: 只有 case_name，搜尋以 case_name 開頭的檔案
        if not current_excel_path and case_name and os.path.exists(excel_dir):
            excel_files = [f for f in os.listdir(excel_dir)
                          if f.startswith(f"{case_name}_") and f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
            if excel_files:
                current_excel_path = os.path.join(excel_dir, excel_files[0])
                print(f"根據案場名稱找到 Excel 檔案: {current_excel_path}")

        # 方法3: 備用 - 使用目錄中的第一個 Excel 檔案
        if not current_excel_path and os.path.exists(excel_dir):
            excel_files = [f for f in os.listdir(excel_dir) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
            if excel_files:
                current_excel_path = os.path.join(excel_dir, excel_files[0])
                print(f"使用備用 Excel 檔案: {current_excel_path}")

        # 設定工具的 Excel 檔案
        if current_excel_path:
            agent.tool_manager.set_finance_excel_file(current_excel_path, sheet_name)
            agent.tool_manager.price_rolling_tool.set_excel_file(current_excel_path, sheet_name)
            print(f"已設定 Excel 檔案: {current_excel_path}")
        else:
            print(f"錯誤: 找不到任何 Excel 檔案，case_name={case_name}, original_filename={original_filename}")
            return jsonify({
                "status": "error",
                "error": "找不到 Excel 檔案。請先上傳 Excel 檔案再進行滾算。"
            }), 400

        # 轉換模式名稱：前端使用 CashMode/RatioMode 等，後端 execute_price_rolling 使用 cash/ratio 等
        mode_mapping = {
            "CashMode": "cash",
            "RatioMode": "ratio",
            "ConditionalMode": "conditional",
            "CustomizeMode": "customize"
        }
        backend_mode = mode_mapping.get(mode, mode.lower())

        # 構造工具參數（使用 calculate_price_rolling 工具，純計算不寫入 Excel）
        # 注意：calculate_price_rolling 使用 CashMode/RatioMode 等格式，不是 cash/ratio
        tool_args = {
            "mode": mode,  # 使用原始模式名稱 CashMode, RatioMode 等
            "boundary": boundary
        }

        # 添加共用參數
        if profit_rate is not None:
            tool_args["profit_rate"] = profit_rate
        if development_fee is not None:
            tool_args["development_fee"] = development_fee
        if sheet_name is not None:
            tool_args["sheet_name"] = sheet_name

        # 添加模式特定參數
        query_desc = f"{mode}"

        if mode == "CashMode":
            step = data.get('step')
            if step is not None:
                tool_args["step"] = int(step)
                query_desc += f" (步伐={step})"

        elif mode == "RatioMode":
            step = data.get('step')
            if step is not None:
                tool_args["step"] = float(step)
                query_desc += f" (比例={step})"

        elif mode == "ConditionalMode":
            max_value = data.get('max_value')
            min_value = data.get('min_value')
            step1 = data.get('step1')
            step2 = data.get('step2')
            step3 = data.get('step3')

            # calculate_price_rolling 使用這些參數名稱
            if max_value is not None:
                tool_args["max_value"] = max_value
            if min_value is not None:
                tool_args["min_value"] = min_value
            if step1 is not None:
                tool_args["step1"] = step1
            if step2 is not None:
                tool_args["step2"] = step2
            if step3 is not None:
                tool_args["step3"] = step3

            query_desc += f" (max={max_value}, min={min_value}, step1={step1}, step2={step2}, step3={step3})"

        elif mode == "CustomizeMode":
            adjust_times = data.get('adjust_times')
            steps = data.get('steps')

            if adjust_times is not None:
                tool_args["adjust_times"] = adjust_times
            if steps is not None:
                tool_args["steps"] = steps

            query_desc += f" (調整次數={adjust_times}, 步伐={steps})"

        print(f"工具參數: {tool_args}")

        # 使用 calculate_price_rolling 工具（純計算，不寫入 Excel，讓用戶自己決定是否儲存）
        tool_result = agent.tool_manager.execute_tool("calculate_price_rolling", tool_args)

        # 格式化結果
        if tool_result.get("success"):
            # 【快取】成功執行後，儲存參數到快取供「執行滾算紀錄」使用
            from services.llm_service import _store_rolling_cache
            _store_rolling_cache(current_excel_path, backend_mode, tool_args)
            print(f"[快取] 已儲存滾算參數到 llm_service 快取")

            agent_response = _format_price_rolling_result(tool_result, mode)

            if agent_response:
                print(f"{mode} 工具調用成功")
                return jsonify({
                    "status": "success",
                    "query": query_desc,
                    "response": agent_response
                })
            else:
                print(f"{mode} 結果格式化失敗")
                return jsonify({
                    "status": "error",
                    "error": "結果格式化失敗"
                }), 500
        else:
            error_msg = tool_result.get("message", "未知錯誤")
            print(f"{mode} 工具調用失敗: {error_msg}")
            return jsonify({
                "status": "error",
                "error": error_msg
            }), 400

    except Exception as e:
        import traceback
        print(f"🚨 價金滾算處理錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({
            "status": "error",
            "error": f"處理請求時發生錯誤: {str(e)}"
        }), 500


@agent_bp.route('/download_excel', methods=['GET'])
def download_excel():
    """
    下載當前聊天室的 Excel 檔案
    確保每個聊天室下載的是自己對應的 Excel 檔案
    """
    try:
        # 獲取案場資訊
        case_name = request.args.get('case_name', '')
        original_filename = request.args.get('original_filename', '')

        if not case_name:
            return jsonify({
                "status": "error",
                "error": "缺少必要參數：case_name"
            }), 400

        excel_dir = os.path.join(parent_dir, 'Excel')
        excel_path = None
        excel_filename = None

        print(f"[下載請求] 案場: {case_name}, 檔案: {original_filename}")

        # 方法1: 使用完整的 case_name + original_filename
        if original_filename:
            excel_filename = f"{case_name}_{original_filename}"
            excel_path = os.path.join(excel_dir, excel_filename)
            if not os.path.exists(excel_path):
                print(f"[下載] 完整路徑不存在: {excel_path}")
                excel_path = None

        # 方法2: 搜尋以 case_name 開頭的檔案
        if not excel_path and os.path.exists(excel_dir):
            excel_files = [f for f in os.listdir(excel_dir)
                          if f.startswith(f"{case_name}_") and f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
            if excel_files:
                excel_filename = excel_files[0]
                excel_path = os.path.join(excel_dir, excel_filename)
                print(f"[下載] 根據案場名稱找到: {excel_path}")

        # 檢查是否找到檔案
        if not excel_path or not os.path.exists(excel_path):
            print(f"[下載錯誤] 找不到任何符合的檔案，case_name={case_name}")
            return jsonify({
                "status": "error",
                "error": f"找不到案場「{case_name}」的 Excel 檔案"
            }), 404

        print(f"[下載] 發送檔案: {excel_path}")

        # 發送檔案
        return send_file(
            excel_path,
            as_attachment=True,
            download_name=excel_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        import traceback
        print(f"🚨 下載 Excel 錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({
            "status": "error",
            "error": f"下載失敗: {str(e)}"
        }), 500


@agent_bp.route('/list_case_sheets', methods=['GET'])
def list_case_sheets():
    """
    列出所有案場及其 Excel 檔案中的 sheets
    用於匯入表格功能
    """
    try:
        import openpyxl

        excel_dir = os.path.join(parent_dir, 'Excel')
        cases_data = []

        if not os.path.exists(excel_dir):
            return jsonify({
                "status": "success",
                "cases": []
            })

        # 掃描 Excel 目錄中的所有檔案
        excel_files = [f for f in os.listdir(excel_dir)
                      if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

        for filename in excel_files:
            file_path = os.path.join(excel_dir, filename)

            try:
                # 解析檔名，格式為 {case_name}_{original_filename}
                # 例如：表 1_公版.xlsx
                parts = filename.rsplit('_', 1)
                if len(parts) == 2:
                    case_name = parts[0]
                    original_filename = parts[1]
                else:
                    case_name = filename.rsplit('.', 1)[0]
                    original_filename = filename

                # 讀取 Excel 檔案的所有 sheet 名稱
                workbook = openpyxl.load_workbook(file_path, read_only=True)
                sheet_names = workbook.sheetnames
                workbook.close()

                cases_data.append({
                    "case_name": case_name,
                    "filename": filename,
                    "original_filename": original_filename,
                    "sheets": sheet_names,
                    "site_type": "single"  # 預設為單站，前端會根據 localStorage 更新
                })

                print(f"[list_case_sheets] {case_name}: {len(sheet_names)} sheets")

            except Exception as e:
                print(f"[list_case_sheets] 讀取 {filename} 失敗: {e}")
                continue

        return jsonify({
            "status": "success",
            "cases": cases_data
        })

    except Exception as e:
        import traceback
        print(f"🚨 列出案場 sheets 錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({
            "status": "error",
            "error": f"列出案場失敗: {str(e)}"
        }), 500


@agent_bp.route('/import_sheets', methods=['POST'])
def import_sheets():
    """
    將選定的 sheets 匯入到目標案場的 Excel 檔案中
    完整複製包含公式
    """
    try:
        import openpyxl
        from openpyxl.utils import get_column_letter
        from copy import copy

        data = request.json
        target_case_name = data.get('target_case_name')
        target_filename = data.get('target_filename')  # 可能為 None
        sheets_to_import = data.get('sheets_to_import', [])

        print(f"\n[import_sheets] 目標案場: {target_case_name}")
        print(f"[import_sheets] 目標檔案: {target_filename}")
        print(f"[import_sheets] 要匯入的 sheets: {sheets_to_import}")

        if not target_case_name:
            return jsonify({"status": "error", "error": "缺少目標案場名稱"}), 400

        if not sheets_to_import:
            return jsonify({"status": "error", "error": "請選擇至少一個 sheet"}), 400

        excel_dir = os.path.join(parent_dir, 'Excel')

        # 確定目標檔案路徑
        target_path = None
        new_filename = None

        if target_filename:
            # 有指定檔名
            full_filename = f"{target_case_name}_{target_filename}"
            target_path = os.path.join(excel_dir, full_filename)
            new_filename = target_filename

        if not target_path or not os.path.exists(target_path):
            # 搜尋以 target_case_name 開頭的檔案
            if os.path.exists(excel_dir):
                excel_files = [f for f in os.listdir(excel_dir)
                              if f.startswith(f"{target_case_name}_") and f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
                if excel_files:
                    target_path = os.path.join(excel_dir, excel_files[0])
                    # 從檔名提取 original_filename
                    new_filename = excel_files[0].replace(f"{target_case_name}_", "", 1)

        # 如果還是找不到，創建新檔案
        if not target_path or not os.path.exists(target_path):
            new_filename = "整合表格.xlsx"
            target_path = os.path.join(excel_dir, f"{target_case_name}_{new_filename}")
            # 創建空白 Excel 檔案
            wb = openpyxl.Workbook()
            wb.save(target_path)
            wb.close()
            print(f"[import_sheets] 創建新檔案: {target_path}")

        # 開啟目標檔案
        target_wb = openpyxl.load_workbook(target_path)
        print(f"[import_sheets] 開啟目標檔案: {target_path}")

        imported_count = 0

        for sheet_info in sheets_to_import:
            source_case_name = sheet_info.get('case_name')
            source_sheet_name = sheet_info.get('sheet_name')
            source_filename = sheet_info.get('filename')

            # 找到來源檔案
            source_path = os.path.join(excel_dir, source_filename)

            if not os.path.exists(source_path):
                print(f"[import_sheets] 來源檔案不存在: {source_path}")
                continue

            try:
                # 開啟來源檔案（不使用 data_only，保留公式）
                source_wb = openpyxl.load_workbook(source_path)

                if source_sheet_name not in source_wb.sheetnames:
                    print(f"[import_sheets] 來源 sheet 不存在: {source_sheet_name}")
                    source_wb.close()
                    continue

                source_sheet = source_wb[source_sheet_name]

                # 新 sheet 名稱：案場名稱_原sheet名稱
                new_sheet_name = f"{source_case_name}_{source_sheet_name}"

                # 確保名稱不重複（Excel sheet 名稱最多 31 字元）
                if len(new_sheet_name) > 31:
                    new_sheet_name = new_sheet_name[:31]

                # 如果已存在同名 sheet，加上編號
                base_name = new_sheet_name
                counter = 1
                while new_sheet_name in target_wb.sheetnames:
                    suffix = f"_{counter}"
                    max_base_len = 31 - len(suffix)
                    new_sheet_name = base_name[:max_base_len] + suffix
                    counter += 1

                # 創建新 sheet
                target_sheet = target_wb.create_sheet(title=new_sheet_name)

                # 複製儲存格內容、公式和樣式
                for row in source_sheet.iter_rows():
                    for cell in row:
                        new_cell = target_sheet.cell(row=cell.row, column=cell.column)

                        # 複製值或公式
                        if cell.data_type == 'f':
                            # 公式
                            new_cell.value = cell.value
                        else:
                            new_cell.value = cell.value

                        # 複製樣式
                        if cell.has_style:
                            new_cell.font = copy(cell.font)
                            new_cell.border = copy(cell.border)
                            new_cell.fill = copy(cell.fill)
                            new_cell.number_format = cell.number_format
                            new_cell.protection = copy(cell.protection)
                            new_cell.alignment = copy(cell.alignment)

                # 複製欄寬
                for col_letter, col_dim in source_sheet.column_dimensions.items():
                    target_sheet.column_dimensions[col_letter].width = col_dim.width

                # 複製列高
                for row_num, row_dim in source_sheet.row_dimensions.items():
                    target_sheet.row_dimensions[row_num].height = row_dim.height

                # 複製合併儲存格
                for merged_range in source_sheet.merged_cells.ranges:
                    target_sheet.merge_cells(str(merged_range))

                source_wb.close()
                imported_count += 1
                print(f"[import_sheets] 成功匯入: {source_case_name}/{source_sheet_name} -> {new_sheet_name}")

            except Exception as e:
                print(f"[import_sheets] 匯入 {source_case_name}/{source_sheet_name} 失敗: {e}")
                continue

        # 儲存目標檔案
        target_wb.save(target_path)
        target_wb.close()

        print(f"[import_sheets] 完成，共匯入 {imported_count} 個 sheets")

        return jsonify({
            "status": "success",
            "message": f"成功匯入 {imported_count} 個 sheets",
            "imported_count": imported_count,
            "new_filename": new_filename
        })

    except Exception as e:
        import traceback
        print(f"🚨 匯入 sheets 錯誤: {e}")
        print(traceback.format_exc())
        return jsonify({
            "status": "error",
            "error": f"匯入失敗: {str(e)}"
        }), 500