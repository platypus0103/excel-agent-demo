"""
AI Agent 藍圖 - 處理聊天請求
"""
from flask import Blueprint, request, jsonify
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
                    excel_files = [f for f in os.listdir(excel_dir) if f.endswith(('.xlsx', '.xls'))]
                    if excel_files:
                        fallback_file = os.path.join(excel_dir, excel_files[0])
                        agent.tool_manager.set_finance_excel_file(fallback_file, sheet_name)
                        print(f"使用備用 Excel 檔案: {fallback_file}")

        # 調用 AI Agent 的 chat 方法
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

    # 顯示原始 IRR
    base_irr = tool_result.get("base_irr", {})
    response_parts.append("### 原始 Excel IRR (對照基準)")
    response_parts.append(f"- **專案法 IRR**: {format_irr(base_irr.get('project_irr'))}")
    response_parts.append(f"- **成本法 IRR**: {format_irr(base_irr.get('cost_method_irr'))}")
    response_parts.append(f"- **權益法 IRR**: {format_irr(base_irr.get('equity_method_irr'))}\n")

    # 顯示使用的參數
    params = tool_result.get("used_parameters", {})
    response_parts.append("### 使用參數")
    response_parts.append(f"- **初始價金 / kW**: {params.get('equipment_cost', 'N/A')}")
    response_parts.append(f"- **信邦利潤率**: {params.get('profit_rate', 'N/A')}")
    response_parts.append(f"- **開發費**: {params.get('development_fee', 'N/A')}")
    response_parts.append(f"- **邊界**: {params.get('boundary', 'N/A')}")

    # 顯示模式特定參數
    if mode == "CashMode" and params.get('step'):
        response_parts.append(f"- **固定步伐**: {params.get('step')}")
    elif mode == "RatioMode" and params.get('step'):
        response_parts.append(f"- **調整比例**: {params.get('step')}")

    response_parts.append("")  # 空行

    # 顯示結果表格
    results = tool_result.get("results_summary", {})
    columns = results.get("columns", [])
    data_rows = results.get("data", [])

    response_parts.append("### 滾算結果")

    # 構造 Markdown 表格
    header_map = {
        "price_per_kw": "價金/kW",
        "profit_per_kw": "信邦利潤/kW",
        "final_price_per_kw": "最終價金/kW",
        "project_irr": "專案法 IRR",
        "cost_method_irr": "成本法 IRR",
        "equity_method_irr": "權益法 IRR"
    }

    headers = [header_map.get(col, col) for col in columns]
    response_parts.append("| " + " | ".join(headers) + " |")
    response_parts.append("| " + " | ".join(["---"] * len(headers)) + " |")

    # IRR 列的索引
    irr_columns = {"project_irr", "cost_method_irr", "equity_method_irr"}

    for row in data_rows:
        formatted_row = []
        for i, val in enumerate(row):
            col_name = columns[i] if i < len(columns) else ""
            # 如果是 IRR 列，加上 %
            if col_name in irr_columns:
                formatted_row.append(format_irr(val))
            else:
                formatted_row.append(str(val))
        row_str = "| " + " | ".join(formatted_row) + " |"
        response_parts.append(row_str)

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
        if case_name and original_filename:
            excel_filename = f"{case_name}_{original_filename}"
            excel_path = os.path.join(parent_dir, 'Excel', excel_filename)

            if os.path.exists(excel_path):
                agent.tool_manager.set_finance_excel_file(excel_path, sheet_name)
                agent.tool_manager.price_rolling_tool.set_excel_file(excel_path, sheet_name)
                print(f"已設定 Excel 檔案: {excel_path}")
            else:
                print(f"警告: Excel 檔案不存在: {excel_path}")

        # 構造工具參數（所有模式統一處理）
        tool_args = {
            "mode": mode,
            "boundary": boundary
        }

        # 添加共用參數
        if equipment_cost is not None:
            tool_args["equipment_cost"] = equipment_cost
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
                tool_args["step"] = step
                query_desc += f" (步伐={step})"

        elif mode == "RatioMode":
            step = data.get('step')
            if step is not None:
                tool_args["step"] = step
                query_desc += f" (比例={step})"

        elif mode == "ConditionalMode":
            max_value = data.get('max_value')
            min_value = data.get('min_value')
            step1 = data.get('step1')
            step2 = data.get('step2')
            step3 = data.get('step3')

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

        # 直接調用工具
        tool_result = agent.tool_manager.execute_tool("calculate_price_rolling", tool_args)

        # 格式化結果
        if tool_result.get("success"):
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