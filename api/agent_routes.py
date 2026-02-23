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


# ========== 待處理修改請求的快取 ==========
# 用於存儲需要用戶補充信息的修改請求
# key: case_name, value: { 'params': {...}, 'missing': [...], 'timestamp': ... }
_pending_edit_requests = {}


def _store_pending_edit(case_name: str, params: dict, missing: list):
    """存儲待處理的修改請求"""
    from datetime import datetime
    _pending_edit_requests[case_name] = {
        'params': params,
        'missing': missing,
        'timestamp': datetime.now()
    }
    print(f"[快取] 已存儲待處理修改請求: case={case_name}, missing={missing}")


def _get_pending_edit(case_name: str) -> dict:
    """獲取待處理的修改請求"""
    from datetime import datetime
    pending = _pending_edit_requests.get(case_name)
    if pending:
        # 檢查是否過期（5分鐘）
        elapsed = (datetime.now() - pending['timestamp']).total_seconds()
        if elapsed > 300:
            del _pending_edit_requests[case_name]
            print(f"[快取] 待處理請求已過期: case={case_name}")
            return None
        return pending
    return None


def _clear_pending_edit(case_name: str):
    """清除待處理的修改請求"""
    if case_name in _pending_edit_requests:
        del _pending_edit_requests[case_name]
        print(f"[快取] 已清除待處理請求: case={case_name}")


# ========== 待確認保存的滾算結果快取 ==========
# 用於存儲等待用戶確認是否保存的滾算結果
# key: case_name, value: { 'mode': ..., 'params': {...}, 'excel_path': ..., 'timestamp': ... }
_pending_rolling_save = {}


def _store_pending_rolling_save(case_name: str, mode: str, params: dict, excel_path: str):
    """存儲待確認保存的滾算結果"""
    from datetime import datetime
    _pending_rolling_save[case_name] = {
        'mode': mode,
        'params': params,
        'excel_path': excel_path,
        'timestamp': datetime.now()
    }
    print(f"[快取] 已存儲待確認滾算: case={case_name}, mode={mode}")


def _get_pending_rolling_save(case_name: str) -> dict:
    """獲取待確認保存的滾算結果"""
    from datetime import datetime
    pending = _pending_rolling_save.get(case_name)
    if pending:
        # 檢查是否過期（5分鐘）
        elapsed = (datetime.now() - pending['timestamp']).total_seconds()
        if elapsed > 300:
            del _pending_rolling_save[case_name]
            print(f"[快取] 待確認滾算已過期: case={case_name}")
            return None
        return pending
    return None


def _clear_pending_rolling_save(case_name: str):
    """清除待確認保存的滾算結果"""
    if case_name in _pending_rolling_save:
        del _pending_rolling_save[case_name]
        print(f"[快取] 已清除待確認滾算: case={case_name}")


def _parse_edit_request(query: str, existing_params: dict = None) -> dict:
    """
    解析修改工作表的請求，提取參數

    Args:
        query: 用戶輸入，如 "修改滾算紀錄4的綜合損益表 2020~2025的保險費改成-40000"
               或 "修改滾算紀錄3 綜合損益表的保險費 2020改為-40000 2023改為-20000 2027改為-50000"
        existing_params: 已有的參數（用於合併補充回答）

    Returns:
        解析結果字典，包含 missing_params 列表標記缺少哪些參數
    """
    import re

    result = {
        'success': False,
        'complete': False,  # 參數是否完整
        'sheet_name': None,
        'field_keyword': None,
        'year_spec': '全部',
        'new_value': None,
        'is_expense': None,
        'section_type': None,
        'year_value_map': None,  # 新增：支持非連續年份的年份-數值映射
        'missing_params': []  # 缺少的參數列表
    }

    # 如果有已存在的參數，先合併
    if existing_params:
        for key in ['sheet_name', 'field_keyword', 'year_spec', 'new_value', 'is_expense', 'section_type', 'year_value_map']:
            if existing_params.get(key) is not None:
                result[key] = existing_params[key]

    try:
        # 1. 提取工作表名稱（滾算紀錄 + 數字或其他名稱）
        sheet_match = re.search(r'(滾算紀錄\s*\d+|滾算紀錄_\d+|滾算紀錄\d+)', query)
        if sheet_match:
            result['sheet_name'] = sheet_match.group(1).replace(' ', '')
        else:
            # 嘗試其他格式
            sheet_match2 = re.search(r'(工作表|sheet)\s*[「「]?([^」」\s]+)[」」]?', query, re.IGNORECASE)
            if sheet_match2:
                result['sheet_name'] = sheet_match2.group(2)

        # 2. 提取區域類型（公版、綜合損益表、現金流量表）
        if re.search(r'公版', query):
            result['section_type'] = '公版'
        elif re.search(r'綜合損益表|損益表', query):
            result['section_type'] = '綜合損益表'
        elif re.search(r'現金流量表|流量表', query):
            result['section_type'] = '現金流量表'

        # 3. 優先檢測多個年份-數值對（非連續年份修改）
        # 匹配格式：2020改為-40000, 2023改為-20000 等
        year_value_pattern = r'(\d{4})\s*(?:改[為成]|設[為成]|變[為成])\s*(-?\d+(?:\.\d+)?)'
        year_value_matches = re.findall(year_value_pattern, query)

        if len(year_value_matches) >= 2:
            # 有多個年份-數值對，使用 year_value_map 模式
            year_value_map = {}
            is_any_negative = False
            for year_str, value_str in year_value_matches:
                year = int(year_str)
                value = float(value_str)
                year_value_map[year] = value
                if value < 0:
                    is_any_negative = True

            result['year_value_map'] = year_value_map
            result['year_spec'] = 'multiple'  # 標記為多年份模式
            # 如果任一值為負數，標記為支出
            if is_any_negative:
                result['is_expense'] = True
            print(f"檢測到多年份-數值對: {year_value_map}")
        else:
            # 4. 提取年份規格（單一年份或範圍）
            year_range = re.search(r'(\d{4})\s*[~\-到至]\s*(\d{4})', query)
            if year_range:
                result['year_spec'] = f"{year_range.group(1)}~{year_range.group(2)}"
            else:
                single_year = re.search(r'(\d{4})\s*年', query)
                if single_year:
                    result['year_spec'] = single_year.group(1)
                elif re.search(r'全部|所有|全年', query):
                    result['year_spec'] = '全部'

            # 5. 提取數值（支援負數）- 只在非 year_value_map 模式時執行
            value_match = re.search(r'(?:改成|改為|設為|設成|變成|變為)\s*(-?\d+(?:\.\d+)?)', query)
            if value_match:
                result['new_value'] = float(value_match.group(1))
            else:
                value_match2 = re.search(r'(?:支出|收入)\s*(-?\d+(?:\.\d+)?)', query)
                if value_match2:
                    result['new_value'] = float(value_match2.group(1))
                else:
                    # 嘗試匹配獨立的數字（可能是補充回答）
                    standalone_num = re.search(r'^(-?\d+(?:\.\d+)?)$', query.strip())
                    if standalone_num:
                        result['new_value'] = float(standalone_num.group(1))

        # 6. 判斷是否為支出（只在非 year_value_map 模式或尚未判斷時執行）
        if result['is_expense'] is None:
            if re.search(r'支出|費用|成本|花費', query):
                result['is_expense'] = True
            elif re.search(r'收入|進帳|營收', query):
                result['is_expense'] = False
            elif result['new_value'] is not None and result['new_value'] < 0:
                result['is_expense'] = True
                result['new_value'] = abs(result['new_value'])

        # 7. 提取欄位關鍵字
        field_query = query
        field_query = re.sub(r'滾算紀錄\s*\d+|滾算紀錄_\d+', '', field_query)
        field_query = re.sub(r'公版|綜合損益表|損益表|現金流量表|流量表', '', field_query)
        # 移除年份-數值對（支持多個）
        field_query = re.sub(r'\d{4}\s*(?:改[為成]|設[為成]|變[為成])\s*-?\d+(?:\.\d+)?', '', field_query)
        field_query = re.sub(r'\d{4}\s*[~\-到至]\s*\d{4}|\d{4}\s*年?|全部|所有|全年', '', field_query)
        field_query = re.sub(r'(?:改成|改為|設為|設成|變成|變為)\s*-?\d+(?:\.\d+)?', '', field_query)
        field_query = re.sub(r'(?:支出|收入)\s*-?\d+(?:\.\d+)?', '', field_query)
        field_query = re.sub(r'我想|想要|想|要|幫我|請|修改|更改|把|的|將|中|裡|內|改|成|為|變|設定|可以|能|能否|麻煩', '', field_query)
        field_query = re.sub(r'[，。、！？\s]+', ' ', field_query).strip()

        words = [w.strip() for w in field_query.split() if len(w.strip()) >= 2]
        if words:
            result['field_keyword'] = max(words, key=len)
        else:
            common_fields = ['保險費', '設備費用', '租金', '管理費', '電費', '維護費', '折舊', '利息']
            for field in common_fields:
                if field in query:
                    result['field_keyword'] = field
                    break

        # 8. 檢查缺少哪些必要參數
        if not result['sheet_name']:
            result['missing_params'].append('sheet_name')
        if not result['section_type']:
            result['missing_params'].append('section_type')
        if not result['field_keyword']:
            result['missing_params'].append('field_keyword')
        # 如果有 year_value_map，則不需要 new_value
        if result['year_value_map'] is None and result['new_value'] is None:
            result['missing_params'].append('new_value')

        # 9. 判斷是否完整
        if not result['missing_params']:
            result['complete'] = True
            result['success'] = True
        elif len(result['missing_params']) < 4:
            # 有部分參數，標記為部分成功（可以反問）
            result['success'] = True

        print(f"解析結果: {result}")
        return result

    except Exception as e:
        result['message'] = f'解析錯誤: {str(e)}'
        return result


def _generate_followup_question(missing_params: list, parsed: dict) -> str:
    """
    根據缺少的參數生成反問訊息
    """
    questions = []

    if 'sheet_name' in missing_params:
        questions.append("請問要修改哪一個工作表？（例如：滾算紀錄1、滾算紀錄2...）")

    if 'section_type' in missing_params:
        questions.append("請問要修改哪個區域？\n- **公版**（第1-36行）\n- **綜合損益表**（第37-64行）\n- **現金流量表**（第86-115行）")

    if 'field_keyword' in missing_params:
        questions.append("請問要修改哪個欄位？（例如：保險費、設備費用、租金...）")

    if 'new_value' in missing_params:
        questions.append("請問要改成多少？（例如：40000 或 支出40000）")

    # 顯示已識別的參數
    recognized = []
    if parsed.get('sheet_name'):
        recognized.append(f"工作表: {parsed['sheet_name']}")
    if parsed.get('section_type'):
        recognized.append(f"區域: {parsed['section_type']}")
    if parsed.get('field_keyword'):
        recognized.append(f"欄位: {parsed['field_keyword']}")
    if parsed.get('year_value_map'):
        # 多年份模式
        year_values = ", ".join([f"{y}年={v}" for y, v in sorted(parsed['year_value_map'].items())])
        recognized.append(f"年份-數值: {year_values}")
    elif parsed.get('new_value') is not None:
        recognized.append(f"數值: {parsed['new_value']}")
    if parsed.get('year_spec') and parsed['year_spec'] != 'multiple':
        recognized.append(f"年份: {parsed['year_spec']}")

    response = "### 需要更多資訊\n\n"

    if recognized:
        response += "**已識別的資訊**:\n"
        for r in recognized:
            response += f"- {r}\n"
        response += "\n"

    response += "**請補充以下資訊**:\n\n"
    for i, q in enumerate(questions, 1):
        response += f"{i}. {q}\n\n"

    return response


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

        # 檢查請求類型並分流處理
        import re

        # 0a. 先檢查是否有待確認保存的滾算結果
        pending_rolling = _get_pending_rolling_save(case_name) if case_name else None

        if pending_rolling:
            # 檢查用戶是否確認保存
            is_confirm_save = re.search(r'^(是|好|要|對|確認|保存|儲存|存|yes|ok|save)$', user_query.strip(), re.IGNORECASE)
            is_decline_save = re.search(r'^(否|不|不要|取消|no|cancel)$', user_query.strip(), re.IGNORECASE)

            if is_confirm_save:
                # 用戶確認保存，執行 execute_price_rolling
                print("用戶確認保存滾算結果，執行 execute_price_rolling...")
                _clear_pending_rolling_save(case_name)

                # 準備保存參數，加入 excel_file
                save_params = pending_rolling['params'].copy()
                save_params['excel_file'] = pending_rolling.get('excel_path')

                # 轉換模式名稱：CashMode -> cash, RatioMode -> ratio 等
                mode_mapping = {
                    "CashMode": "cash",
                    "RatioMode": "ratio",
                    "ConditionalMode": "conditional",
                    "CustomizeMode": "customize"
                }
                original_mode = save_params.get('mode', '')
                save_params['mode'] = mode_mapping.get(original_mode, original_mode.lower())

                print(f"保存參數: {save_params}")

                # 執行保存
                save_result = agent.tool_manager.execute_tool("execute_price_rolling", save_params)

                if save_result.get("success"):
                    agent_response = "滾算紀錄已儲存"
                    excel_modified = True
                else:
                    agent_response = f"儲存失敗：{save_result.get('message', '未知錯誤')}"
                    excel_modified = False

                return jsonify({
                    "query": user_query,
                    "response": agent_response,
                    "excel_modified": excel_modified
                })

            elif is_decline_save:
                # 用戶拒絕保存
                print("用戶拒絕保存滾算結果")
                _clear_pending_rolling_save(case_name)

                return jsonify({
                    "query": user_query,
                    "response": "好的，滾算結果不會儲存。您可以繼續進行其他操作。",
                    "excel_modified": False
                })

            else:
                # 用戶輸入其他內容，視為不保存並繼續處理新請求
                print("用戶未確認保存，清除待確認快取，處理新請求...")
                _clear_pending_rolling_save(case_name)
                # 繼續往下執行，處理新請求

        # 0b. 檢查是否有待處理的修改請求（用戶可能在回答反問）
        pending = _get_pending_edit(case_name) if case_name else None

        if pending:
            print(f"發現待處理的修改請求，嘗試合併用戶回答...")
            # 將用戶的回答與之前的參數合併
            parsed = _parse_edit_request(user_query, pending['params'])

            if parsed.get('complete'):
                # 參數完整，執行修改
                print("參數已完整，執行修改...")
                _clear_pending_edit(case_name)

                # 設定 Excel 檔案路徑
                if case_name and original_filename:
                    excel_path = os.path.join(parent_dir, 'Excel', f"{case_name}_{original_filename}")
                    if os.path.exists(excel_path):
                        agent.tool_manager.excel_tool.file_path = excel_path

                # 直接調用工具
                tool_args = {
                    'sheet_name': parsed['sheet_name'],
                    'field_keyword': parsed['field_keyword'],
                    'year_spec': parsed['year_spec'],
                }
                # 如果有 year_value_map，使用多年份模式
                if parsed.get('year_value_map'):
                    tool_args['year_value_map'] = parsed['year_value_map']
                else:
                    tool_args['new_value'] = parsed['new_value']

                if parsed.get('is_expense') is not None:
                    tool_args['is_expense'] = parsed['is_expense']
                if parsed.get('section_type'):
                    tool_args['section_type'] = parsed['section_type']

                result = agent.tool_manager.execute_tool('edit_sheet_by_field', tool_args)

                if result.get('success'):
                    excel_modified = True
                    agent_response = f"### 修改完成\n\n{result.get('message', '已成功修改')}"
                    if result.get('matched_field'):
                        agent_response += f"\n\n**匹配欄位**: {result['matched_field']}"
                    if result.get('years_modified'):
                        years = result['years_modified']
                        # 判斷是多年份模式還是傳統模式
                        if result.get('year_value_map'):
                            # 多年份模式：顯示每個年份的值
                            agent_response += f"\n**修改年份**: {', '.join(map(str, sorted(years)))}（共 {len(years)} 年）"
                        elif len(years) > 1:
                            agent_response += f"\n**修改年份**: {min(years)} ~ {max(years)}（共 {len(years)} 年）"
                        else:
                            agent_response += f"\n**修改年份**: {years[0]}"
                    # 顯示數值詳情
                    if result.get('year_value_map'):
                        # 多年份模式：顯示每個年份對應的值
                        agent_response += "\n**設定值明細**:"
                        for year, value in sorted(result['year_value_map'].items(), key=lambda x: int(x[0])):
                            agent_response += f"\n  - {year}年: {value}"
                    elif result.get('new_value') is not None:
                        agent_response += f"\n**新數值**: {result['new_value']}"
                else:
                    excel_modified = False
                    agent_response = f"### 修改失敗\n\n{result.get('message', '未知錯誤')}"

                return jsonify({
                    "query": user_query,
                    "response": agent_response,
                    "excel_modified": excel_modified
                })

            else:
                # 參數仍不完整，繼續反問
                print(f"參數仍不完整，缺少: {parsed.get('missing_params')}")
                _store_pending_edit(case_name, parsed, parsed.get('missing_params', []))
                agent_response = _generate_followup_question(parsed.get('missing_params', []), parsed)
                return jsonify({
                    "query": user_query,
                    "response": agent_response,
                    "excel_modified": False
                })

        # 1. 檢查是否為「修改 Excel」請求（優先級最高）
        is_edit_request = re.search(r'修改|更改|改成|設定|把.+改|變更', user_query, re.IGNORECASE)
        # 判斷是否為工作表修改請求：需要有工作表名稱，或者有常見欄位名稱
        has_sheet_name = re.search(r'滾算紀錄|sheet|工作表', user_query, re.IGNORECASE)
        # 匹配所有包含「費」的詞（如回收費、保險費、運維費等），以及其他常見欄位
        has_field_keyword = re.search(r'\w*費\w*|租金|折舊|利息|成本|收入|支出|價金|利潤', user_query)
        is_sheet_edit = is_edit_request and (has_sheet_name or has_field_keyword)

        if is_sheet_edit:
            # 修改工作表的請求，直接解析參數並調用工具（繞過 LLM）
            print("檢測到修改工作表請求，直接調用 edit_sheet_by_field 工具...")

            # 解析用戶輸入
            parsed = _parse_edit_request(user_query)

            if parsed.get('complete'):
                # 參數完整，直接執行
                print("參數完整，直接執行修改...")

                # 設定 Excel 檔案路徑
                if case_name and original_filename:
                    excel_path = os.path.join(parent_dir, 'Excel', f"{case_name}_{original_filename}")
                    if os.path.exists(excel_path):
                        agent.tool_manager.excel_tool.file_path = excel_path
                        print(f"已設定 Excel 路徑: {excel_path}")

                # 直接調用工具
                tool_args = {
                    'sheet_name': parsed['sheet_name'],
                    'field_keyword': parsed['field_keyword'],
                    'year_spec': parsed['year_spec'],
                }

                # 如果有 year_value_map，使用多年份模式
                if parsed.get('year_value_map'):
                    tool_args['year_value_map'] = parsed['year_value_map']
                else:
                    tool_args['new_value'] = parsed['new_value']

                # 可選參數
                if parsed.get('is_expense') is not None:
                    tool_args['is_expense'] = parsed['is_expense']
                if parsed.get('section_type'):
                    tool_args['section_type'] = parsed['section_type']

                print(f"工具參數: {tool_args}")

                # 執行工具
                result = agent.tool_manager.execute_tool('edit_sheet_by_field', tool_args)

                if result.get('success'):
                    excel_modified = True
                    agent_response = f"### 修改完成\n\n{result.get('message', '已成功修改')}"

                    # 添加詳細信息
                    if result.get('matched_field'):
                        agent_response += f"\n\n**匹配欄位**: {result['matched_field']}"
                    if result.get('years_modified'):
                        years = result['years_modified']
                        # 判斷是多年份模式還是傳統模式
                        if result.get('year_value_map'):
                            # 多年份模式：顯示每個年份
                            agent_response += f"\n**修改年份**: {', '.join(map(str, sorted(years)))}（共 {len(years)} 年）"
                        elif len(years) > 1:
                            agent_response += f"\n**修改年份**: {min(years)} ~ {max(years)}（共 {len(years)} 年）"
                        else:
                            agent_response += f"\n**修改年份**: {years[0]}"
                    # 顯示數值詳情
                    if result.get('year_value_map'):
                        # 多年份模式：顯示每個年份對應的值
                        agent_response += "\n**設定值明細**:"
                        for year, value in sorted(result['year_value_map'].items(), key=lambda x: int(x[0])):
                            agent_response += f"\n  - {year}年: {value}"
                    elif result.get('new_value') is not None:
                        agent_response += f"\n**新數值**: {result['new_value']}"
                else:
                    excel_modified = False
                    agent_response = f"### 修改失敗\n\n{result.get('message', '未知錯誤')}"
                    if result.get('available_fields'):
                        agent_response += f"\n\n**可用欄位**:\n" + "\n".join([f"- {f}" for f in result['available_fields'][:10]])
                    if result.get('available_sheets'):
                        agent_response += f"\n\n**可用工作表**: {', '.join(result['available_sheets'])}"
            elif parsed.get('success') and parsed.get('missing_params'):
                # 部分參數已解析，生成反問（防呆機制）
                print(f"參數不完整，缺少: {parsed.get('missing_params')}，生成反問...")

                # 存儲待處理請求
                _store_pending_edit(case_name, parsed, parsed.get('missing_params', []))

                # 生成反問
                agent_response = _generate_followup_question(parsed.get('missing_params', []), parsed)
                excel_modified = False

            else:
                # 完全無法解析，交給 LLM 處理（混合方案）
                print(f"直接解析失敗（{parsed.get('message')}），交給 LLM 處理...")

                # 設定 Excel 檔案路徑
                if case_name and original_filename:
                    excel_path = os.path.join(parent_dir, 'Excel', f"{case_name}_{original_filename}")
                    if os.path.exists(excel_path):
                        agent.tool_manager.excel_tool.file_path = excel_path

                # 調用 LLM
                agent_response = agent.chat(user_query)

                # 檢查是否有使用 Excel 編輯工具
                excel_tools = ['write_excel_cell', 'delete_excel_cell', 'read_excel_cell',
                              'add_excel_sheet', 'delete_excel_sheet', 'edit_sheet_by_field']
                excel_modified = any(tool in agent.last_used_tools for tool in excel_tools)

                # 如果 LLM 也沒有成功處理，提供格式提示
                if not excel_modified and 'edit_sheet_by_field' not in agent.last_used_tools:
                    agent_response += "\n\n---\n**提示**: 如果修改未成功，請嘗試以下格式：\n"
                    agent_response += "- 修改滾算紀錄2全部的保險費 變成支出40000\n"
                    agent_response += "- 修改滾算紀錄3的公版 保險費改成40000\n"
                    agent_response += "- 修改滾算紀錄4的綜合損益表 2020~2025的租金改成-50000\n"
                    agent_response += "- 修改滾算紀錄3的綜合損益表 保險費 2020改為-40000 2023改為-20000 2027改為-50000（非連續年份）"

        # 2. 檢查是否為滾算相關請求
        elif re.search(r'滾算|價金|price.*rolling|IRR|設備成本|計算|分析|模擬|cashmode|ratiomode|conditional|customize|執行', user_query, re.IGNORECASE):
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
            # 3. 其他請求調用 AI Agent 的 chat 方法
            agent_response = agent.chat(user_query)
            # 檢查是否有使用 Excel 工具
            excel_tools = ['write_excel_cell', 'delete_excel_cell', 'read_excel_cell',
                          'add_excel_sheet', 'delete_excel_sheet', 'edit_sheet_by_field']
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

        # 使用 data_only=True 讀取計算後的值（用於前端顯示）
        # 原始 Excel 檔案中的公式會被保留，匯出時直接下載原始檔案
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
            # 【快取】成功執行後，儲存參數到快取供確認保存使用
            from services.llm_service import _store_rolling_cache
            _store_rolling_cache(current_excel_path, backend_mode, tool_args)
            print(f"[快取] 已儲存滾算參數到 llm_service 快取")

            # 【新增】存儲待確認保存的滾算結果
            if case_name:
                _store_pending_rolling_save(case_name, backend_mode, tool_args, current_excel_path)

            agent_response = _format_price_rolling_result(tool_result, mode)

            if agent_response:
                # 【新增】添加詢問是否保存的消息
                agent_response += "\n\n---\n**是否要將此滾算結果儲存到 Excel？**\n請回覆「是」或「儲存」來保存，或直接進行下一步操作。"

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


@agent_bp.route('/save_excel', methods=['POST'])
def save_excel():
    """
    將 Luckysheet 數據保存回 Excel 檔案（保留公式）

    重要：此函數會保護原始 Excel 中的公式不被覆蓋。
    如果原始儲存格有公式，且 Luckysheet 數據中沒有公式資訊，則保留原始公式。
    """
    try:
        import openpyxl
        from openpyxl.utils import get_column_letter

        data = request.json
        case_name = data.get('case_name', '')
        sheets_data = data.get('sheets', [])

        if not case_name:
            return jsonify({"status": "error", "error": "缺少案場名稱"}), 400

        excel_dir = os.path.join(parent_dir, 'Excel')
        excel_path = None

        # 搜尋對應的 Excel 檔案
        if os.path.exists(excel_dir):
            for f in os.listdir(excel_dir):
                if f.startswith(f"{case_name}_") and f.endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                    excel_path = os.path.join(excel_dir, f)
                    break

        if not excel_path or not os.path.exists(excel_path):
            return jsonify({"status": "error", "error": f"找不到案場「{case_name}」的 Excel 檔案"}), 404

        print(f"[save_excel] 保存到: {excel_path}")

        # 讀取原始 Excel（保留公式）
        workbook = openpyxl.load_workbook(excel_path, data_only=False)

        # 建立 sheet 名稱到索引的映射
        sheet_name_map = {sheet.get('name'): sheet for sheet in sheets_data}

        formulas_preserved = 0  # 統計保留的公式數量

        for sheet_name in workbook.sheetnames:
            if sheet_name not in sheet_name_map:
                continue

            ws = workbook[sheet_name]
            sheet_data = sheet_name_map[sheet_name]
            celldata = sheet_data.get('celldata', [])

            # 建立座標到儲存格數據的映射
            cell_map = {}
            for cell in celldata:
                r = cell.get('r', 0)
                c = cell.get('c', 0)
                cell_map[(r, c)] = cell

            # 更新儲存格
            for (r, c), cell_info in cell_map.items():
                v = cell_info.get('v', {})
                excel_cell = ws.cell(row=r+1, column=c+1)

                # 檢查原始儲存格是否有公式
                original_has_formula = False
                if excel_cell.value and isinstance(excel_cell.value, str) and excel_cell.value.startswith('='):
                    original_has_formula = True

                if isinstance(v, dict):
                    # 如果 Luckysheet 數據有公式，使用公式
                    if 'f' in v and v['f']:
                        excel_cell.value = '=' + v['f'] if not v['f'].startswith('=') else v['f']
                    elif original_has_formula:
                        # 原始儲存格有公式，但 Luckysheet 沒有公式資訊，保留原始公式
                        formulas_preserved += 1
                        # 不做任何修改，保留原始公式
                    elif 'v' in v:
                        # 原始沒有公式，使用 Luckysheet 的值
                        excel_cell.value = v['v']
                else:
                    if original_has_formula:
                        # 保留原始公式
                        formulas_preserved += 1
                    else:
                        excel_cell.value = v

        # 保存檔案
        workbook.save(excel_path)
        workbook.close()

        print(f"[save_excel] 保存成功，保留了 {formulas_preserved} 個公式")
        return jsonify({"status": "success", "message": f"Excel 已保存（保留了 {formulas_preserved} 個公式）"})

    except Exception as e:
        import traceback
        print(f"[save_excel] 保存失敗: {e}")
        print(traceback.format_exc())
        return jsonify({"status": "error", "error": f"保存失敗: {str(e)}"}), 500


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


def create_aggregation_sheet(target_wb, cases_info):
    """
    在 target_wb 中建立「彙整總表」工作表。
    cases_info: list of {
        'display_name': str,   # B欄顯示的案場名稱
        'sheet_name':   str,   # 實際工作表名稱（用於公式引用）
        'start_year':   int,
        'end_year':     int
    }
    """
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import PatternFill, Font

    ITEMS = [
        {"name": "預估發電度數", "row": 40},
        {"name": "預估收入",     "row": 41},
        {"name": "租金",         "row": 45},
        {"name": "運維",         "row": 48},
        {"name": "保險",         "row": 49},
        {"name": "模組回收費",   "row": 50},
        {"name": "所得稅",       "row": 54},
        {"name": "設備折舊",     "row": 67},
    ]

    # 各項目年總計列的背景色
    ITEM_TOTAL_COLORS = {
        "預估發電度數": "FFFF99",  # 淺黃
        "預估收入":     "F2DCDB",  # 紅色輔色2較淺80%
        "租金":         "FCD5B5",  # 橙色輔色6較淺60%
        "運維":         "DCE6F1",  # 藍色輔色1較淺80%
        "保險":         "F2F2F2",  # 白色背景1較深5%
        "模組回收費":   "F2DCDB",  # 紅色輔色2較淺80%
        "所得稅":       "FFFF99",  # 淺黃
        "設備折舊":     "DAEEF3",  # 青色輔色5較淺80%
    }

    YEAR_ROW_COLOR = "1F3864"  # 藍色輔色1較深50%

    def fill_row(ws, row_num, start_col, end_col, hex_color, font_color=None):
        """將指定列的 start_col ~ end_col 填上背景色，並可選擇指定字體顏色"""
        fill = PatternFill(fill_type="solid", fgColor=hex_color)
        font = Font(color=font_color) if font_color else None
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.fill = fill
            if font:
                cell.font = font

    valid_cases = [c for c in cases_info if c.get('start_year') and c.get('end_year')]
    if not valid_cases:
        print("[彙整] 沒有有效的案場年份資料，跳過彙整")
        return

    overall_start = min(c['start_year'] for c in valid_cases)
    overall_end   = max(c['end_year']   for c in valid_cases)
    year_axis     = list(range(overall_start, overall_end + 1))
    n_years       = len(year_axis)
    last_col      = 3 + n_years  # C=3, D=4, 最後一個年份欄 = 4+(n_years-1) = 3+n_years

    # 建立（或清空重用）彙整總表
    SHEET_NAME = "彙整總表"
    if SHEET_NAME in target_wb.sheetnames:
        # 清空原有內容，保留工作表位置（通常在第一頁）
        agg_ws = target_wb[SHEET_NAME]
        for row in agg_ws.iter_rows():
            for cell in row:
                cell.value = None
    else:
        agg_ws = target_wb.create_sheet(title=SHEET_NAME, index=0)

    # Row 4：全域年份標題（只出現一次）
    agg_ws['C4'] = "年份"
    for i, year in enumerate(year_axis):
        agg_ws.cell(row=4, column=4 + i, value=year)
    fill_row(agg_ws, 4, 3, last_col, YEAR_ROW_COLOR, font_color="FFFFFF")

    current_row = 5

    for item_idx, item in enumerate(ITEMS):
        item_name   = item["name"]
        source_row  = item["row"]
        total_color = ITEM_TOTAL_COLORS.get(item_name, "FFFFFF")

        # 項目名稱列
        agg_ws.cell(row=current_row, column=3, value=item_name)
        # 第一個項目不重複年份（row 4 已有），後續項目在同列補上年份數字並上色
        if item_idx > 0:
            for i, year in enumerate(year_axis):
                agg_ws.cell(row=current_row, column=4 + i, value=year)
            fill_row(agg_ws, current_row, 3, last_col, YEAR_ROW_COLOR, font_color="FFFFFF")
        current_row += 1

        case_start_row = current_row

        # 各案場列（公式一律加 ABS 確保正數）
        for case in valid_cases:
            display_name = case['display_name']
            sheet_name   = case['sheet_name']
            case_start   = case['start_year']
            case_end     = case['end_year']

            agg_ws.cell(row=current_row, column=2, value=display_name)

            for i, year in enumerate(year_axis):
                if case_start <= year <= case_end:
                    src_col        = 4 + (year - case_start)
                    src_col_letter = get_column_letter(src_col)
                    safe_name      = sheet_name.replace("'", "''")
                    agg_ws.cell(row=current_row, column=4 + i,
                                value=f"=ABS('{safe_name}'!{src_col_letter}{source_row})")
                # else: 留空

            current_row += 1

        case_end_row = current_row - 1

        # 年總計列（填入 SUM 公式並套上對應顏色）
        agg_ws.cell(row=current_row, column=3, value=f"{item_name}/年總計")
        for i in range(n_years):
            col        = 4 + i
            col_letter = get_column_letter(col)
            agg_ws.cell(row=current_row, column=col,
                        value=f"=SUM({col_letter}{case_start_row}:{col_letter}{case_end_row})")
        fill_row(agg_ws, current_row, 3, last_col, total_color)
        current_row += 1

    print(f"[彙整] 已建立「{SHEET_NAME}」，年份 {overall_start}~{overall_end}，"
          f"{len(ITEMS)} 個項目，{len(valid_cases)} 個案場")
    return current_row  # 回傳第一個空白列（供下方附加使用）


def create_income_cashflow_section(target_wb, cases_info, agg_last_row):
    """
    在彙整總表工作表的 agg_last_row 下方（空 6 列）附加：
      - 綜合損益表：範本 B37:W80，資料欄從 D(col4) 開始
      - 現金流量表：範本 B82:W109，資料欄從 E(col5) 開始
    完整複製範本的值（含公式）和樣式，不做任何清除或覆寫。
    超過 20 年時，以指定欄位為樣板向右擴充欄位樣式。
    """
    import os, openpyxl
    from copy import copy
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell

    # ── 範本路徑 ──
    current_dir   = os.path.dirname(os.path.abspath(__file__))
    parent_dir    = os.path.dirname(current_dir)
    template_path = os.path.join(parent_dir, 'Excel final', '滾算後記錄.xlsx')
    if not os.path.exists(template_path):
        print(f"[損益/現金流] 找不到範本: {template_path}")
        return

    valid_cases = [c for c in cases_info if c.get('start_year') and c.get('end_year')]
    if not valid_cases:
        return

    overall_start = min(c['start_year'] for c in valid_cases)
    overall_end   = max(c['end_year']   for c in valid_cases)
    n_years       = overall_end - overall_start + 1

    template_wb = openpyxl.load_workbook(template_path)
    template_ws = template_wb['滾算紀錄1']
    agg_ws = target_wb["彙整總表"]

    print(f"[損益/現金流] 使用工作表: {template_ws.title}, agg_last_row={agg_last_row}, income_start={agg_last_row + 6}")

    def copy_cell(src, dst):
        """完整複製一個儲存格的值（含公式）與樣式；MergedCell 子格略過"""
        if isinstance(src, MergedCell):
            return
        dst.value = src.value
        if src.has_style:
            try:
                dst.font          = copy(src.font)
                dst.fill          = copy(src.fill)
                dst.border        = copy(src.border)
                dst.alignment     = copy(src.alignment)
                dst.number_format = src.number_format
            except Exception as e:
                print(f"[損益/現金流] 樣式複製警告 ({src.coordinate}): {e}")

    def copy_style_only(src, dst):
        """只複製樣式（用於擴充欄位）"""
        if isinstance(src, MergedCell):
            return
        if src.has_style:
            try:
                dst.font          = copy(src.font)
                dst.fill          = copy(src.fill)
                dst.border        = copy(src.border)
                dst.alignment     = copy(src.alignment)
                dst.number_format = src.number_format
            except Exception as e:
                print(f"[損益/現金流] 擴充樣式警告: {e}")

    def copy_block(t_row_start, t_row_end, t_col_start, t_col_end,
                   data_col_start, dst_row_start, ext_template_col):
        """
        將範本 [t_row_start:t_row_end] × [t_col_start:t_col_end] 完整複製到目標。
        若 n_years 超過範本的 20 年，以 ext_template_col 欄為樣板向右擴充樣式。
        """
        template_data_cols = t_col_end - data_col_start + 1  # 範本資料欄數 = 20
        row_off = dst_row_start - t_row_start

        # 1. 逐格複製值 + 樣式
        for r in range(t_row_start, t_row_end + 1):
            for c in range(t_col_start, t_col_end + 1):
                copy_cell(
                    template_ws.cell(row=r, column=c),
                    agg_ws.cell(row=r + row_off, column=c)
                )

        # 2. 複製合併儲存格（只取落在此範圍內的）
        for merge in list(template_ws.merged_cells.ranges):
            if (merge.min_row >= t_row_start and merge.max_row <= t_row_end and
                    merge.min_col >= t_col_start and merge.max_col <= t_col_end):
                agg_ws.merge_cells(
                    start_row=merge.min_row + row_off, start_column=merge.min_col,
                    end_row=merge.max_row   + row_off, end_column=merge.max_col
                )

        # 3. 複製列高
        for r in range(t_row_start, t_row_end + 1):
            rd = template_ws.row_dimensions.get(r)
            if rd and rd.height:
                agg_ws.row_dimensions[r + row_off].height = rd.height

        # 4. 複製欄寬（只做一次，對目標 sheet 整欄設定）
        for c in range(t_col_start, t_col_end + 1):
            col_letter = get_column_letter(c)
            cd = template_ws.column_dimensions.get(col_letter)
            if cd and cd.width:
                agg_ws.column_dimensions[col_letter].width = cd.width

        # 5. 年份超過 20 時，向右擴充欄位樣式
        extra_cols = max(0, n_years - template_data_cols)
        for extra in range(extra_cols):
            new_col = t_col_end + 1 + extra
            for r in range(t_row_start, t_row_end + 1):
                copy_style_only(
                    template_ws.cell(row=r, column=ext_template_col),
                    agg_ws.cell(row=r + row_off, column=new_col)
                )

    year_axis = list(range(overall_start, overall_end + 1))

    # ── 綜合損益表（B37:W80，資料從 D=col4，延伸樣板 D 欄）──
    income_start = agg_last_row + 6
    copy_block(37, 80, 2, 23, 4, income_start, ext_template_col=4)
    income_end = income_start + (80 - 37)

    # ── 現金流量表（B82:W109，資料從 E=col5，延伸樣板 E 欄）──
    cashflow_start = income_end + 1
    copy_block(82, 109, 2, 23, 5, cashflow_start, ext_template_col=5)
    cashflow_end = cashflow_start + (109 - 82)

    # ── 更新年份列（template 第 37、60 列 → D欄起；第 82 列 → E欄起）──
    # 綜合損益表 第一年份列（template row 37）
    dst_row_37 = income_start + (37 - 37)   # = income_start
    for i, year in enumerate(year_axis):
        agg_ws.cell(row=dst_row_37, column=4 + i, value=year)

    # 綜合損益表 第二年份列（template row 60）
    dst_row_60 = income_start + (60 - 37)
    for i, year in enumerate(year_axis):
        agg_ws.cell(row=dst_row_60, column=4 + i, value=year)

    # 現金流量表 年份列（template row 82，年份從 D 欄起）
    dst_row_82 = cashflow_start + (82 - 82)  # = cashflow_start
    for i, year in enumerate(year_axis):
        agg_ws.cell(row=dst_row_82, column=4 + i, value=year)

    # ── 延伸欄位補填序號（template 第 39、62、84 列，序號 1~20 已在模板，21起補填）──
    TEMPLATE_DATA_END_COL = 23   # W = 第 20 年最後一欄
    TEMPLATE_SEQ_END      = 20   # 模板序號最大值
    extra_seq_cols = max(0, n_years - TEMPLATE_SEQ_END)
    for extra in range(extra_seq_cols):
        new_col  = TEMPLATE_DATA_END_COL + 1 + extra  # X, Y, Z...
        seq_val  = TEMPLATE_SEQ_END + 1 + extra        # 21, 22, 23...
        # 綜合損益表 row 39、62
        for t_row, base_t in [(39, 37), (62, 37)]:
            agg_ws.cell(row=income_start + (t_row - base_t), column=new_col, value=seq_val)
        # 現金流量表 row 84
        agg_ws.cell(row=cashflow_start + (84 - 82), column=new_col, value=seq_val)

    template_wb.close()
    print(f"[損益/現金流] 綜合損益表 row {income_start}~{income_end}，"
          f"現金流量表 row {cashflow_start}~{cashflow_end}，"
          f"年份已更新 {overall_start}~{overall_end}")


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
            # 創建新檔案，並將預設空白 Sheet 直接命名為「彙整總表」
            wb = openpyxl.Workbook()
            wb.active.title = "彙整總表"
            wb.save(target_path)
            wb.close()
            print(f"[import_sheets] 創建新檔案: {target_path}")

        # 開啟目標檔案
        target_wb = openpyxl.load_workbook(target_path)
        print(f"[import_sheets] 開啟目標檔案: {target_path}")

        imported_count = 0
        cases_info = []  # 收集每個案場的年份資訊，供彙整總表使用

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

                # 只複製第 1~81 列的儲存格內容、公式和樣式
                COPY_MAX_ROW = 81
                for row in source_sheet.iter_rows(min_row=1, max_row=COPY_MAX_ROW):
                    for cell in row:
                        new_cell = target_sheet.cell(row=cell.row, column=cell.column)

                        # 複製值或公式
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

                # 複製列高（只複製 1~81 列）
                for row_num, row_dim in source_sheet.row_dimensions.items():
                    if row_num <= COPY_MAX_ROW:
                        target_sheet.row_dimensions[row_num].height = row_dim.height
                        target_sheet.row_dimensions[row_num].hidden = row_dim.hidden

                # 複製合併儲存格（只複製範圍在 1~81 列內的）
                for merged_range in source_sheet.merged_cells.ranges:
                    if merged_range.min_row >= 1 and merged_range.max_row <= COPY_MAX_ROW:
                        target_sheet.merge_cells(str(merged_range))

                # 將利息費用 D70:W70 設為 0
                for col in range(4, 24):  # D=4, W=23
                    target_sheet.cell(row=70, column=col).value = 0

                # 將所得稅 D54:W54 設為引用同欄第 77 列的公式
                for col in range(4, 24):  # D=4, W=23
                    col_letter = get_column_letter(col)
                    target_sheet.cell(row=54, column=col).value = f"={col_letter}77"

                # 讀取來源工作表的起始/結束年度，供彙整總表使用
                # 滾算紀錄工作表結構：C5=起始年度數字, E5=結束年度數字
                # 用 data_only=True 開啟，確保取得計算後的實際數值而非公式字串
                try:
                    _wb_data = openpyxl.load_workbook(source_path, data_only=True)
                    _ws_data = _wb_data[source_sheet_name]
                    start_year = int(_ws_data['C5'].value)
                    end_year   = int(_ws_data['E5'].value)
                    _wb_data.close()
                    cases_info.append({
                        'display_name': source_case_name,
                        'sheet_name':   new_sheet_name,
                        'start_year':   start_year,
                        'end_year':     end_year,
                    })
                    print(f"[import_sheets] 年份讀取成功: {source_case_name} {start_year}~{end_year}")
                except (TypeError, ValueError) as e:
                    print(f"[import_sheets] 無法讀取 {new_sheet_name} 的年份資料（C5/E5）: {e}")

                source_wb.close()
                imported_count += 1
                print(f"[import_sheets] 成功匯入: {source_case_name}/{source_sheet_name} -> {new_sheet_name}")

            except Exception as e:
                print(f"[import_sheets] 匯入 {source_case_name}/{source_sheet_name} 失敗: {e}")
                continue

        # 建立彙整總表（自動觸發）
        print(f"[import_sheets] cases_info 共 {len(cases_info)} 筆: {[c['display_name'] for c in cases_info]}")
        if imported_count > 0 and cases_info:
            agg_last_row = create_aggregation_sheet(target_wb, cases_info)
            try:
                create_income_cashflow_section(target_wb, cases_info, agg_last_row)
            except Exception as e:
                import traceback
                print(f"[損益/現金流] 建立失敗（不影響主流程）: {e}")
                print(traceback.format_exc())

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