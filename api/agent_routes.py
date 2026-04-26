"""
AI Agent 藍圖 - 處理聊天請求
"""
from flask import Blueprint, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import sys
import os
import re

# 調整匯入路徑
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from utils.app_logger import log_action, log_error


def _get_excel_dir():
    """根據登入用戶回傳用戶專屬的 Excel 目錄（以 email 命名），若未登入則使用根目錄"""
    base = os.path.join(parent_dir, 'Excel User Data')
    user_email = session.get('user_email')
    d = os.path.join(base, user_email) if user_email else base
    os.makedirs(d, exist_ok=True)
    return d


def _find_excel_file(case_id='', case_name='', original_filename=''):
    """
    通用 Excel 檔案定址，優先順序：
    1. case_id + original_filename  → {case_id}_{original_filename}（最精準）
    2. case_name + original_filename → {case_name}_{original_filename}（相容舊檔）
    3. case_id 前綴搜尋              → {case_id}_*.xlsx
    4. case_name 前綴搜尋            → {case_name}_*.xlsx
    5. original_filename 包含搜尋
    6. 目錄內任意 xlsx               → 最後保底
    """
    excel_dir = _get_excel_dir()
    all_files = [f for f in os.listdir(excel_dir)
                 if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')] if os.path.exists(excel_dir) else []

    # 方法1: case_id + original_filename
    if case_id and original_filename:
        candidate = os.path.join(excel_dir, f"{case_id}_{original_filename}")
        if os.path.exists(candidate):
            return candidate

    # 方法2: case_name + original_filename
    if case_name and original_filename:
        candidate = os.path.join(excel_dir, f"{case_name}_{original_filename}")
        if os.path.exists(candidate):
            return candidate

    # 方法3: case_id 前綴
    if case_id:
        matches = [f for f in all_files if f.startswith(f"{case_id}_")]
        if matches:
            return os.path.join(excel_dir, matches[0])

    # 方法4: case_name 前綴
    if case_name:
        matches = [f for f in all_files if f.startswith(f"{case_name}_")]
        if matches:
            return os.path.join(excel_dir, matches[0])

    # 方法5: original_filename 包含搜尋
    if original_filename:
        matches = [f for f in all_files if original_filename in f]
        if matches:
            return os.path.join(excel_dir, matches[0])

    # 方法6: 保底，取第一個 xlsx
    if all_files:
        return os.path.join(excel_dir, all_files[0])

    return None

try:
    from core.agent import AIAgent
    from config.settings import DEFAULT_CONFIG
except ImportError as e:
    print(f"嚴重警告: 無法匯入 AI Agent 核心模組: {e}")
    AIAgent = None
    DEFAULT_CONFIG = None

agent_bp = Blueprint('agent_bp', __name__)

# ── 使用說明文字 ──────────────────────────────────────────────
_HELP_TEXT = """## 財模助手 使用說明

歡迎使用財模助手！以下是主要功能的操作方式。

---

### 一、建立案場與上傳 Excel
1. 點擊左側「＋ 新增案場」建立案場
2. 在案場中點擊「上傳 Excel」，選擇財務模型檔案（.xlsx）
3. 上傳後系統會自動載入試算表

---

### 二、查詢財務數據
直接用自然語言詢問，例如：
- `p1 的專案法 IRR 是多少？`
- `查詢 p2 的 2025 年現金流`
- `比較所有情境的 IRR`

---

### 三、修改 Excel 數值
直接說明要改什麼，例如：
- `把 p3 的保險費 2025~2030 改成 -50000`
- `修改第 48 行 2026 年改成 -40000`
- `p2 公版的設備費用改成 28000`

系統會顯示確認訊息，回覆「**是**」後執行修改。

---

### 四、價金滾算
點擊畫面上方的「**滾算**」按鈕，選擇模式並填入參數：

| 模式 | 說明 |
|------|------|
| 現金模式 | 以固定金額逐步遞減 |
| 比率模式 | 以固定比率逐步遞減 |
| 自訂模式 | 手動輸入每次遞減金額 |

計算完成後可選擇是否存入 Excel。

---

### 五、多站彙整
在「匯入表格」功能中，選取多個案場的工作表合併，系統自動產生「多站彙整」總表。

---

### 支援的關鍵字
輸入以下任一關鍵字可重新顯示此說明：
`使用說明`、`說明`、`help`、`如何使用`、`怎麼用`
"""

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




# ========== 待確認保存的滾算結果快取 ==========
# 用於存儲等待用戶確認是否保存的滾算結果
# key: case_id（優先）或 case_name，value: { 'mode': ..., 'params': {...}, 'excel_path': ..., 'timestamp': ... }
_pending_rolling_save = {}


def _store_pending_rolling_save(case_name: str, mode: str, params: dict, excel_path: str, case_id: str = '', tool_result: dict = None):
    """存儲待確認保存的滾算結果"""
    from datetime import datetime
    key = case_id if case_id else case_name
    _pending_rolling_save[key] = {
        'mode': mode,
        'params': params,
        'excel_path': excel_path,
        'tool_result': tool_result,  # 存入前端顯示的計算結果，確認時直接使用，不重算
        'timestamp': datetime.now()
    }
    print(f"[快取] 已存儲待確認滾算: key={key}, mode={mode}")


def _get_pending_rolling_save(case_name: str, case_id: str = '') -> dict:
    """獲取待確認保存的滾算結果"""
    from datetime import datetime
    key = case_id if case_id else case_name
    pending = _pending_rolling_save.get(key)
    if pending:
        # 檢查是否過期（5分鐘）
        elapsed = (datetime.now() - pending['timestamp']).total_seconds()
        if elapsed > 300:
            del _pending_rolling_save[key]
            print(f"[快取] 待確認滾算已過期: key={key}")
            return None
        return pending
    return None


def _clear_pending_rolling_save(case_name: str, case_id: str = ''):
    """清除待確認保存的滾算結果"""
    key = case_id if case_id else case_name
    if key in _pending_rolling_save:
        del _pending_rolling_save[key]
        print(f"[快取] 已清除待確認滾算: key={key}")


# ========== 待確認修改快取 ==========
_pending_modification = {}

# 欄位名稱 → section_type 推斷對應表
_FIELD_SECTION_MAP = {
    '綜合損益表': ['保險費', '租金', '運維費', '回收費', '維運費', '管理費', '電費', '折舊', '利息',
                   '稅', '收入', '費用', '修繕', '雜支', '人事', '水費'],
    '公版':       ['設備費用', '開發費', '建置費', '容量', '利潤率', '裝置容量', '每kw', '價金'],
    '現金流量表': ['現金流', '借款', '還款', '貸款', '融資', '資本'],
}


def _infer_section_type(field_keyword: str) -> str:
    """根據欄位名稱推斷 section_type"""
    if not field_keyword:
        return None
    for section, keywords in _FIELD_SECTION_MAP.items():
        for kw in keywords:
            if kw in field_keyword or field_keyword in kw:
                return section
    return None


def _parse_modification(query: str, current_sheet: str = None) -> dict:
    """從 query 解析修改參數，回傳 dict（參數不足時部分欄位為 None）"""
    import re
    result = {
        'sheet_name': current_sheet,
        'field_keyword': None,
        'year_spec': '全部',
        'new_value': None,
        'section_type': None,
        'year_value_map': None,
        'row_hint': None,
    }

    # 工作表名稱（p1/p2/p3 或其他格式）
    sheet_m = re.search(r'(?<![a-zA-Z\d])(p\d+)(?![a-zA-Z\d])', query, re.IGNORECASE)
    if sheet_m:
        result['sheet_name'] = sheet_m.group(1)

    # 行號提取（優先，避免行號被當成欄位名稱）：「第48行」或「48行」（1-3位數，避免誤抓年份）
    row_m = re.search(r'第?\s*(\d{1,3})\s*行', query)
    if row_m:
        result['row_hint'] = int(row_m.group(1))

    # 區域
    if re.search(r'公版', query):
        result['section_type'] = '公版'
    elif re.search(r'綜合損益表|損益表', query):
        result['section_type'] = '綜合損益表'
    elif re.search(r'現金流量表|流量表', query):
        result['section_type'] = '現金流量表'

    # 多年份-數值對
    yv_matches = re.findall(r'(\d{4})\s*(?:改[為成]|設[為成]|→|->)\s*(-?\d+(?:\.\d+)?)', query)
    if len(yv_matches) >= 2:
        result['year_value_map'] = {int(y): float(v) for y, v in yv_matches}
        result['year_spec'] = 'multiple'
    else:
        # 年份範圍或單年
        yr = re.search(r'(\d{4})\s*[~～\-到至]\s*(\d{4})', query)
        if yr:
            result['year_spec'] = f"{yr.group(1)}~{yr.group(2)}"
        else:
            # 支援「年份是2025」、「年分是2025」、「年份=2025」等格式
            sy_labeled = re.search(r'年[份分]?\s*[是為=:：]\s*(\d{4})', query)
            if sy_labeled:
                result['year_spec'] = sy_labeled.group(1)
            else:
                sy = re.search(r'(\d{4})\s*年', query)
                if sy:
                    result['year_spec'] = sy.group(1)
                elif re.search(r'全部|所有|全年', query):
                    result['year_spec'] = '全部'

        # 數值
        vm = re.search(r'(?:改成|改為|設為|設成|變成|變為|→|->)\s*(-?\d+(?:\.\d+)?)', query)
        if vm:
            result['new_value'] = float(vm.group(1))
        else:
            # 末尾獨立數字
            vm2 = re.search(r'(?<!\d)(-?\d{4,})(?!\d)', query)
            if vm2:
                candidate = float(vm2.group(1))
                # 排除年份（1900~2100）
                if not (1900 <= candidate <= 2100):
                    result['new_value'] = candidate

    # 欄位關鍵字：移除干擾詞後，優先取已知欄位關鍵字，否則取最長詞
    fq = query
    for pattern in [r'p\d+', r'公版|綜合損益表|損益表|現金流量表|流量表',
                    r'\d{4}\s*[~～\-到至]\s*\d{4}', r'\d{4}\s*年?',
                    r'全部年份|所有年份|全年份|年份',              # 避免時間描述詞被誤選為欄位
                    r'第?\s*\d{1,3}\s*行',   # 行號（如「第48行」「48行」），避免誤判為欄位名
                    r'(?:改成|改為|設為|設成|變成|變為|→|->)\s*-?\d+(?:\.\d+)?',
                    r'我想|想要|想|要|幫我|請|調整|變更|修改|更改|把|的|將|中|裡|內|改|成|為|變|設定|可以|能|麻煩']:
        fq = re.sub(pattern, ' ', fq, flags=re.IGNORECASE)
    words = [w.strip() for w in re.split(r'[\s，。、！？]+', fq) if len(w.strip()) >= 2]
    if words:
        # 優先選取符合已知欄位清單的詞（避免「全部」「年份」等描述詞搶佔欄位位置）
        all_known_keywords = [kw for kws in _FIELD_SECTION_MAP.values() for kw in kws]
        known_matches = [w for w in words if any(kw in w or w in kw for kw in all_known_keywords)]
        result['field_keyword'] = max(known_matches, key=len) if known_matches else max(words, key=len)

    # 推斷 section_type
    if not result['section_type'] and result['field_keyword']:
        result['section_type'] = _infer_section_type(result['field_keyword'])

    return result


def _set_pending_modification(case_id: str, case_name: str, params: dict):
    from datetime import datetime
    key = case_id if case_id else case_name
    _pending_modification[key] = {**params, 'timestamp': datetime.now()}
    print(f"[修改快取] 已存入: key={key}, params={params}")


def _get_pending_modification(case_id: str, case_name: str) -> dict:
    from datetime import datetime
    key = case_id if case_id else case_name
    pending = _pending_modification.get(key)
    if pending:
        if (datetime.now() - pending['timestamp']).total_seconds() > 300:
            del _pending_modification[key]
            return None
        return pending
    return None


def _clear_pending_modification(case_id: str, case_name: str):
    key = case_id if case_id else case_name
    _pending_modification.pop(key, None)
    print(f"[修改快取] 已清除: key={key}")


# ── 現金流量表參數關鍵字定義 ──
_CASHFLOW_PARAM_KEYWORDS = {
    '貸款成數': '貸款成數(%)',
    '還款期數': '還款期數',
    '股利比率': '股利比率(%)',
    '年底減資': '年底減資攤還期數',
    '減資攤還': '年底減資攤還期數',
    '銀行利率': '銀行利率(%)',
}

def _detect_cashflow_param(query: str):
    """
    偵測 query 是否為修改現金流量表參數的請求。
    回傳 (canonical_name, new_value) 或 (None, None)。
    canonical_name 為 B 欄實際字串，new_value 為浮點數或 None。
    """
    import re
    from difflib import SequenceMatcher

    matched_key = None
    best_score = 0.0

    for kw, canonical in _CASHFLOW_PARAM_KEYWORDS.items():
        # 完全包含
        if kw in query:
            matched_key = canonical
            break
        # 模糊相似度
        score = SequenceMatcher(None, kw, query).ratio()
        if score > best_score:
            best_score = score
            if score >= 0.6:
                matched_key = canonical

    if not matched_key:
        return None, None

    # 提取數值
    vm = re.search(r'(?:改成|改為|設為|設成|變成|變為|→|->|=|：|:)\s*(-?\d+(?:\.\d+)?)', query)
    new_value = float(vm.group(1)) if vm else None
    if new_value is None:
        vm2 = re.search(r'(?<!\d)(\d{1,3})(?!\d)', query)
        if vm2:
            new_value = float(vm2.group(1))

    return matched_key, new_value


def _update_cashflow_params(excel_path: str, param_name: str, new_value: float) -> dict:
    """
    修改「多站彙整」工作表中的現金流量表參數（B 欄搜尋），
    並在修改「還款期數」或「年底減資攤還期數」時重填對應公式列。
    回傳 {'success': bool, 'message': str}
    """
    import openpyxl
    from openpyxl.utils import get_column_letter
    from difflib import SequenceMatcher

    try:
        wb = openpyxl.load_workbook(excel_path)
        if '多站彙整' not in wb.sheetnames:
            return {'success': False, 'message': '找不到「多站彙整」工作表'}

        ws = wb['多站彙整']

        # ── 1. 掃 B 欄找「貸款成數(%)」定位參數區起點 ──
        param_start_row = None
        for row in ws.iter_rows(min_col=2, max_col=2):
            cell = row[0]
            if cell.value and '貸款成數' in str(cell.value):
                param_start_row = cell.row
                break

        if param_start_row is None:
            return {'success': False, 'message': '找不到現金流量表參數區（B 欄缺少「貸款成數」）'}

        # ── 2. 在參數區 5 列中找目標列（模糊比對）──
        target_row = None
        for offset in range(5):
            r = param_start_row + offset
            b_val = str(ws.cell(row=r, column=2).value or '')
            # 完全包含
            if param_name in b_val or b_val in param_name:
                target_row = r
                break
            # 模糊
            if SequenceMatcher(None, param_name, b_val).ratio() >= 0.6:
                target_row = r
                break

        if target_row is None:
            return {'success': False, 'message': f'在參數區找不到「{param_name}」'}

        old_value = ws.cell(row=target_row, column=3).value
        ws.cell(row=target_row, column=3).value = new_value
        print(f"[cashflow_param] {param_name} 第 {target_row} 列: {old_value} → {new_value}")

        # ── 3. 讀 row 4 年份標題，算 n_years ──
        n_years = 0
        for col in range(4, ws.max_column + 1):
            if ws.cell(row=4, column=col).value is not None:
                n_years += 1
            else:
                break
        if n_years == 0:
            wb.save(excel_path)
            wb.close()
            return {'success': True, 'message': f'已更新「{param_name}」為 {new_value}（無法讀取年份，跳過公式重填）'}

        # ── 4. 找出 r93（還款列）和 r96（年底減資列）的實際列號 ──
        # 策略：掃 B 欄找「理財活動」區塊後的「還款」和「年底減資」
        r93_actual = None  # 還款公式列
        r96_actual = None  # 年底減資公式列
        p_repay_cell = f"C{param_start_row + 1}"   # 還款期數的 C 欄位置（固定 offset）
        p_cap_cell   = f"C{param_start_row + 3}"   # 年底減資攤還期數

        # 找 r92（借款列）→ 往下一列就是 r93（還款）
        for row in ws.iter_rows(min_col=2, max_col=2):
            b = str(row[0].value or '')
            if '借款' in b and '理財' in b:
                r93_actual = row[0].row + 1
            if '年底減資' in b:
                r96_actual = row[0].row

        # ── 5. 重填公式 ──
        new_int = int(new_value)

        if '還款' in param_name and r93_actual is not None:
            # 清空 r93 整列（D 欄起到最後年份）
            for col in range(4, 4 + n_years):
                ws.cell(row=r93_actual, column=col).value = None

            # r92 的借款在 D 欄
            r92_actual = r93_actual - 1
            for i in range(min(new_int, n_years - 1)):
                ws.cell(row=r93_actual, column=5 + i,
                        value=f"=-D{r92_actual}/{p_repay_cell}")
            print(f"[cashflow_param] 重填還款公式 r93={r93_actual}，期數={new_int}")

        if '減資' in param_name and r96_actual is not None:
            # 清空 r96 整列
            for col in range(4, 4 + n_years):
                ws.cell(row=r96_actual, column=col).value = None

            # r94（現金增資）在 D 欄
            r94_actual = r96_actual - 2
            for i in range(min(new_int, n_years)):
                col = 3 + n_years - i   # 從最後一欄往前
                ws.cell(row=r96_actual, column=col,
                        value=f"=-D{r94_actual}/{p_cap_cell}")
            print(f"[cashflow_param] 重填年底減資公式 r96={r96_actual}，期數={new_int}")

        wb.save(excel_path)
        wb.close()

        # LibreOffice recalc
        try:
            from utils.recalc import recalc as _recalc
            _recalc(excel_path, timeout=90)
        except Exception:
            pass

        return {'success': True, 'message': f'已將「{param_name}」從 {old_value} 更新為 {new_value}'}

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return {'success': False, 'message': f'更新失敗：{e}'}


def _parse_edit_request_REMOVED(query: str, existing_params: dict = None) -> dict:  # 已廢棄，保留以備參考
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
        'year_value_map': None,  # 年份-數值映射，用於非連續年份的修改需求
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
    case_id = str(data.get('case_id', ''))
    original_filename = data.get('original_filename', '')
    sheet_name = data.get('sheet_name', None)  # 工作表名稱，可選

    if not user_query:
        return jsonify({"error": "請提供 'query' 參數"}), 400

    # ── 千分位逗號正規化：-40,000 → -40000（避免 regex 解析與 LLM tool call 錯誤）──
    user_query = re.sub(
        r'(?<!\d)(-?\d{1,3}(?:,\d{3})+)(?!\d)',
        lambda m: m.group(0).replace(',', ''),
        user_query
    )

    # ── 使用說明快捷回應（不走 AI，直接回傳固定說明）──
    _HELP_KEYWORDS = {'使用說明', '說明', 'help', '幫助', '如何使用', '怎麼用', '功能介紹', '教學'}
    if user_query.strip().lower() in _HELP_KEYWORDS or user_query.strip() in _HELP_KEYWORDS:
        return jsonify({
            "query": user_query,
            "response": _HELP_TEXT,
            "excel_modified": False
        })

    print(f"\n--- API 請求: 執行 Agent 對話 ---")
    print(f"收到查詢: {user_query}")
    print(f"案場資訊: case_id={case_id}, case_name={case_name}, original_filename={original_filename}, sheet_name={sheet_name}")

    user_email = session.get('user_email', 'anonymous')
    log_action(user_email, 'agent_chat',
               f"case={case_name}({case_id}) query={user_query[:200]}")

    try:
        # 獲取 Agent 實例
        agent = get_agent()

        # 使用通用定址 helper 取得 Excel 檔案路徑
        excel_path = _find_excel_file(case_id=case_id, case_name=case_name, original_filename=original_filename)
        if excel_path:
            agent.tool_manager.set_finance_excel_file(excel_path, sheet_name)
            print(f"已設定財務工具 Excel 檔案: {excel_path}")
        else:
            print(f"警告: 找不到 Excel 檔案 (case_id={case_id}, case_name={case_name})")

        # ── 取得可用工作表清單，注入到 query，避免模型幻覺 ──
        sheet_names_prefix = ''
        current_excel_path = excel_path
        if current_excel_path and os.path.exists(current_excel_path):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(current_excel_path, read_only=True)
                sheet_names = wb.sheetnames
                wb.close()
                sheet_names_prefix = f"[目前 Excel 可用工作表: {', '.join(sheet_names)}]\n"
            except Exception:
                pass

        # ── 0. 待確認保存的滾算結果 ──
        pending_rolling = _get_pending_rolling_save(case_name, case_id=case_id) if (case_name or case_id) else None

        if pending_rolling:
            is_confirm = re.search(r'^(是|好|要|對|確認|保存|儲存|存|yes|ok|save)$',
                                   user_query.strip(), re.IGNORECASE)
            is_decline = re.search(r'^(否|不|不要|取消|no|cancel)$',
                                   user_query.strip(), re.IGNORECASE)

            if is_confirm:
                print("用戶確認保存滾算結果...")
                _clear_pending_rolling_save(case_name, case_id=case_id)
                save_params = pending_rolling['params'].copy()
                save_params['excel_file'] = pending_rolling.get('excel_path')
                mode_mapping = {
                    "CashMode": "cash", "RatioMode": "ratio",
                    "CustomizeMode": "customize"
                }
                save_params['mode'] = mode_mapping.get(
                    save_params.get('mode', ''), save_params.get('mode', '').lower())

                # 將前端已計算好的 IRR 結果傳入，避免重算導致數值不一致
                stored_tool_result = pending_rolling.get('tool_result')
                # calculate_price_rolling 結果結構: {"results_summary": {"data": data_rows}}
                data_rows = (stored_tool_result or {}).get('results_summary', {}).get('data')
                if data_rows:
                    # data_rows 欄位順序: [price_per_kw, profit_per_kw, final_price_per_kw, project_irr, cost_method_irr, equity_method_irr]
                    save_params['precomputed_irr_results'] = [
                        {
                            'project_irr':       row[3],
                            'cost_method_irr':   row[4],
                            'equity_method_irr': row[5],
                            'profit_per_kw':     row[1],  # 直接用前端已算好的每kw信邦利潤
                        }
                        for row in data_rows
                    ]
                    print(f"[快取] 使用預計算 IRR，共 {len(data_rows)} 筆")

                    # CustomizeMode：從預計算價格反推實際 steps，避免儲存時重新隨機生成
                    # 問題根源：execute_price_rolling 的參數是 adjustment_times（預設10），
                    # 而 save_params 帶的是 adjust_times，導致重算出 11 筆且數字錯誤
                    if save_params.get('mode') == 'customize':
                        prices = [row[0] for row in data_rows]
                        computed_steps = [prices[i] - prices[i + 1] for i in range(len(prices) - 1)]
                        save_params['custom_steps'] = computed_steps
                        save_params['auto_config'] = False
                        save_params['adjustment_times'] = len(computed_steps)
                        print(f"[CustomizeMode] 反推 steps：{computed_steps}")
                else:
                    print(f"[快取] 無預計算 IRR，將重新計算（stored_tool_result keys: {list((stored_tool_result or {}).keys())})")

                save_result = agent.tool_manager.execute_tool("execute_price_rolling", save_params)
                if save_result.get("success"):
                    return jsonify({"query": user_query, "response": "滾算紀錄已儲存",
                                    "excel_modified": True})
                else:
                    return jsonify({"query": user_query,
                                    "response": f"儲存失敗：{save_result.get('message', '未知錯誤')}",
                                    "excel_modified": False})

            elif is_decline:
                print("用戶拒絕保存滾算結果")
                _clear_pending_rolling_save(case_name, case_id=case_id)
                return jsonify({"query": user_query,
                                "response": "好的，滾算結果不會儲存。您可以繼續進行其他操作。",
                                "excel_modified": False})
            else:
                # 其他輸入視為不保存，繼續處理新請求
                _clear_pending_rolling_save(case_name, case_id=case_id)

        # ── 0b. 待確認修改快取 ──
        pending_mod = _get_pending_modification(case_id, case_name)
        if pending_mod:
            # 取消語境偵測
            _cancel_keywords = ['取消修改', '算了，不改了', '不做調整', '取消本次調整', '不弄了']
            _is_cancel = (
                user_query.strip() in ['取消', '算了'] or
                any(kw in user_query for kw in _cancel_keywords)
            )
            if _is_cancel:
                _clear_pending_modification(case_id, case_name)
                return jsonify({"query": user_query,
                                "response": "收到，已為你取消本次調整，如需任何協助，請告訴我。",
                                "excel_modified": False})

            is_confirm = re.search(r'^(是|好|要|對|y|yes|確認|ok)$', user_query.strip(), re.IGNORECASE)
            is_decline = re.search(r'^(否|不|不要|n|no|cancel)$', user_query.strip(), re.IGNORECASE)

            if is_confirm:
                _clear_pending_modification(case_id, case_name)

                # 現金流量表參數修改走專屬函式
                if pending_mod.get('is_cashflow_param'):
                    cf_excel = pending_mod.get('excel_path') or current_excel_path
                    cf_result = _update_cashflow_params(
                        cf_excel,
                        pending_mod['field_keyword'],
                        pending_mod['new_value']
                    )
                    msg = f"調整已完成。\n（{cf_result['message']}）" if cf_result['success'] else f"修改失敗：{cf_result['message']}"
                    return jsonify({"query": user_query, "response": msg, "excel_modified": cf_result['success']})

                exec_params = {k: v for k, v in pending_mod.items() if k != 'timestamp'}
                exec_result = agent.tool_manager.execute_tool("edit_sheet_by_field", exec_params)
                if exec_result.get("success"):
                    msg = f"調整已完成，請查收。\n（{exec_result.get('message', '')}）"
                    return jsonify({"query": user_query, "response": msg, "excel_modified": True})
                else:
                    return jsonify({"query": user_query,
                                    "response": f"修改失敗：{exec_result.get('message', '未知錯誤')}",
                                    "excel_modified": False})

            elif is_decline:
                _clear_pending_modification(case_id, case_name)
                clarify_query = (
                    f"{sheet_names_prefix}"
                    f"[使用者拒絕了修改確認，請詢問他想調整哪些地方。格式：請告訴我哪裡需要調整（分頁/區域/項目/金額/年份）]\n"
                    f"{user_query}"
                )
                agent_response = agent.chat(clarify_query)
                return jsonify({"query": user_query, "response": agent_response, "excel_modified": False})

            else:
                # 非確認/拒絕 → 嘗試合併補充資料
                _pending_has_value = (
                    pending_mod.get('new_value') is not None or
                    pending_mod.get('year_value_map') is not None
                )
                supplement = _parse_modification(user_query, current_sheet=pending_mod.get('sheet_name'))
                _supp_has_value = (
                    supplement.get('new_value') is not None or
                    supplement.get('year_value_map') is not None
                )
                _supp_new_field = (
                    supplement.get('field_keyword') and
                    supplement.get('field_keyword') != pending_mod.get('field_keyword')
                )
                _supp_year = supplement.get('year_spec') and supplement['year_spec'] != '全部'

                # 判斷是否為補充（數值補充 or 年份修正），而非全新請求
                _is_new_modify_request = bool(
                    re.search(r'修改|更改|改成|把.+改|變更', user_query) and _supp_new_field
                )

                if not _is_new_modify_request and (_supp_has_value or _supp_year):
                    # 合併補充資料
                    merged = {k: v for k, v in pending_mod.items() if k != 'timestamp'}
                    if supplement.get('new_value') is not None:
                        merged['new_value'] = supplement['new_value']
                    if supplement.get('year_value_map') is not None:
                        merged['year_value_map'] = supplement['year_value_map']
                    if _supp_year:
                        merged['year_spec'] = supplement['year_spec']
                    # 合併後若資料完整，直接執行
                    _merged_has_value = (
                        merged.get('new_value') is not None or
                        merged.get('year_value_map') is not None
                    )
                    if bool(merged.get('field_keyword')) and _merged_has_value:
                        _clear_pending_modification(case_id, case_name)
                        exec_result = agent.tool_manager.execute_tool("edit_sheet_by_field", merged)
                        if exec_result.get("success"):
                            msg = f"調整已完成，請查收。\n（{exec_result.get('message', '')}）"
                            return jsonify({"query": user_query, "response": msg, "excel_modified": True})
                        else:
                            return jsonify({"query": user_query,
                                            "response": f"修改失敗：{exec_result.get('message', '未知錯誤')}",
                                            "excel_modified": False})
                    else:
                        # 更新快取，繼續等待補充
                        _set_pending_modification(case_id, case_name, merged)
                        _clear_pending_modification(case_id, case_name)
                else:
                    # 視為新請求，清除快取後繼續往下處理
                    _clear_pending_modification(case_id, case_name)

        # ── 0c. 現金流量表參數修改（多站彙整專用）──
        _is_edit_intent_cf = re.search(r'修改|更改|改成|設定|編輯|把.+改|變更|調整', user_query)
        if _is_edit_intent_cf:
            _is_agg_context = (sheet_name == '多站彙整') or ('多站彙整' in user_query)
            cf_param, cf_value = _detect_cashflow_param(user_query)
            if cf_param:
                # 驗證：必須在多站彙整工作表，或 query 中明確指定「多站彙整」
                is_agg_sheet = _is_agg_context
                if not is_agg_sheet:
                    agent_response = f'「{cf_param}」參數只存在於「多站彙整」工作表，請先切換到該分頁再操作。'
                    return jsonify({"query": user_query, "response": agent_response, "excel_modified": False})

                if cf_value is None:
                    # 缺少數值，先詢問
                    agent_response = f'已找到「{cf_param}」參數，請問要改為多少？'
                    _set_pending_modification(case_id, case_name, {
                        'sheet_name': '多站彙整',
                        'field_keyword': cf_param,
                        'year_spec': '全部',
                        'new_value': None,
                        'section_type': '現金流量表參數',
                        'year_value_map': None,
                        'row_hint': None,
                        'is_cashflow_param': True,
                    })
                    return jsonify({"query": user_query, "response": agent_response, "excel_modified": False})

                # 有完整資訊，先確認再執行
                _set_pending_modification(case_id, case_name, {
                    'sheet_name': '多站彙整',
                    'field_keyword': cf_param,
                    'year_spec': '全部',
                    'new_value': cf_value,
                    'section_type': '現金流量表參數',
                    'year_value_map': None,
                    'row_hint': None,
                    'is_cashflow_param': True,
                    'excel_path': current_excel_path,
                })
                agent_response = (
                    f'請問這是你要調整的方式嗎：\n'
                    f'1. 調整目標分頁：多站彙整\n'
                    f'2. 調整項目：{cf_param}\n'
                    f'3. 調整後數值：{cf_value}\n'
                    f'目標資料已經鎖定，確定要執行修改嗎？(y / n)'
                )
                return jsonify({"query": user_query, "response": agent_response, "excel_modified": False})

            elif _is_agg_context:
                # 在多站彙整但找不到對應參數 → 提示使用者，不往下走
                _valid_params = '、'.join(_CASHFLOW_PARAM_KEYWORDS.keys())
                agent_response = (
                    f'找不到你指定的參數，可能是字打錯了。\n\n'
                    f'多站彙整可修改的參數為：\n'
                    f'**{_valid_params}**\n\n'
                    f'請確認後重新輸入。'
                )
                return jsonify({"query": user_query, "response": agent_response, "excel_modified": False})

        # ── 1. 價金滾算請求 → llm_service ──
        # 條件：明確的滾算動作關鍵字，且不包含修改/編輯意圖
        # 「修改滾算紀錄1」中的「滾算紀錄」只是 sheet 名稱，不應進入此路由
        _is_edit_intent = re.search(r'修改|更改|改成|設定|編輯|把.+改|變更|調整', user_query)
        _is_rolling_action = re.search(
            r'執行滾算紀錄|儲存滾算|保存滾算'          # 明確存檔
            r'|價金\s*滾算'                             # 價金滾算
            r'|price\s*rolling'                         # 英文
            r'|設備成本'                                # 設備成本滾算
            r'|cashmode|ratiomode|customizemode',  # 模式關鍵字
            user_query, re.IGNORECASE
        )

        if _is_rolling_action and not _is_edit_intent:
            print("檢測到滾算相關請求，使用 llm_service 處理...")
            from services.llm_service import process_user_query
            agent_response, excel_modified = process_user_query(
                user_query, excel_path=current_excel_path, sheet_name=sheet_name)

        # ── 2. 所有其他請求（含 IRR 查詢、修改 Excel）→ LLM function calling ──
        else:
            print("交由 LLM function calling 處理...")
            _is_modify = re.search(r'修改|更改|改成|設定|把.+改|調整|變更', user_query)

            if _is_modify and not _is_rolling_action:
                # 路由層解析修改參數
                parsed = _parse_modification(user_query, current_sheet=sheet_name)

                _has_field = bool(parsed.get('field_keyword')) or parsed.get('row_hint') is not None
                _has_value = (
                    parsed.get('new_value') is not None or
                    parsed.get('year_value_map') is not None
                )

                if not _has_field:
                    # 找不到修改目標 → 告知使用者
                    inject = (
                        "[系統無法識別修改目標，請用以下格式原文回覆（禁止呼叫工具）：\n"
                        "對不起，我找不到你要調的資料項目，可能原因如下：\n"
                        "1. 資料不存在當前分頁。\n"
                        "2. 設定資訊嚴重不完整。\n"
                        "3. 要調的內容不符合表格結構。\n"
                        "請確認原因後重新送出回覆，我會盡力為你提供協助。]\n"
                    )
                    enhanced_query = f"{sheet_names_prefix}{inject}{user_query}"
                    agent_response = agent.chat(enhanced_query)
                    excel_modified = False

                elif not _has_value:
                    # 找到欄位但缺少數值 → 存入快取，詢問數值
                    _set_pending_modification(case_id, case_name, parsed)
                    _sheet_disp = parsed.get('sheet_name') or sheet_name or '目前分頁'
                    _field_disp = parsed.get('field_keyword')
                    inject = (
                        f"[系統已找到修改目標，但缺少修改數值：\n"
                        f"分頁={_sheet_disp}，項目={_field_disp}\n"
                        f"請用繁體中文詢問使用者要將「{_field_disp}」改為多少，禁止呼叫工具]\n"
                    )
                    enhanced_query = f"{sheet_names_prefix}{inject}{user_query}"
                    agent_response = agent.chat(enhanced_query)
                    excel_modified = False

                else:
                    # 資訊完整 → 存入快取，直接用 Python 產生格式固定的確認訊息（不走 LLM，確保格式穩定）
                    _set_pending_modification(case_id, case_name, parsed)
                    year_desc = (
                        "、".join([f"{y}年={v}" for y, v in sorted(parsed['year_value_map'].items())])
                        if parsed.get('year_value_map')
                        else parsed.get('year_spec', '全部')
                    )
                    _location_desc = (
                        f"第{parsed['row_hint']}行"
                        if parsed.get('row_hint') and not parsed.get('field_keyword')
                        else (parsed.get('field_keyword') or '未指定')
                    )
                    _sheet_disp   = parsed.get('sheet_name') or '未指定'
                    _section_disp = parsed.get('section_type') or '待確認'
                    _value_disp   = (
                        year_desc
                        if parsed.get('year_value_map')
                        else str(parsed.get('new_value'))
                    )
                    _row_line = (
                        f"\n行號：第{parsed['row_hint']}行"
                        if parsed.get('row_hint') else ""
                    )
                    agent_response = (
                        f"根據系統已解析的修改參數，Double Check 確認訊息如下：\n\n"
                        f"分頁：{_sheet_disp}\n"
                        f"區域：{_section_disp}\n"
                        f"項目：{_location_desc}{_row_line}\n"
                        f"修改值：{_value_disp}\n"
                        f"年份範圍：{year_desc}\n\n"
                        f"請確認是否要將「{_sheet_disp}」工作表中 {year_desc} 的「{_location_desc}」，"
                        f"統一修改為 {_value_disp}？（y / n）"
                    )
                    excel_modified = False
            else:
                current_sheet_info = f"[使用者目前觀看的工作表：{sheet_name}]\n" if sheet_name else ""
                enhanced_query = f"{sheet_names_prefix}{current_sheet_info}{user_query}"
                agent_response = agent.chat(enhanced_query)
                excel_tools = ['write_excel_cell', 'delete_excel_cell', 'read_excel_cell',
                               'edit_sheet_by_field']
                excel_modified = any(tool in agent.last_used_tools for tool in excel_tools)

        print(f"Agent 回應成功 (Excel modified: {excel_modified})")
        log_action(user_email, 'agent_chat_done',
                   f"case={case_name}({case_id}) excel_modified={excel_modified} "
                   f"tools_used={agent.last_used_tools} response_len={len(agent_response)}")
        return jsonify({
            "query": user_query,
            "response": agent_response,
            "excel_modified": excel_modified
        })

    except Exception as e:
        import traceback
        print(f"AI Agent 處理錯誤: {e}")
        print(traceback.format_exc())
        log_error(user_email, 'agent_chat',
                  e, f"case={case_name}({case_id}) query={user_query[:200]}")
        return jsonify({"error": "系統處理時發生問題，請稍後再試。"}), 500

AVAILABLE_MODELS = ['qwen3:4b', 'qwen3:14b', 'qwen3:32b']

@agent_bp.route('/agent/model', methods=['GET'])
def get_model():
    """取得目前使用的模型"""
    try:
        agent = get_agent()
        return jsonify({"status": "success", "model": agent.config.model_name})
    except Exception as e:
        log_error(session.get('user_email', 'anonymous'), 'get_model', e)
        return jsonify({"status": "error", "error": str(e)}), 500

@agent_bp.route('/agent/model', methods=['POST'])
def set_model():
    """切換模型"""
    try:
        data = request.get_json()
        model = data.get('model', '')
        if model not in AVAILABLE_MODELS:
            return jsonify({"status": "error", "error": f"不支援的模型，可選: {AVAILABLE_MODELS}"}), 400
        agent = get_agent()
        agent.config.model_name = model
        agent.connection.model_name = model
        log_action(session.get('user_email', 'anonymous'), 'set_model', f"model={model}")
        print(f"[模型切換] → {model}")
        return jsonify({"status": "success", "model": model})
    except Exception as e:
        log_error(session.get('user_email', 'anonymous'), 'set_model', e, f"model={data.get('model', '') if 'data' in dir() else ''}")
        return jsonify({"status": "error", "error": str(e)}), 500

@agent_bp.route('/reset', methods=['POST'])
def reset_agent():
    """重置 Agent 對話歷史"""
    try:
        agent = get_agent()
        agent.reset(keep_system=True)
        log_action(session.get('user_email', 'anonymous'), 'reset_agent', '')
        return jsonify({
            "status": "success",
            "message": "對話已重置"
        })
    except Exception as e:
        log_error(session.get('user_email', 'anonymous'), 'reset_agent', e)
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
        log_error(session.get('user_email', 'anonymous'), 'get_history', e)
        return jsonify({"error": str(e)}), 500

@agent_bp.route('/get_excel_defaults', methods=['GET'])
def get_excel_defaults():
    """獲取 Excel 輸入公版的預設值（C16 初始價金、C17 利潤率、C18 開發費）"""
    try:
        from tool.equipment_cost_tool import EquipmentCostTool

        tool = EquipmentCostTool()

        # 優先使用前端傳入的案場參數來找到對應 Excel 檔案
        case_id = request.args.get('case_id', '')
        case_name = request.args.get('case_name', '')
        original_filename = request.args.get('original_filename', '')

        excel_file = _find_excel_file(case_id=case_id, case_name=case_name, original_filename=original_filename)

        # 先設定基本預設值（備用值）
        defaults = {
            "equipment_cost": 30000,
            "profit_rate": 0.2,
            "development_fee": 0,
            "boundary": 20000,
            "cash_step": 2000,
            "ratio_step": 0.05,
            "max_value": 50000,
            "min_value": 30000,
            "cond_step_1": 2000,
            "cond_step_2": 1000,
            "cond_step_3": 500,
            "adjust_times": 5
        }

        # 如果有 Excel 檔案，則用 Excel 中的實際值覆寫預設值
        if excel_file:
            try:
                data = tool._read_excel_data(excel_file)
                defaults.update({
                    "equipment_cost": data.get("equipment_cost", 30000),
                    "profit_rate": data.get("profit_rate", 0.2),
                    "development_fee": data.get("development_fee", 0)
                })
                print(f"已從Excel讀取預設值：equipment_cost={defaults['equipment_cost']}, profit_rate={defaults['profit_rate']}, development_fee={defaults['development_fee']}")
            except Exception as e:
                print(f"警告：讀取 Excel 檔案失敗，使用備用預設值: {e}")

        response = jsonify({
            "status": "success",
            "defaults": defaults,
            "excel_file": excel_file
        })
        response.headers['Cache-Control'] = 'no-store'
        return response

    except Exception as e:
        log_error(session.get('user_email', 'anonymous'), 'get_excel_defaults', e)
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
        excel_dir = _get_excel_dir()
        if not os.path.exists(excel_dir):
            os.makedirs(excel_dir)
            print(f"創建 Excel 目錄: {excel_dir}")

        # 優先用 case_id 作為前綴（永久唯一，不受改名影響）
        if case_id:
            filename = f"{case_id}_{original_filename}"  # 例如：1_公版.xlsx
        else:
            filename = f"{case_name}_{original_filename}"  # fallback
        file_path = os.path.join(excel_dir, filename)
        file.save(file_path)
        print(f"Excel 檔案已儲存: {file_path}")
        print(f"檔案資訊: case_id={case_id}, case_name={case_name}, filename={filename}")

        # 嘗試從 Excel B2 讀取案場名稱
        project_name = None
        try:
            from openpyxl import load_workbook as _lw
            _wb = _lw(file_path, data_only=True, read_only=True)
            _ws = _wb.active
            _val = _ws['B2'].value
            if _val and str(_val).strip() and str(_val).strip() != '-':
                project_name = str(_val).strip()
            _wb.close()
        except Exception:
            pass

        log_action(session.get('user_email', 'anonymous'), 'upload_excel',
                   f"case={case_name}({case_id}) file={original_filename}")
        return jsonify({
            "status": "success",
            "message": f"檔案已上傳: {original_filename}",
            "file_path": file_path,
            "filename": filename,
            "original_filename": original_filename,
            "project_name": project_name
        })
    except Exception as e:
        import traceback
        print(f"檔案上傳錯誤: {e}")
        print(traceback.format_exc())
        log_error(session.get('user_email', 'anonymous'), 'upload_excel', e)
        return jsonify({"error": f"檔案上傳失敗: {str(e)}"}), 500


@agent_bp.route('/read_excel/<case_id>', methods=['GET'])
def read_excel(case_id):
    """
    讀取指定案場的 Excel 檔案並轉換為前端格式
    """
    try:
        excel_dir = _get_excel_dir()

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

        # 使用 openpyxl 讀取 Excel（data_only=True 取計算後的值）
        import openpyxl

        wb = openpyxl.load_workbook(excel_file, data_only=True)
        sheets_data = []

        for sheet_index, sheet_name in enumerate(wb.sheetnames):
            sheet = wb[sheet_name]
            celldata = []

            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None or cell.value == '':
                        continue

                    v_obj = {
                        'v': cell.value,
                        'm': str(cell.value),
                        'ct': {
                            'fa': 'General',
                            't': 'n' if isinstance(cell.value, (int, float)) else 's'
                        }
                    }
                    celldata.append({
                        'r': cell.row - 1,
                        'c': cell.column - 1,
                        'v': v_obj
                    })

            max_row = sheet.max_row    if sheet.max_row    else 20
            max_col = sheet.max_column if sheet.max_column else 10

            sheets_data.append({
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
            })

        print(f"成功讀取 Excel: {excel_file}, 共 {len(sheets_data)} 個工作表")
        log_action(session.get('user_email', 'anonymous'), 'read_excel',
                   f"case_id={case_id} file={os.path.basename(excel_file)} sheets={len(sheets_data)}")
        return jsonify({
            'status': 'success',
            'data': sheets_data,
            'filename': os.path.basename(excel_file)
        })

    except Exception as e:
        import traceback
        print(f"讀取 Excel 錯誤: {e}")
        print(traceback.format_exc())
        log_error(session.get('user_email', 'anonymous'), 'read_excel', e, f"case_id={case_id}")
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

        excel_dir = _get_excel_dir()

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
        log_action(session.get('user_email', 'anonymous'), 'rename_excel',
                   f"old={old_filename} new={new_filename}")
        return jsonify({
            "status": "success",
            "message": "檔案重新命名成功",
            "new_filename": new_filename
        })

    except Exception as e:
        import traceback
        print(f"🚨 重新命名錯誤: {e}")
        print(traceback.format_exc())
        log_error(session.get('user_email', 'anonymous'), 'rename_excel', e,
                  f"old={old_filename if 'old_filename' in dir() else ''}")
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

        excel_dir = _get_excel_dir()
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
        log_action(session.get('user_email', 'anonymous'), 'delete_excel',
                   f"file={filename}")

        return jsonify({
            "status": "success",
            "message": "檔案刪除成功",
            "deleted_filename": filename
        })

    except Exception as e:
        import traceback
        print(f"🚨 刪除檔案錯誤: {e}")
        print(traceback.format_exc())
        log_error(session.get('user_email', 'anonymous'), 'delete_excel', e,
                  f"file={filename if 'filename' in dir() else ''}")
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

    # 格式化 IRR 加上 %，固定顯示兩位小數
    def format_irr(val):
        if val is None or val == 'N/A':
            return 'N/A'
        try:
            return f"{float(val):.2f}%"
        except (TypeError, ValueError):
            return f"{val}%"

    # 格式化金額，加千分位逗號
    def fmt_money(val):
        if val is None or val == 'N/A':
            return 'N/A'
        try:
            return f"{int(round(float(val))):,}"
        except (TypeError, ValueError):
            return str(val)

    # 檢查是 execute_price_rolling 還是 calculate_price_rolling 的結果
    if "result" in tool_result:
        # execute_price_rolling 格式 - 從 result 字段提取數據
        result_data = tool_result.get("result", {})
        summary = tool_result.get("summary", {})

        response_parts.append("### 滾算摘要")
        response_parts.append(f"- **初始價金**: {fmt_money(summary.get('initial_cost'))} 元/kW")
        response_parts.append(f"- **最終價金**: {fmt_money(summary.get('final_cost'))} 元/kW")
        response_parts.append(f"- **總降幅**: {fmt_money(summary.get('total_reduction'))} 元/kW\n")

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

                response_parts.append(f"| {fmt_money(price)} | {fmt_money(profit)} | {project_irr} | {cost_irr} | {equity_irr} |")

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
        response_parts.append(f"- **初始價金**: {fmt_money(initial_cost)} 元/kW")
        response_parts.append(f"- **最終價金**: {fmt_money(final_cost)} 元/kW")
        response_parts.append(f"- **總降幅**: {fmt_money(total_reduction)} 元/kW\n")

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

        response_parts.append(f"| {fmt_money(price)} | {fmt_money(profit)} | {project_irr} | {cost_irr} | {equity_irr} |")

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
    mode = data.get('mode')  # CashMode, RatioMode, CustomizeMode
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
    case_id = str(data.get('case_id', ''))
    original_filename = data.get('original_filename', '')
    sheet_name = data.get('sheet_name', None)

    # 驗證必要參數
    if not mode or boundary is None:
        return jsonify({"error": "缺少必要參數：mode 和 boundary"}), 400

    print(f"\n--- API 請求: 價金滾算 (直接調用工具) ---")
    print(f"模式: {mode}")
    print(f"參數: equipment_cost={equipment_cost}, profit_rate={profit_rate}, development_fee={development_fee}, boundary={boundary}")
    print(f"案場資訊: case_id={case_id}, case_name={case_name}, original_filename={original_filename}")
    log_action(session.get('user_email', 'anonymous'), 'calculate_price_rolling',
               f"case={case_name}({case_id}) mode={mode} boundary={boundary}")

    try:
        # 直接建立 ToolManager，不依賴 AIAgent（滾算不需要 LLM）
        from tool.tool_manager import ToolManager
        tool_manager = ToolManager()

        # 設定財務工具的 Excel 檔案（使用通用定址 helper）
        current_excel_path = _find_excel_file(case_id=case_id, case_name=case_name, original_filename=original_filename)
        if current_excel_path:
            print(f"已找到 Excel 檔案: {current_excel_path}")
        else:
            print(f"警告: 找不到 Excel 檔案 (case_id={case_id}, case_name={case_name})")

        # 設定工具的 Excel 檔案
        if current_excel_path:
            tool_manager.set_finance_excel_file(current_excel_path, sheet_name)
            tool_manager.price_rolling_tool.set_excel_file(current_excel_path, sheet_name)
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
        tool_result = tool_manager.execute_tool("calculate_price_rolling", tool_args)

        # 格式化結果
        if tool_result.get("success"):
            # 成功執行後，儲存滾算參數到快取，供後續確認保存時使用
            from services.llm_service import _store_rolling_cache
            _store_rolling_cache(current_excel_path, backend_mode, tool_args)

            # 儲存待確認保存的滾算結果，等待使用者決定是否寫回 Excel
            if case_name or case_id:
                _store_pending_rolling_save(case_name, backend_mode, tool_args, current_excel_path, case_id=case_id, tool_result=tool_result)

            agent_response = _format_price_rolling_result(tool_result, mode)

            if agent_response:
                # 詢問使用者是否確認將滾算結果寫回 Excel
                agent_response += "\n\n---\n**是否要將此滾算結果儲存到 Excel？**\n請回覆「是」或「儲存」來保存，或直接進行下一步操作。"

                log_action(session.get('user_email', 'anonymous'), 'calculate_price_rolling_done',
                           f"case={case_name}({case_id}) mode={mode}")
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
        log_error(session.get('user_email', 'anonymous'), 'calculate_price_rolling', e,
                  f"case={case_name}({case_id}) mode={mode}")
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

        excel_dir = _get_excel_dir()
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
        log_action(session.get('user_email', 'anonymous'), 'save_excel',
                   f"case={case_name} file={os.path.basename(excel_path)} formulas_preserved={formulas_preserved}")
        return jsonify({"status": "success", "message": f"Excel 已保存（保留了 {formulas_preserved} 個公式）"})

    except Exception as e:
        import traceback
        print(f"[save_excel] 保存失敗: {e}")
        print(traceback.format_exc())
        log_error(session.get('user_email', 'anonymous'), 'save_excel', e,
                  f"case={case_name if 'case_name' in dir() else ''}")
        return jsonify({"status": "error", "error": f"保存失敗: {str(e)}"}), 500


@agent_bp.route('/download_excel', methods=['GET'])
def download_excel():
    """
    下載當前聊天室的 Excel 檔案
    確保每個聊天室下載的是自己對應的 Excel 檔案
    """
    try:
        # 獲取案場資訊
        case_id = request.args.get('case_id', '')
        case_name = request.args.get('case_name', '')
        original_filename = request.args.get('original_filename', '')

        print(f"[下載請求] case_id={case_id}, 案場: {case_name}, 檔案: {original_filename}")

        excel_path = _find_excel_file(case_id=case_id, case_name=case_name, original_filename=original_filename)

        # 檢查是否找到檔案
        if not excel_path or not os.path.exists(excel_path):
            print(f"[下載錯誤] 找不到任何符合的檔案，case_id={case_id}, case_name={case_name}")
            return jsonify({
                "status": "error",
                "error": f"找不到案場「{case_name}」的 Excel 檔案"
            }), 404

        print(f"[下載] 發送檔案: {excel_path}")
        log_action(session.get('user_email', 'anonymous'), 'download_excel',
                   f"case={case_name}({case_id}) file={os.path.basename(excel_path)}")
        # 發送檔案
        return send_file(
            excel_path,
            as_attachment=True,
            download_name=os.path.basename(excel_path),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        import traceback
        print(f"🚨 下載 Excel 錯誤: {e}")
        print(traceback.format_exc())
        log_error(session.get('user_email', 'anonymous'), 'download_excel', e,
                  f"case={case_name}({case_id})")
        return jsonify({
            "status": "error",
            "error": f"下載失敗: {str(e)}"
        }), 500


@agent_bp.route('/download_template', methods=['GET'])
def download_template():
    """
    下載 Excel Generic Template/excel公版.xlsx 的第三個工作表（含公式），
    供使用者取得標準計算公版。
    """
    try:
        import io
        import openpyxl
        from openpyxl import Workbook
        from copy import copy

        template_path = os.path.join(parent_dir, 'Excel Generic Template', 'excel公版.xlsx')

        if not os.path.exists(template_path):
            return jsonify({'status': 'error', 'error': '找不到公版檔案'}), 404

        # 以 data_only=False 開啟，保留所有公式字串
        wb_src = openpyxl.load_workbook(template_path, data_only=False)

        if '空白範例' not in wb_src.sheetnames:
            return jsonify({'status': 'error', 'error': f'公版檔案中找不到「空白範例」工作表，可用工作表: {wb_src.sheetnames}'}), 400

        src_sheet = wb_src['空白範例']

        # 建立新 workbook，將來源工作表的儲存格逐一複製
        from openpyxl.formula.translate import Translator
        from openpyxl.utils import get_column_letter, column_index_from_string

        wb_new = Workbook()
        ws_new = wb_new.active
        ws_new.title = src_sheet.title

        DOWNLOAD_ROWS = 35
        SRC_START_COL = 2   # 來源從 B 欄（col=2）開始
        COL_OFFSET = -1     # 整體向左平移 1 欄，B→A、C→B …

        for row in src_sheet.iter_rows(max_row=DOWNLOAD_ROWS, min_col=SRC_START_COL):
            for cell in row:
                dst_col = cell.column + COL_OFFSET
                value = cell.value

                # 公式欄位：同步平移公式內的欄位參照
                if isinstance(value, str) and value.startswith('='):
                    src_coord = f"{get_column_letter(cell.column)}{cell.row}"
                    dst_coord = f"{get_column_letter(dst_col)}{cell.row}"
                    try:
                        value = Translator(value, src_coord).translate_formula(dst_coord)
                    except Exception:
                        pass  # 翻譯失敗保留原公式

                new_cell = ws_new.cell(row=cell.row, column=dst_col, value=value)
                if cell.has_style:
                    new_cell.font        = copy(cell.font)
                    new_cell.fill        = copy(cell.fill)
                    new_cell.border      = copy(cell.border)
                    new_cell.alignment   = copy(cell.alignment)
                    new_cell.number_format = cell.number_format

        # 複製欄寬與列高（欄號同步平移）
        for col_letter, col_dim in src_sheet.column_dimensions.items():
            src_col_idx = column_index_from_string(col_letter)
            if src_col_idx >= SRC_START_COL:
                dst_col_letter = get_column_letter(src_col_idx + COL_OFFSET)
                ws_new.column_dimensions[dst_col_letter].width = col_dim.width
        for row_num, row_dim in src_sheet.row_dimensions.items():
            if row_num <= DOWNLOAD_ROWS:
                ws_new.row_dimensions[row_num].height = row_dim.height

        # 複製合併儲存格（欄號同步平移，排除原始 A 欄跨入的範圍）
        for merged_range in src_sheet.merged_cells.ranges:
            if merged_range.min_row <= DOWNLOAD_ROWS and merged_range.min_col >= SRC_START_COL:
                ws_new.merge_cells(
                    start_row=merged_range.min_row,
                    start_column=merged_range.min_col + COL_OFFSET,
                    end_row=merged_range.max_row,
                    end_column=merged_range.max_col + COL_OFFSET
                )

        # 輸出為記憶體串流並回傳
        output = io.BytesIO()
        wb_new.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name='excel公版.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        import traceback
        print(f"下載公版錯誤: {e}\n{traceback.format_exc()}")
        log_error(session.get('user_email', 'anonymous'), 'download_template', e)
        return jsonify({'status': 'error', 'error': str(e)}), 500


@agent_bp.route('/list_case_sheets', methods=['GET'])
def list_case_sheets():
    """
    列出所有案場及其 Excel 檔案中的 sheets
    用於匯入表格功能
    """
    try:
        import openpyxl

        excel_dir = _get_excel_dir()
        cases_data = []

        if not os.path.exists(excel_dir):
            return jsonify({
                "status": "success",
                "cases": []
            })

        # 取得資料庫中仍有效的 case id 集合，過濾孤立檔案
        valid_case_ids = set()
        try:
            from models.db_models import Case
            from models.database import db
            user_id = session.get('user_id')
            if user_id:
                valid_case_ids = {str(c.id) for c in Case.query.filter_by(user_id=user_id).all()}
        except Exception as _dbe:
            print(f"[list_case_sheets] 無法查詢 DB，略過驗證: {_dbe}")

        # 掃描 Excel 目錄中的所有檔案
        excel_files = [f for f in os.listdir(excel_dir)
                      if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

        for filename in excel_files:
            # 若能取得 valid_case_ids，跳過已刪除案場的孤立檔案
            if valid_case_ids:
                file_case_id = filename.split('_', 1)[0]
                if file_case_id not in valid_case_ids:
                    print(f"[list_case_sheets] 跳過孤立檔案（case 已刪除）: {filename}")
                    continue
            file_path = os.path.join(excel_dir, filename)

            try:
                # 檔名格式為 {case_id}_{original_filename}，case_id 是數字前綴
                parts = filename.split('_', 1)
                if len(parts) == 2:
                    case_name = parts[0]       # 數字 ID（用於路由）
                    original_filename = parts[1]
                else:
                    case_name = filename.rsplit('.', 1)[0]
                    original_filename = filename

                # 讀取 Excel 檔案的所有 sheet 名稱，順便讀 B2 取得真正的案場名稱
                workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                sheet_names = workbook.sheetnames
                # 嘗試從第一個工作表的 B2 讀案場名稱
                display_name = None
                try:
                    _first_ws = workbook[sheet_names[0]] if sheet_names else workbook.active
                    _b2 = _first_ws['B2'].value if _first_ws else None
                    if _b2 and str(_b2).strip() and str(_b2).strip() != '-':
                        display_name = str(_b2).strip()
                except Exception:
                    pass
                workbook.close()

                if not display_name:
                    display_name = original_filename.rsplit('.', 1)[0]  # 去副檔名

                cases_data.append({
                    "case_name": case_name,
                    "display_name": display_name,
                    "filename": filename,
                    "original_filename": original_filename,
                    "sheets": sheet_names,
                    "site_type": "single"  # 預設為單站，前端會根據 localStorage 更新
                })

                print(f"[list_case_sheets] {display_name} ({case_name}): {len(sheet_names)} sheets")

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
        log_error(session.get('user_email', 'anonymous'), 'list_case_sheets', e)
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
        return None, {}  # 回傳 (last_row, item_total_rows)

    overall_start = min(c['start_year'] for c in valid_cases)
    overall_end   = max(c['end_year']   for c in valid_cases)
    year_axis     = list(range(overall_start, overall_end + 1))
    n_years       = len(year_axis)
    last_col      = 3 + n_years  # C=3, D=4, 最後一個年份欄 = 4+(n_years-1) = 3+n_years

    # 建立（或清空重用）彙整總表
    SHEET_NAME = "多站彙整"
    from openpyxl.styles import DEFAULT_FONT
    _no_fill  = PatternFill(fill_type=None)
    _def_font = Font()
    if SHEET_NAME in target_wb.sheetnames:
        # 清空原有內容「值＋樣式」，避免舊顏色殘留到新位置
        agg_ws = target_wb[SHEET_NAME]
        for row in agg_ws.iter_rows():
            for cell in row:
                cell.value = None
                cell.fill  = _no_fill
                cell.font  = _def_font
    else:
        agg_ws = target_wb.create_sheet(title=SHEET_NAME, index=0)

    # Row 4：全域年份標題（只出現一次）
    agg_ws['C4'] = "年份"
    for i, year in enumerate(year_axis):
        agg_ws.cell(row=4, column=4 + i, value=year)
    fill_row(agg_ws, 4, 3, last_col, YEAR_ROW_COLOR, font_color="FFFFFF")

    current_row = 5
    item_total_rows = {}  # 記錄每個項目「年總計」列的實際列號

    for item_idx, item in enumerate(ITEMS):
        item_name   = item["name"]
        source_row  = item["row"]
        total_color = ITEM_TOTAL_COLORS.get(item_name, "FFFFFF")

        # 每個項目上方都有一排獨立的年份列（深藍底 + 白字）
        # item_idx == 0 已有 row 4 作為全域年份列，不重複插入
        if item_idx > 0:
            agg_ws.cell(row=current_row, column=3, value="年份")
            for i, year in enumerate(year_axis):
                agg_ws.cell(row=current_row, column=4 + i, value=year)
            fill_row(agg_ws, current_row, 3, last_col, YEAR_ROW_COLOR, font_color="FFFFFF")
            current_row += 1

        # 項目名稱列（無背景色）
        agg_ws.cell(row=current_row, column=3, value=item_name)
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
                    c = agg_ws.cell(row=current_row, column=4 + i,
                                    value=f"=ROUND(ABS('{safe_name}'!{src_col_letter}{source_row}),0)")
                    c.number_format = '#,##0'
                # else: 留空

            current_row += 1

        case_end_row = current_row - 1

        # 年總計列（填入 SUM 公式並套上對應顏色）
        agg_ws.cell(row=current_row, column=3, value=f"{item_name}/年總計")
        for i in range(n_years):
            col        = 4 + i
            col_letter = get_column_letter(col)
            c = agg_ws.cell(row=current_row, column=col,
                            value=f"=SUM({col_letter}{case_start_row}:{col_letter}{case_end_row})")
            c.number_format = '#,##0'
        fill_row(agg_ws, current_row, 3, last_col, total_color)
        item_total_rows[item_name] = current_row  # 記錄此項目年總計列號
        current_row += 1

    # 自動調整欄寬（依內容長度計算，避免顯示 #####）
    col_widths = {}
    for row in agg_ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            val_str = str(cell.value) if not str(cell.value).startswith('=') else ""
            if val_str:
                col_widths[cell.column] = max(col_widths.get(cell.column, 0), len(val_str))
    for col_num, width in col_widths.items():
        agg_ws.column_dimensions[get_column_letter(col_num)].width = width + 4  # 加 4 作為留白
    # 年份欄（D 欄起）的公式值無法靠字串長度判斷，確保至少 14 寬（足夠顯示千萬級數字）
    for i in range(n_years):
        col_letter = get_column_letter(4 + i)
        current_width = agg_ws.column_dimensions[col_letter].width or 0
        if current_width < 14:
            agg_ws.column_dimensions[col_letter].width = 14

    print(f"[彙整] 已建立「{SHEET_NAME}」，年份 {overall_start}~{overall_end}，"
          f"{len(ITEMS)} 個項目，{len(valid_cases)} 個案場")
    return current_row, item_total_rows  # 回傳第一個空白列 + 各項目年總計列號


def create_income_cashflow_section(target_wb, cases_info, agg_last_row, item_total_rows=None, finance_params=None):
    """
    在彙整總表工作表的 agg_last_row 下方（空 6 列）附加：
      - 綜合損益表：範本 B37:W80，資料欄從 D(col4) 開始
      - 現金流量表：範本 B82:W109，資料欄從 E(col5) 開始
    完整複製範本的值（含公式）和樣式，不做任何清除或覆寫。
    超過 20 年時，以指定欄位為樣板向右擴充欄位樣式。
    item_total_rows: 各項目在彙整總表上方「年總計」列號，用於填入綜合損益表公式。
    """
    import os, openpyxl
    from copy import copy
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell

    # ── 範本路徑 ──
    current_dir   = os.path.dirname(os.path.abspath(__file__))
    parent_dir    = os.path.dirname(current_dir)
    template_path = os.path.join(parent_dir, 'Excel Generic Template', 'excel公版.xlsx')
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
    template_ws = template_wb['空白範例']
    agg_ws = target_wb["多站彙整"]

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

    # 用彙整總表上方各項目的年總計列填入綜合損益表資料欄
    if item_total_rows:
        row_off = income_start - 37  # 範本列號 → 彙整總表實際列號的偏移量

        def t2a(tmpl_row):
            """將範本列號轉換為彙整總表的實際列號"""
            return tmpl_row + row_off

        # ── 第一張綜合損益表（範本 37~58）──

        # 一般項目：D 欄起引用上方總表年總計（44~54 行為支出項，加負號）
        FIRST_TABLE_ITEMS = [
            {"name": "預估發電度數", "template_row": 40, "expense": False},
            {"name": "預估收入",     "template_row": 41, "expense": False},
            {"name": "租金",         "template_row": 45, "expense": True},
            {"name": "運維",         "template_row": 48, "expense": True},
            {"name": "保險",         "template_row": 49, "expense": True},
            {"name": "模組回收費",   "template_row": 50, "expense": True},
        ]
        for item in FIRST_TABLE_ITEMS:
            total_row = item_total_rows.get(item["name"])
            if total_row is None:
                continue
            dst_row = t2a(item["template_row"])
            sign = "-" if item["expense"] else ""
            for i in range(n_years):
                col = 4 + i
                col_letter = get_column_letter(col)
                agg_ws.cell(row=dst_row, column=col, value=f"={sign}{col_letter}{total_row}")

        # 設備費用（C44）：C 欄橫向加總上方總表設備折舊年總計，D 欄之後全填 0
        equip_total_row = item_total_rows.get("設備折舊")
        if equip_total_row is not None:
            dst_row_43 = t2a(44)
            first_col_letter = get_column_letter(4)
            last_col_letter  = get_column_letter(3 + n_years)
            agg_ws.cell(row=dst_row_43, column=3,
                        value=f"=-SUM({first_col_letter}{equip_total_row}:{last_col_letter}{equip_total_row})")

        # 所得稅（row 54）：引用第二張綜合損益表的所得稅（row 77），支出加負號
        dst_row_54 = t2a(54)
        dst_row_77 = t2a(77)
        for i in range(n_years):
            col = 4 + i
            col_letter = get_column_letter(col)
            agg_ws.cell(row=dst_row_54, column=col, value=f"=-{col_letter}{dst_row_77}")

        # 現金流（row 55）：D 欄起 =SUM(D41:D54)；C 欄 = 設備費用（C44）
        dst_row_55 = t2a(55)
        dst_row_41 = t2a(41)
        for i in range(n_years):
            col = 4 + i
            col_letter = get_column_letter(col)
            agg_ws.cell(row=dst_row_55, column=col,
                        value=f"=SUM({col_letter}{dst_row_41}:{col_letter}{dst_row_54})")
        # C55 = 設備費用（C44），作為 IRR 的初始現金流（負值投入）
        agg_ws.cell(row=dst_row_55, column=3, value=f"=C{t2a(44)}")

        # IRR（C58）：=IRR(C55:last_col55)，從 C 欄（設備費用）開始計算
        dst_row_58 = t2a(58)
        last_data_letter = get_column_letter(3 + n_years)
        agg_ws.cell(row=dst_row_58, column=3,
                    value=f"=IRR(C{dst_row_55}:{last_data_letter}{dst_row_55})")

        # ── 第二張綜合損益表（範本 60~80）──

        # 一般項目（含設備折舊）：D 欄起引用上方總表年總計；支出項加負號（利息另處理）
        SECOND_TABLE_ITEMS = [
            {"name": "預估發電度數", "template_row": 63, "expense": False},
            {"name": "預估收入",     "template_row": 64, "expense": False},
            {"name": "設備折舊",     "template_row": 67, "expense": True},
            {"name": "租金",         "template_row": 68, "expense": True},
            {"name": "運維",         "template_row": 71, "expense": True},
            {"name": "保險",         "template_row": 72, "expense": True},
            {"name": "模組回收費",   "template_row": 73, "expense": True},
        ]
        for item in SECOND_TABLE_ITEMS:
            total_row = item_total_rows.get(item["name"])
            if total_row is None:
                continue
            dst_row = t2a(item["template_row"])
            sign = "-" if item["expense"] else ""
            for i in range(n_years):
                col = 4 + i
                col_letter = get_column_letter(col)
                agg_ws.cell(row=dst_row, column=col, value=f"={sign}{col_letter}{total_row}")

        # 稅前淨利（row 76）：=SUM(D64:D75)
        dst_row_76 = t2a(76)
        dst_row_64 = t2a(64)
        dst_row_75 = t2a(75)
        for i in range(n_years):
            col = 4 + i
            col_letter = get_column_letter(col)
            agg_ws.cell(row=dst_row_76, column=col,
                        value=f"=SUM({col_letter}{dst_row_64}:{col_letter}{dst_row_75})")

        # 所得稅（row 77）：=稅前淨利 * 0.2
        for i in range(n_years):
            col = 4 + i
            col_letter = get_column_letter(col)
            agg_ws.cell(row=dst_row_77, column=col, value=f"={col_letter}{dst_row_76}*0.2")

        # 稅後淨利（row 78）：=稅前淨利 - 所得稅
        dst_row_78 = t2a(78)
        for i in range(n_years):
            col = 4 + i
            col_letter = get_column_letter(col)
            agg_ws.cell(row=dst_row_78, column=col,
                        value=f"={col_letter}{dst_row_76}-{col_letter}{dst_row_77}")

    # ── 現金流量表（B82:W109，資料從 E=col5，延伸樣板 E 欄）──
    cashflow_start = income_end + 1
    copy_block(82, 109, 2, 23, 5, cashflow_start, ext_template_col=5)
    cashflow_end = cashflow_start + (109 - 82)

    # 現金流量表公式填入（引用第二張綜合損益表的值）
    if item_total_rows:
        def cf_row(tmpl_row):
            return cashflow_start + (tmpl_row - 82)

        _dst_78 = income_start + (78 - 37)   # 第二張損益表：稅後淨利
        dst_r85 = cf_row(85)
        dst_r86 = cf_row(86)
        dst_r87 = cf_row(87)

        # row 86 設備折舊：引用最上方總表的設備折舊年總計（正數）
        _equip_agg_row = item_total_rows.get("設備折舊")

        for i in range(n_years):
            col = 4 + i
            col_letter = get_column_letter(col)
            # 稅後淨利（row 85）：引用第二張損益表稅後淨利（正數）
            agg_ws.cell(row=dst_r85, column=col, value=f"={col_letter}{_dst_78}")
            # 設備折舊（row 86）：引用最上方總表的設備折舊年總計（正數）
            if _equip_agg_row:
                agg_ws.cell(row=dst_r86, column=col, value=f"={col_letter}{_equip_agg_row}")
            # 營運活動現金流量（row 87）：稅後淨利 + 設備折舊
            agg_ws.cell(row=dst_r87, column=col,
                        value=f"={col_letter}{dst_r85}+{col_letter}{dst_r86}")

    # ── 現金流量表延伸公式（89, 92, 93, 94, 95, 96 行）──
    if finance_params:
        def cf2(r):
            return cashflow_start + (r - 82)

        # 參數格位置（同張工作表）
        p_loan   = f"C{agg_last_row + 1}"   # 貸款成數(%)
        p_repay  = f"C{agg_last_row + 2}"   # 還款期數
        p_div    = f"C{agg_last_row + 3}"   # 股利比率(%)
        p_cap    = f"C{agg_last_row + 4}"   # 年底減資攤還期數
        p_bank   = f"C{agg_last_row + 5}"   # 銀行利率(%)

        # 實際數值（用來決定填值範圍）
        repay_val  = int(finance_params.get('repay_periods') or 0)
        cap_val    = int(finance_params.get('cap_reduction_periods') or 0)

        r86 = cf2(86)   # 設備折舊行
        r85 = cf2(85)   # 稅後淨利行（現金流量表）
        r89 = cf2(89);  r92 = cf2(92);  r93 = cf2(93)
        r94 = cf2(94);  r95 = cf2(95);  r96 = cf2(96)

        first_col_l = get_column_letter(4)
        last_col_l  = get_column_letter(3 + n_years)

        # Row 89：設備支出 = 第一張綜合損益表的設備費用（C44）
        _income_r44 = income_start + (44 - 37)
        agg_ws.cell(row=r89, column=4,
                    value=f"=C{_income_r44}")

        # Row 92：借款（負數，現金流出）= -設備支出絕對值 * 貸款成數(小數)（D欄單格）
        agg_ws.cell(row=r92, column=4,
                    value=f"=-D{r89}*{p_loan}")

        # Row 93：還款 = -|借款|/攤還期數（負數，現金流出），從 E 欄開始填 repay_val 年
        if repay_val > 0:
            for i in range(min(repay_val, n_years - 1)):
                agg_ws.cell(row=r93, column=5 + i,
                            value=f"=-D{r92}/{p_repay}")

        # Row 94：現金增資（負數，現金流出）= -設備支出絕對值 * (1-貸款成數(小數))（D欄單格）
        agg_ws.cell(row=r94, column=4,
                    value=f"=-D{r89}*(1-{p_loan})")

        # Row 95：現金股利 = 前一年稅後淨利(row85) * 股利比率(小數)，從 E 欄起
        for i in range(n_years - 1):
            prev_col_l = get_column_letter(4 + i)
            agg_ws.cell(row=r95, column=5 + i,
                        value=f"=-{prev_col_l}{r85}*{p_div}")

        # Row 96：年底減資 = -現金增資/攤還期數，從最後一年往前回填 cap_val 年
        if cap_val > 0:
            for i in range(min(cap_val, n_years)):
                col = 3 + n_years - i   # 最後一欄往前
                agg_ws.cell(row=r96, column=col,
                            value=f"=-D{r94}/{p_cap}")

        # ── row 99/100/101：現金淨流入出、期初、期末 ──
        r99  = cf2(99);  r100 = cf2(100);  r101 = cf2(101)
        r87_act = cf2(87);  r96_act = cf2(96)

        for i in range(n_years):
            col   = 4 + i
            col_l = get_column_letter(col)

            # Row 99：現金淨流入(出) = SUM(row87:row96) 同欄
            agg_ws.cell(row=r99, column=col,
                        value=f"=SUM({col_l}{r87_act}:{col_l}{r96_act})")

            # Row 100：期初淨現金
            if i == 0:
                agg_ws.cell(row=r100, column=col, value=0)
            else:
                prev_l = get_column_letter(col - 1)
                agg_ws.cell(row=r100, column=col, value=f"={prev_l}{r101}")

            # Row 101：期末淨現金 = 99 + 100
            agg_ws.cell(row=r101, column=col,
                        value=f"={col_l}{r99}+{col_l}{r100}")

        # ── row 104/105/106/107/109：成本法、權益法、借款餘額 ──
        r104 = cf2(104);  r105 = cf2(105)
        r106 = cf2(106);  r107 = cf2(107)
        r109 = cf2(109)

        for i in range(n_years):
            col   = 4 + i
            col_l = get_column_letter(col)

            # Row 104：成本法實際現金流
            # 首年 = -現金增資；其餘 = -現金股利 - 年底減資
            if i == 0:
                agg_ws.cell(row=r104, column=col, value=f"=-D{r94}")
            else:
                agg_ws.cell(row=r104, column=col,
                            value=f"=-{col_l}{r95}-{col_l}{r96}")

            # Row 106：權益法實際現金流
            # 首年 = 稅後淨利 - 現金增資；其餘 = 稅後淨利 - 年底減資
            if i == 0:
                agg_ws.cell(row=r106, column=col, value=f"={col_l}{r85}-D{r94}")
            else:
                agg_ws.cell(row=r106, column=col,
                            value=f"={col_l}{r85}-{col_l}{r96}")

            # Row 109：借款餘額
            # 首年 = 借款；其餘 = 前一年餘額 + 當年還款
            if i == 0:
                agg_ws.cell(row=r109, column=col, value=f"=D{r92}")
            else:
                prev_l = get_column_letter(col - 1)
                agg_ws.cell(row=r109, column=col,
                            value=f"={prev_l}{r109}+{col_l}{r93}")

        # Row 105：成本法 IRR = IRR(D104:last_col104)
        agg_ws.cell(row=r105, column=3,
                    value=f"=IRR({first_col_l}{r104}:{last_col_l}{r104})")

        # Row 107：權益法 IRR = IRR(D106:last_col106)
        agg_ws.cell(row=r107, column=3,
                    value=f"=IRR({first_col_l}{r106}:{last_col_l}{r106})")

        # ── 第二張綜合損益表 row 70：利息費用 = -借款餘額 * 銀行利率(小數) ──
        # 借款餘額 = row 109（現金流量表），銀行利率 = p_bank
        dst_r70 = income_start + (70 - 37)
        for i in range(n_years):
            col   = 4 + i
            col_l = get_column_letter(col)
            agg_ws.cell(row=dst_r70, column=col,
                        value=f"=-{col_l}{r109}*{p_bank}")

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

    # ── 自動調整欄寬（參數列 + 綜合損益表 + 現金流量表全範圍）──
    def _cell_width(val):
        """估算儲存格顯示寬度；公式格用預設 16，CJK 字元計 2"""
        s = str(val)
        if s.startswith('='):
            return 16
        return sum(2 if '\u4e00' <= ch <= '\u9fff' else 1 for ch in s)

    col_widths = {}
    for row in agg_ws.iter_rows(min_row=agg_last_row + 1, max_row=cashflow_end):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value is not None:
                w = _cell_width(cell.value)
                col_widths[cell.column] = max(col_widths.get(cell.column, 0), w)

    for col_num, w in col_widths.items():
        col_letter = get_column_letter(col_num)
        agg_ws.column_dimensions[col_letter].width = min(max(w + 2, 10), 60)

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
        finance_params = data.get('finance_params', {})

        print(f"\n[import_sheets] 目標案場: {target_case_name}")
        print(f"[import_sheets] 目標檔案: {target_filename}")
        print(f"[import_sheets] 要匯入的 sheets: {sheets_to_import}")

        if not target_case_name:
            return jsonify({"status": "error", "error": "缺少目標案場名稱"}), 400

        if not sheets_to_import:
            return jsonify({"status": "error", "error": "請選擇至少一個 sheet"}), 400

        excel_dir = _get_excel_dir()

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
            wb.active.title = "多站彙整"
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

                # 從 C2 讀電站名稱，作為 sheet 命名與彙整總表標題
                plant_name = str(source_sheet['C2'].value).strip() if source_sheet['C2'].value else source_case_name

                # 新 sheet 名稱：電站名稱_原sheet名稱（例：台中電站_p2）
                new_sheet_name = f"{plant_name}_{source_sheet_name}"

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
                        'display_name': plant_name,
                        'sheet_name':   new_sheet_name,
                        'start_year':   start_year,
                        'end_year':     end_year,
                    })
                    print(f"[import_sheets] 年份讀取成功: {plant_name} {start_year}~{end_year}")
                except (TypeError, ValueError) as e:
                    print(f"[import_sheets] 無法讀取 {new_sheet_name} 的年份資料（C5/E5）: {e}")

                source_wb.close()
                imported_count += 1
                print(f"[import_sheets] 成功匯入: {source_case_name}/{source_sheet_name} -> {new_sheet_name}")

            except Exception as e:
                print(f"[import_sheets] 匯入 {source_case_name}/{source_sheet_name} 失敗: {e}")
                continue

        # 補入目標檔案中已存在的 sheet（追加彙總時保留舊紀錄）
        new_sheet_names = {c['sheet_name'] for c in cases_info}
        _system_sheets = {"多站彙整"}
        for sn in target_wb.sheetnames:
            if sn in _system_sheets or sn in new_sheet_names:
                continue
            try:
                _ws = target_wb[sn]
                _plant = str(_ws['C2'].value).strip() if _ws['C2'].value else sn
                _start = int(_ws['C5'].value)
                _end   = int(_ws['E5'].value)
                cases_info.insert(0, {
                    'display_name': _plant,
                    'sheet_name':   sn,
                    'start_year':   _start,
                    'end_year':     _end,
                })
                print(f"[import_sheets] 已存在 sheet 納入彙整: {sn} ({_plant}) {_start}~{_end}")
            except Exception:
                pass  # 非資料 sheet，略過

        # 建立彙整總表（自動觸發）
        print(f"[import_sheets] cases_info 共 {len(cases_info)} 筆: {[c['display_name'] for c in cases_info]}")
        agg_last_row = 0
        if imported_count > 0 and cases_info:
            agg_last_row, item_total_rows = create_aggregation_sheet(target_wb, cases_info)
            try:
                create_income_cashflow_section(target_wb, cases_info, agg_last_row, item_total_rows, finance_params)
            except Exception as e:
                import traceback
                print(f"[損益/現金流] 建立失敗（不影響主流程）: {e}")
                print(traceback.format_exc())

        # 將財務參數寫入彙整總表（總表與綜合損益表之間的空白區域）
        if agg_last_row > 0 and finance_params:
            agg_ws = target_wb["多站彙整"]
            param_start_row = agg_last_row + 1  # 緊接在總表後開始
            params_to_write = [
                ("貸款成數(%)",     finance_params.get('loan_ratio')),
                ("還款期數",       finance_params.get('repay_periods')),
                ("股利比率(%)",    finance_params.get('dividend_ratio')),
                ("年底減資攤還期數", finance_params.get('cap_reduction_periods')),
                ("銀行利率(%)",    finance_params.get('bank_rate')),
            ]
            _pct_names = {'貸款成數(%)', '股利比率(%)', '銀行利率(%)'}
            for i, (name, value) in enumerate(params_to_write):
                row = param_start_row + i
                agg_ws.cell(row=row, column=2, value=name)   # B 欄放名稱
                if value is not None:
                    # 百分比欄位除以 100 存成小數，讓 Excel 百分比格式正確顯示
                    write_val = value / 100 if name in _pct_names else value
                    agg_ws.cell(row=row, column=3, value=write_val)  # C 欄放值
            print(f"[import_sheets] 財務參數已寫入彙整總表 row {param_start_row}~{param_start_row+3}")

        # 儲存目標檔案
        target_wb.save(target_path)
        target_wb.close()

        # 用 LibreOffice 重算公式（否則公式欄位顯示 0）
        try:
            from utils.recalc import recalc as _recalc
            recalc_result = _recalc(target_path, timeout=90)
            print(f"[import_sheets] LibreOffice recalc: {recalc_result}")
        except Exception as _e:
            print(f"[import_sheets] recalc 失敗（略過）: {_e}")

        print(f"[import_sheets] 完成，共匯入 {imported_count} 個 sheets")
        log_action(session.get('user_email', 'anonymous'), 'import_sheets',
                   f"target={target_case_name} imported={imported_count}")
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
        log_error(session.get('user_email', 'anonymous'), 'import_sheets', e,
                  f"target={target_case_name if 'target_case_name' in dir() else ''}")
        return jsonify({
            "status": "error",
            "error": f"匯入失敗: {str(e)}"
        }), 500


@agent_bp.route('/delete_sheet', methods=['POST'])
def delete_sheet():
    """
    從後端 Excel 檔案中刪除指定工作表（前端手動刪除觸發）
    Body: { case_id, case_name, original_filename, sheet_name }
    """
    try:
        data = request.json
        case_id = data.get('case_id', '')
        case_name = data.get('case_name', '')
        original_filename = data.get('original_filename', '')
        sheet_name = data.get('sheet_name', '').strip()

        user_email = session.get('user_email', 'anonymous')

        if not sheet_name:
            return jsonify({"error": "缺少工作表名稱"}), 400

        excel_path = _find_excel_file(
            case_id=case_id,
            case_name=case_name,
            original_filename=original_filename
        )
        if not excel_path:
            return jsonify({"error": "找不到對應的 Excel 檔案"}), 404

        from tool.tool_manager import ToolManager
        tm = ToolManager()
        tm.set_finance_excel_file(excel_path)
        result = tm.execute_tool("delete_excel_sheet", {"sheet_name": sheet_name})

        if result.get("success"):
            log_action(user_email, 'delete_sheet',
                       f"case={case_name}({case_id}) sheet={sheet_name}")
            return jsonify({
                "status": "success",
                "message": result.get("message", ""),
                "deleted_sheet": result.get("deleted_sheet", sheet_name),
                "remaining_sheets": result.get("remaining_sheets", [])
            })
        else:
            return jsonify({"error": result.get("message", "刪除失敗")}), 400

    except Exception as e:
        import traceback
        print(f"🚨 刪除工作表錯誤: {e}")
        print(traceback.format_exc())
        log_error(session.get('user_email', 'anonymous'), 'delete_sheet', e)
        return jsonify({"error": f"刪除工作表失敗: {str(e)}"}), 500