"""
LLM 服務模組 - 簡化版
直接使用工具執行價金滾算，工具會自動輸出 Excel 滾算記錄
"""
import re
from datetime import datetime
from tool.tool_manager import ToolManager

# 初始化工具管理器
tool_manager = ToolManager()

# ========== 滾算參數快取 ==========
# 用於儲存每個案件（以 excel_path 為 key）的最後一次滾算參數
# 當使用者執行「執行滾算紀錄」時，會使用這裡的快取參數
_last_rolling_cache = {}


def _store_rolling_cache(excel_path: str, mode: str, params: dict):
    """
    儲存滾算參數到快取

    Args:
        excel_path: Excel 檔案路徑（作為快取的 key）
        mode: 滾算模式
        params: 滾算參數
    """
    cache_key = excel_path or "default"
    _last_rolling_cache[cache_key] = {
        "mode": mode,
        "params": params.copy() if params else {},
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    print(f"[快取] 已儲存滾算參數: key={cache_key}, mode={mode}")
    print(f"[快取] 參數內容: {params}")


def _get_rolling_cache(excel_path: str) -> dict:
    """
    從快取取得滾算參數

    Args:
        excel_path: Excel 檔案路徑（作為快取的 key）

    Returns:
        dict: 快取的滾算參數，若無快取則返回 None
    """
    cache_key = excel_path or "default"
    cached = _last_rolling_cache.get(cache_key)
    if cached:
        print(f"[快取] 找到快取: key={cache_key}, mode={cached['mode']}, 時間={cached['timestamp']}")
    else:
        print(f"[快取] 無快取: key={cache_key}")
    return cached


def _clear_rolling_cache(excel_path: str = None):
    """
    清除滾算參數快取

    Args:
        excel_path: 指定要清除的 key，若為 None 則清除全部
    """
    if excel_path:
        cache_key = excel_path or "default"
        if cache_key in _last_rolling_cache:
            del _last_rolling_cache[cache_key]
            print(f"[快取] 已清除: key={cache_key}")
    else:
        _last_rolling_cache.clear()
        print("[快取] 已清除全部快取")


def initialize_agent(sheet_name):
    """
    使用特定工作表名稱初始化 Agent。
    """
    print(f"Initializing Agent for sheet: {sheet_name}")


def process_user_query(query, simulation_amount=0, excel_path=None, rolling_mode=None, rolling_params=None, sheet_name=None):
    """
    處理使用者查詢。
    簡化流程：
    1. 獲取網頁輸入的模式和參數
    2. 直接調用對應的工具
    3. 工具會自動輸出 Excel 滾算記錄

    Args:
        query (str): 使用者查詢
        simulation_amount (float): 模擬金額（來自網頁輸入）
        excel_path (str, optional): 使用者上傳的 Excel 檔案路徑（當前聊天室的 Excel）
        rolling_mode (str, optional): 價金滾算模式（CashMode, RatioMode, ConditionalMode, CustomizeMode）
        rolling_params (dict, optional): 價金滾算參數（從網頁傳來）
        sheet_name (str, optional): Excel 工作表名稱
    """

    print(f"\n========== 處理使用者請求 ==========")
    print(f"查詢: {query}")
    print(f"Excel 路徑: {excel_path}")
    print(f"工作表: {sheet_name}")
    print(f"網頁指定模式: {rolling_mode}")
    print(f"網頁滾算參數: {rolling_params}")

    # 如果有提供 excel_path，更新 tool_manager 的設定
    if excel_path:
        tool_manager.set_finance_excel_file(excel_path, sheet_name)
        print(f"已設定 tool_manager Excel 路徑: {excel_path}")
    
    # ========== 直接使用網頁提供的模式和參數調用工具 ==========
    
    # 情況 1: 網頁明確指定了滾算模式和參數
    if rolling_mode and rolling_params:
        print(f"使用網頁參數直接調用滾算工具...")
        print(f"  模式: {rolling_mode}")

        # 直接使用 execute_price_rolling 工具（會自動輸出 Excel 記錄）
        return _execute_price_rolling_with_params(rolling_mode, rolling_params, excel_path, sheet_name)
    
    # 情況 2a: 用戶要求儲存滾算紀錄到 Excel
    elif re.search(r'執行滾算紀錄|儲存滾算|保存滾算|寫入excel|存檔', query, re.IGNORECASE):
        print(f"檢測到儲存滾算紀錄請求，執行 execute_price_rolling...")

        # 嘗試從快取取得上次的滾算參數
        cached = _get_rolling_cache(excel_path)

        if cached:
            # 使用快取的參數
            mode = cached["mode"]
            params = cached["params"]
            print(f"使用快取的滾算參數: mode={mode}, params={params}")
        else:
            # 無快取，提示用戶先執行滾算
            return """### ⚠️ 尚無滾算紀錄

您尚未執行價金滾算，無法儲存紀錄。

**請先執行以下步驟：**
1. 點擊上方的 **「價金滾算」** 按鈕
2. 選擇滾算模式並輸入參數
3. 執行滾算計算
4. 再輸入「執行滾算紀錄」來儲存結果

---
💡 **提示**: 滾算紀錄會使用您上一次執行的滾算參數。"""

        # 使用 execute_price_rolling 工具寫入 Excel
        result = _execute_equipment_cost_tool(mode, params, excel_path, sheet_name)

        if "執行失敗" not in result:
            # 取得快取中的參數摘要
            mode_display = {
                "cash": "CashMode（固定金額）",
                "ratio": "RatioMode（比例調整）",
                "conditional": "ConditionalMode（條件調整）",
                "customize": "CustomizeMode（自訂調整）"
            }.get(mode, mode)

            boundary = params.get("boundary", "N/A")
            step = params.get("step", "N/A")

            return f"""### ✅ 儲存完畢

滾算紀錄已成功寫入 Excel 檔案。

**使用的滾算參數：**
- **模式**: {mode_display}
- **邊界**: {boundary}
- **步伐**: {step}
- **儲存時間**: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}"""
        else:
            return result

    # 情況 2b: 其他滾算相關關鍵字 - 提示使用對話框
    elif re.search(r'滾算|價金|price.*rolling|IRR|設備成本|計算|分析|模擬|cashmode|ratiomode|conditional|customize', query, re.IGNORECASE):
        print(f"檢測到滾算相關請求，提示用戶使用對話框...")

        return """### 📊 價金滾算

請點擊上方的 **「價金滾算」** 按鈕來執行滾算計算。

若要將結果儲存至 Excel，請在聊天框輸入「執行滾算紀錄」。"""
    
    # 情況 3: 其他查詢 - 提供更友善的回應
    else:
        return f"""### 💡 您好！我是價金滾算分析助手

我可以幫您進行以下分析：

**📊 快速開始：**
- 輸入「計算 IRR」或「分析價金」開始滾算分析
- 或點擊上方「**價金滾算**」按鈕，選擇模式並輸入參數

**🔧 支援的模式：**
| 模式 | 說明 | 使用場景 |
|------|------|---------|
| **CashMode** | 固定金額調整 | 每次降價固定金額 |
| **RatioMode** | 比例調整 | 按百分比遞減 |
| **ConditionalMode** | 條件調整 | 不同價格區間不同步伐 |
| **CustomizeMode** | 自訂調整 | 手動指定每次降價 |

**📝 範例輸入：**
- 「用 CashMode 計算，邊界 20000，步伐 2000」
- 「幫我分析 IRR」
- 「執行價金滾算」

工具會自動讀取 Excel 數據並計算 IRR！
"""


def _parse_query_for_rolling_params(query, simulation_amount):
    """
    從查詢中解析滾算參數（簡化版，不使用 LLM）
    """
    params = {
        "mode": "cash",
        "boundary": 20000,
        "step": 1000
    }
    
    # 檢測模式
    if re.search(r'比率|比例|ratio', query, re.IGNORECASE):
        params["mode"] = "ratio"
        params["step"] = 0.05
    elif re.search(r'條件|conditional', query, re.IGNORECASE):
        params["mode"] = "conditional"
        params["maximum_value"] = 50000
        params["minimum_value"] = 30000
        params["condition_step_1"] = 2000
        params["condition_step_2"] = 1000
        params["condition_step_3"] = 500
    elif re.search(r'自訂|customize', query, re.IGNORECASE):
        params["mode"] = "customize"
        params["adjustment_times"] = 10
    
    # 從查詢中提取數字作為步長
    if simulation_amount > 0:
        params["step"] = int(simulation_amount)
    else:
        match = re.search(r'(\d+)', query)
        if match:
            params["step"] = int(match.group(1))
    
    # 提取邊界
    boundary_match = re.search(r'邊界.*?(\d+)', query)
    if boundary_match:
        params["boundary"] = int(boundary_match.group(1))
    
    return params


# ========== 輔助函數：價金滾算相關 ==========

def _execute_calculate_price_rolling(mode, params, excel_path):
    """
    使用純計算模式執行價金滾算（不輸出 Excel 記錄）
    用於聊天框輸入時的快速回應
    """
    print(f"執行純計算價金滾算...")
    print(f"  模式: {mode}")
    print(f"  參數: {params}")
    
    # 準備工具參數
    tool_params = {
        "mode": mode.capitalize() + "Mode" if not mode.endswith("Mode") else mode,
        "boundary": params.get("boundary", 20000),
        "step": params.get("step", 1000),
        "profit_rate": params.get("profit_rate", 0.2)
    }
    
    # 調用 calculate_price_rolling 工具（純計算，不輸出檔案）
    result = tool_manager.execute_tool("calculate_price_rolling", tool_params)
    
    if result.get("success"):
        return _format_calculate_result(result)
    else:
        return f"### ❌ 計算失敗\n\n**錯誤訊息:** {result.get('message', '未知錯誤')}\n\n請嘗試點擊「價金滾算」按鈕，使用完整的參數輸入表單。"


def _format_calculate_result(result):
    """
    格式化純計算結果為 Markdown
    """
    mode = result.get("mode", "CashMode")
    base_irr = result.get("base_irr", {})
    params = result.get("used_parameters", {})
    results = result.get("results_summary", {})
    data_rows = results.get("data", [])
    
    # 構建回應
    response = f"""## 📊 價金滾算分析結果 ({mode})

### 原始 Excel IRR (對照基準)
- **專案法 IRR**: {base_irr.get('project_irr', 'N/A')}%
- **成本法 IRR**: {base_irr.get('cost_method_irr', 'N/A')}%
- **權益法 IRR**: {base_irr.get('equity_method_irr', 'N/A')}%

### 使用參數
- **初始價金 / kW**: {params.get('equipment_cost', 'N/A')}
- **信邦利潤率**: {params.get('profit_rate', 'N/A')}
- **開發費**: {params.get('development_fee', 'N/A')}
- **邊界**: {params.get('boundary', 'N/A')}
- **步伐**: {params.get('step', 'N/A')}

### 滾算結果
| 價金/kW | 信邦利潤/kW | 最終價金/kW | 專案法 IRR | 成本法 IRR | 權益法 IRR |
| --- | --- | --- | --- | --- | --- |
"""
    
    for row in data_rows:
        price = row[0]
        profit = row[1]
        final = row[2]
        p_irr = f"{row[3]}%" if row[3] != 'N/A' else 'N/A'
        c_irr = f"{row[4]}%" if row[4] != 'N/A' else 'N/A'
        e_irr = f"{row[5]}%" if row[5] != 'N/A' else 'N/A'
        response += f"| {price} | {profit} | {final} | {p_irr} | {c_irr} | {e_irr} |\n"
    
    response += "\n---\n💡 **提示**: 若需保存記錄到 Excel，請點擊「執行滾算紀錄」按鈕。"
    
    return response


def _execute_equipment_cost_tool(mode, params, excel_path, sheet_name=None):
    """
    使用 execute_price_rolling 工具執行價金滾算
    這個工具會自動輸出 Excel 記錄到 'Excel final' 資料夾
    """
    print(f"執行 execute_price_rolling 工具（equipment_cost_tool）...")
    print(f"  模式: {mode}")
    print(f"  參數: {params}")
    print(f"  Excel 路徑: {excel_path}")

    # 準備工具參數
    tool_params = {
        "mode": mode.lower() if mode.lower() in ["cash", "ratio", "conditional", "customize"] else "cash",
        "boundary": params.get("boundary", 20000),
        "step": params.get("step", 1000),
    }

    # 【重要】添加 excel_file 參數，確保使用正確的檔案
    if excel_path:
        tool_params["excel_file"] = excel_path

    # 添加工作表名稱
    if sheet_name:
        tool_params["sheet_name"] = sheet_name

    # 添加可選參數
    if params.get("profit_rate"):
        tool_params["profit_rate"] = params["profit_rate"]
    if params.get("development_fee"):
        tool_params["development_fee"] = params["development_fee"]

    print(f"  工具參數: {tool_params}")
    
    # 調用 execute_price_rolling 工具
    result = tool_manager.execute_tool("execute_price_rolling", tool_params)
    
    if result.get("success"):
        return _format_equipment_cost_result(result)
    else:
        return f"### ❌ 執行失敗\n\n**錯誤訊息:** {result.get('message', '未知錯誤')}\n\n請嘗試點擊「價金滾算」按鈕，使用完整的參數輸入表單。"


def _format_equipment_cost_result(result):
    """
    格式化 execute_price_rolling 的結果為 Markdown
    """
    summary = result.get("summary", {})
    output_file = result.get("output_file", "")
    result_data = result.get("result", {})
    mode = result_data.get("mode", "cash")
    
    # 獲取滾算結果
    irr_results = result_data.get("irr_results", [])
    adjustment_record = result_data.get("adjustment_record", [])
    profit_record = result_data.get("profit_record", [])
    cost_structure_adjusted = result_data.get("cost_structure_adjusted", [])
    
    response = f"""## ✅ 價金滾算完成

### 📁 Excel 記錄已輸出
- **檔案位置**: `{output_file}`

### 滾算摘要
| 項目 | 數值 |
|------|------|
| 初始設備成本 | {summary.get('initial_cost', 'N/A')} 元/kWp |
| 最終設備成本 | {summary.get('final_cost', 'N/A')} 元/kWp |
| 調整次數 | {summary.get('adjustment_count', 'N/A')} 次 |
| 總降幅 | {summary.get('total_reduction', 'N/A')} 元/kWp |
| 滾算模式 | {mode} |

### 滾算結果（含 IRR）
| 價金/kW | 信邦利潤/kW | 最終價金/kW | 專案法 IRR | 成本法 IRR | 權益法 IRR |
| --- | --- | --- | --- | --- | --- |
"""
    
    # 顯示所有滾算結果
    for i, price in enumerate(adjustment_record):
        profit = profit_record[i] if i < len(profit_record) else 'N/A'
        final = cost_structure_adjusted[i] if i < len(cost_structure_adjusted) else 'N/A'
        
        if i < len(irr_results):
            irr = irr_results[i]
            # IRR 已經是百分比形式，不需要再乘以 100
            p_irr_val = irr.get('project_irr')
            c_irr_val = irr.get('cost_method_irr')
            e_irr_val = irr.get('equity_method_irr')
            
            p_irr = f"{p_irr_val:.2f}%" if p_irr_val is not None else 'N/A'
            c_irr = f"{c_irr_val:.2f}%" if c_irr_val is not None else 'N/A'
            e_irr = f"{e_irr_val:.2f}%" if e_irr_val is not None else 'N/A'
        else:
            p_irr = c_irr = e_irr = 'N/A'
        
        response += f"| {price} | {profit} | {final} | {p_irr} | {c_irr} | {e_irr} |\n"
    
    response += f"\n---\n📂 **完整結果已保存至**: `{output_file}`"
    
    return response


def _execute_price_rolling_with_params(rolling_mode, rolling_params, excel_path, sheet_name=None):
    """
    使用網頁提供的參數執行完整的價金滾算流程
    工具會自動輸出 Excel 滾算記錄
    """
    print(f"執行價金滾算並輸出 Excel 記錄...")
    print(f"  模式: {rolling_mode}")
    print(f"  參數: {rolling_params}")
    print(f"  Excel 路徑: {excel_path}")
    print(f"  工作表: {sheet_name}")

    # 準備工具參數
    tool_params = _convert_web_params_to_tool_params(rolling_mode, rolling_params)

    # 【重要】添加 excel_file 參數，確保使用正確的檔案
    if excel_path:
        tool_params["excel_file"] = excel_path

    # 添加工作表名稱（如果有提供且 rolling_params 中沒有）
    if sheet_name and "sheet_name" not in tool_params:
        tool_params["sheet_name"] = sheet_name

    # 調用 execute_price_rolling 工具（會自動輸出 Excel 記錄）
    result = tool_manager.execute_tool("execute_price_rolling", tool_params)

    if result.get("success"):
        # 【快取】成功執行後，儲存參數到快取供「執行滾算紀錄」使用
        _store_rolling_cache(excel_path, tool_params.get("mode", "cash"), tool_params)
        return _format_rolling_result_with_file(result)
    else:
        return f"### 執行價金滾算時發生錯誤\n\n- **錯誤訊息:** {result.get('message', '未知錯誤')}"


def _convert_web_params_to_tool_params(rolling_mode, rolling_params):
    """
    將網頁參數轉換為工具所需的參數格式
    
    Args:
        rolling_mode: 網頁指定的模式 (CashMode, RatioMode, ConditionalMode, CustomizeMode)
        rolling_params: 網頁提供的參數字典
    
    Returns:
        dict: 工具所需的參數格式
    """
    # 模式名稱映射（網頁用的可能是大寫開頭）
    mode_map = {
        "CashMode": "cash",
        "RatioMode": "ratio", 
        "ConditionalMode": "conditional",
        "CustomizeMode": "customize",
        "cash": "cash",
        "ratio": "ratio",
        "conditional": "conditional",
        "customize": "customize"
    }
    
    # 基本參數（所有模式都需要）
    tool_params = {
        "mode": mode_map.get(rolling_mode, "cash"),
        "boundary": rolling_params.get("boundary", 20000)
    }
    
    # 可選的基本參數
    if "equipment_cost" in rolling_params:
        tool_params["equipment_cost"] = rolling_params["equipment_cost"]
    
    if "profit_rate" in rolling_params:
        tool_params["profit_rate"] = rolling_params["profit_rate"]
    
    if "development_fee" in rolling_params:
        tool_params["development_fee"] = rolling_params["development_fee"]
    
    if "sheet_name" in rolling_params:
        tool_params["sheet_name"] = rolling_params["sheet_name"]
    
    # 根據不同模式添加特定參數
    if rolling_mode in ["CashMode", "cash"]:
        tool_params["step"] = rolling_params.get("step", 1000)
    
    elif rolling_mode in ["RatioMode", "ratio"]:
        tool_params["step"] = rolling_params.get("step", 0.05)
    
    elif rolling_mode in ["ConditionalMode", "conditional"]:
        tool_params["maximum_value"] = rolling_params.get("max_value", rolling_params.get("maximum_value", 50000))
        tool_params["minimum_value"] = rolling_params.get("min_value", rolling_params.get("minimum_value", 30000))
        tool_params["condition_step_1"] = rolling_params.get("step1", rolling_params.get("condition_step_1", 2000))
        tool_params["condition_step_2"] = rolling_params.get("step2", rolling_params.get("condition_step_2", 1000))
        tool_params["condition_step_3"] = rolling_params.get("step3", rolling_params.get("condition_step_3", 500))
    
    elif rolling_mode in ["CustomizeMode", "customize"]:
        tool_params["adjustment_times"] = rolling_params.get("adjust_times", rolling_params.get("adjustment_times", 10))
        if "steps" in rolling_params:
            tool_params["custom_steps"] = rolling_params["steps"]
        if "custom_steps" in rolling_params:
            tool_params["custom_steps"] = rolling_params["custom_steps"]
    
    print(f"  轉換後的工具參數: {tool_params}")
    return tool_params


def _format_rolling_result_with_file(result):
    """
    格式化包含檔案輸出的滾算結果
    """
    summary = result.get("summary", {})
    output_file = result.get("output_file", "")
    result_data = result.get("result", {})
    
    report = f"""
### ✅ **價金滾算完成**

**執行模式:** {result_data.get('mode', 'N/A')}

---

#### **滾算摘要**
- **初始價金:** ${summary.get('initial_cost', 0):,.0f} / kW
- **最終價金:** ${summary.get('final_cost', 0):,.0f} / kW
- **調整次數:** {summary.get('adjustment_count', 0)} 次
- **總降幅:** ${summary.get('total_reduction', 0):,.0f}

---

#### **📁 Excel 滾算記錄已輸出**
結果已保存至: `{output_file}`

---

#### **IRR 變化（前 5 筆）**

| 價金/kW | 專案法 IRR | 成本法 IRR | 權益法 IRR |
| :--- | :--- | :--- | :--- |
"""
    
    # 顯示前 5 筆 IRR 結果
    irr_results = result_data.get("irr_results", [])
    adjustment_record = result_data.get("adjustment_record", [])
    
    for i, (price, irr_data) in enumerate(zip(adjustment_record[:5], irr_results[:5])):
        p_irr = irr_data.get("project_irr")
        c_irr = irr_data.get("cost_method_irr")
        e_irr = irr_data.get("equity_method_irr")
        
        p_str = f"{p_irr*100:.2f}%" if p_irr is not None else "N/A"
        c_str = f"{c_irr*100:.2f}%" if c_irr is not None else "N/A"
        e_str = f"{e_irr*100:.2f}%" if e_irr is not None else "N/A"
        
        report += f"| ${price:,.0f} | {p_str} | {c_str} | {e_str} |\n"
    
    if len(adjustment_record) > 5:
        report += f"\n*（共 {len(adjustment_record)} 筆記錄，完整結果請查看 Excel 檔案）*\n"
    
    return report
