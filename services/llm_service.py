"""
LLM 服務模組 - 簡化版
直接使用工具執行價金滾算，工具會自動輸出 Excel 滾算記錄
"""
import re
from tool.tool_manager import ToolManager

# 初始化工具管理器
tool_manager = ToolManager()


def initialize_agent(sheet_name):
    """
    使用特定工作表名稱初始化 Agent。
    """
    print(f"Initializing Agent for sheet: {sheet_name}")


def process_user_query(query, simulation_amount=0, excel_path=None, rolling_mode=None, rolling_params=None):
    """
    處理使用者查詢。
    簡化流程：
    1. 獲取網頁輸入的模式和參數
    2. 直接調用對應的工具
    3. 工具會自動輸出 Excel 滾算記錄

    Args:
        query (str): 使用者查詢
        simulation_amount (float): 模擬金額（來自網頁輸入）
        excel_path (str, optional): 使用者上傳的 Excel 檔案路徑
        rolling_mode (str, optional): 價金滾算模式（CashMode, RatioMode, ConditionalMode, CustomizeMode）
        rolling_params (dict, optional): 價金滾算參數（從網頁傳來）
    """

    print(f"\n========== 處理使用者請求 ==========")
    print(f"查詢: {query}")
    print(f"網頁指定模式: {rolling_mode}")
    print(f"網頁滾算參數: {rolling_params}")
    
    # ========== 直接使用網頁提供的模式和參數調用工具 ==========
    
    # 情況 1: 網頁明確指定了滾算模式和參數
    if rolling_mode and rolling_params:
        print(f"使用網頁參數直接調用滾算工具...")
        print(f"  模式: {rolling_mode}")
        
        # 直接使用 execute_price_rolling 工具（會自動輸出 Excel 記錄）
        return _execute_price_rolling_with_params(rolling_mode, rolling_params, excel_path)
    
    # 情況 2: 從查詢中檢測滾算相關關鍵字
    elif re.search(r'滾算|價金|price.*rolling', query, re.IGNORECASE):
        print(f"檢測到滾算相關請求，使用預設參數...")
        
        # 解析查詢中的基本參數
        params = _parse_query_for_rolling_params(query, simulation_amount)
        mode = params.pop("mode", "cash")
        
        return _execute_price_rolling_with_params(mode, params, excel_path)
    
    # 情況 3: 其他查詢（暫不支援）
    else:
        return f"""### 請使用價金滾算功能

請點擊「價金滾算」按鈕，選擇模式並輸入參數後開始計算。

**支援的模式：**
- **CashMode** - 固定金額調整
- **RatioMode** - 比例調整  
- **ConditionalMode** - 條件調整
- **CustomizeMode** - 自訂調整

工具會自動將結果記錄到 Excel 檔案。
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

def _execute_price_rolling_with_params(rolling_mode, rolling_params, excel_path):
    """
    使用網頁提供的參數執行完整的價金滾算流程
    工具會自動輸出 Excel 滾算記錄
    """
    print(f"執行價金滾算並輸出 Excel 記錄...")
    print(f"  模式: {rolling_mode}")
    print(f"  參數: {rolling_params}")
    
    # 準備工具參數
    tool_params = _convert_web_params_to_tool_params(rolling_mode, rolling_params)
    
    # 調用 execute_price_rolling 工具（會自動輸出 Excel 記錄）
    result = tool_manager.execute_tool("execute_price_rolling", tool_params)
    
    if result.get("success"):
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
