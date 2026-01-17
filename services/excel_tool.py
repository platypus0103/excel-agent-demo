import pandas as pd
import numpy_financial as npf
import os
from datetime import datetime

# 預設路徑，可由函數參數覆寫
DEFAULT_EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'system_analysis.xlsx')
DEFAULT_SHEET_NAME = 'IRR (中庸2.5595%) 貸8成'

def find_row_by_label(df, label):
    """
    在 DataFrame 的前三欄中尋找標籤並返回其索引。
    此版本具有以下增強功能：
    1. 移除所有空格並忽略大小寫進行比對。
    2. 處理合併儲存格：如果儲存格為空，則向上回溯尋找最近的非空值。
    """
    # 預處理目標標籤：移除所有空格並轉為小寫
    target_label = label.replace(" ", "").lower()
    
    # 為了處理合併儲存格，我們需要一個「填補後」的 DataFrame 版本
    # ffill() 會將上一個非空值填入當前的空值中，模擬合併儲存格的視覺效果
    # 擴大搜尋範圍至前 3 欄 (因為 '年份' 可能在第 C 欄)
    df_filled = df.iloc[:, :3].ffill()
    
    for index, row in df_filled.iterrows():
        # 遍歷前三欄
        for i in range(min(3, len(row))):
            cell_value = row.iloc[i]
            if isinstance(cell_value, str):
                # 預處理儲存格內容：移除所有空格並轉為小寫
                clean_cell_value = cell_value.replace(" ", "").lower()
                if clean_cell_value == target_label:
                    return index
    return None

def get_irr_calculation_data(equipment_cost_adj=1.0, rent_adj=1.0, loan_amount=None, excel_path=None, sheet_name=None):
    """
    讀取 Excel 檔案並提取用於 IRR 計算的特定列。
    此版本新增了對關鍵參數進行敏感性分析的功能，並允許使用者自訂借款金額。

    Args:
        equipment_cost_adj (float): 設備成本的調整乘數 (例如 1.1 表示增加 10%)。
        rent_adj (float): 租金成本的調整乘數 (例如 0.9 表示減少 10%)。
        loan_amount (float, optional): 使用者自訂的借款金額。如果為 None，則從 Excel 讀取。
        excel_path (str, optional): 自訂 Excel 檔案路徑。若為 None，使用預設路徑。
        sheet_name (str, optional): 自訂工作表名稱。若為 None，使用預設名稱。

    返回包含提取數據和計算出的 IRR 的字典。
    """
    # 使用提供的路徑或預設路徑
    if excel_path is None:
        excel_path = DEFAULT_EXCEL_PATH
    if sheet_name is None:
        sheet_name = DEFAULT_SHEET_NAME

    if not os.path.exists(excel_path):
        return {"error": f"Excel file not found at {excel_path}"}

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        
        # --- 1. 定義關鍵列標籤 ---
        # 改為直接讀取「稅後淨利」與「現金增資」，不再逐項加總營運費用，以避免漏項導致計算錯誤。
        row_definitions = {
            "net_profit": "稅後淨利",
            "pre_tax_profit": "稅前淨利",
            "project_irr_items": [
                "Inflow : 電費收入",
                "Inflow : 政府補助",
                "Outflow : 設備費用",
                "Outflow : 設備放置區域租金",
                "Outflow : 維運費",
                "Outflow : 保險費",
                "Outflow : 模組回收費",
                "Outflow : 其它費用(人事薪資…等)",
                "1%管理費",
                "績效獎金",
                "Outflow : 所得稅(20%)"
            ],
            "capital": {
                "loan": "理財活動-借款",
                "equipment_cost": "投資活動-設備支出",
                "equity_investment": "理財活動-現金增資" # 直接讀取
            },
            "rent": "Outflow : 設備放置區域租金", # 用於敏感性分析調整
            "reduction": "年底減資",
            "year_row_label": "年份"
        }

        # --- 2. 尋找列索引 ---
        # 稅後淨利
        net_profit_row_idx = find_row_by_label(df, row_definitions["net_profit"])
        if net_profit_row_idx is None: return {"error": f"找不到列: '{row_definitions['net_profit']}'"}

        # 稅前淨利
        pre_tax_profit_row_idx = find_row_by_label(df, row_definitions["pre_tax_profit"])
        if pre_tax_profit_row_idx is None: return {"error": f"找不到列: '{row_definitions['pre_tax_profit']}'"}

        # 資本項目
        loan_row_idx = find_row_by_label(df, row_definitions["capital"]["loan"])
        equip_cost_row_idx = find_row_by_label(df, row_definitions["capital"]["equipment_cost"])
        equity_investment_row_idx = find_row_by_label(df, row_definitions["capital"]["equity_investment"])

        if loan_row_idx is None: return {"error": f"找不到列: '{row_definitions['capital']['loan']}'"}
        if equip_cost_row_idx is None: return {"error": f"找不到列: '{row_definitions['capital']['equipment_cost']}'"}
        if equity_investment_row_idx is None: return {"error": f"找不到列: '{row_definitions['capital']['equity_investment']}'"}

        # 租金 (用於調整)
        rent_row_idx = find_row_by_label(df, row_definitions["rent"])
        # 如果找不到租金列，我們假設租金為 0，不報錯 (因為可能真的沒有租金)
        
        # 減資
        capital_reduction_row_idx = find_row_by_label(df, row_definitions["reduction"])
        if capital_reduction_row_idx is None: return {"error": f"找不到列: '{row_definitions['reduction']}'"}

        # 專案法 IRR 項目
        project_irr_rows = {}
        for label in row_definitions["project_irr_items"]:
            row_idx = find_row_by_label(df, label)
            if row_idx is not None:
                project_irr_rows[label] = row_idx
            else:
                print(f"Warning: 找不到專案法 IRR 項目列 '{label}'，將在計算中忽略此項目 (視為 0)。")

        # 年份
        year_row_idx = find_row_by_label(df, row_definitions["year_row_label"])
        if year_row_idx is None: return {"error": f"找不到年份列: '{row_definitions['year_row_label']}'"}

        # --- 3. 確定數據範圍 ---
        row_years = df.iloc[year_row_idx]
        
        # 動態尋找起始年份所在的欄位
        start_col = -1
        for col in range(df.shape[1]):
            val = row_years.iloc[col]
            # 檢查是否為年份 (例如大於 1900 的數字)
            try:
                if pd.notna(val) and isinstance(val, (int, float)) and val > 1900:
                    start_col = col
                    break
            except:
                pass
        
        if start_col == -1:
            return {"error": "無法在年份列中找到起始年份 (需大於 1900)"}

        last_col = start_col
        for col in range(start_col + 1, df.shape[1]):
            if pd.isna(row_years.iloc[col]):
                break
            last_col = col
        
        num_years = last_col - start_col + 1

        # --- 4. 讀取基礎數據 ---
        
        # 讀取原始稅後淨利
        net_profit_row = df.iloc[net_profit_row_idx, start_col:last_col+1]
        base_net_profits = [float(x) if pd.notna(x) and x != '-' else 0.0 for x in net_profit_row]

        # 讀取原始稅前淨利
        pre_tax_profit_row = df.iloc[pre_tax_profit_row_idx, start_col:last_col+1]
        base_pre_tax_profits = [float(x) if pd.notna(x) and x != '-' else 0.0 for x in pre_tax_profit_row]

        # 讀取設備成本 (通常為負值，取絕對值)
        equip_row = df.iloc[equip_cost_row_idx, start_col:last_col+1]
        base_total_equip_cost = abs(sum([float(x) for x in equip_row if pd.notna(x) and x != '-']))
        
        # 讀取現金增資 (Capital Increase)
        equity_inv_row = df.iloc[equity_investment_row_idx, start_col:last_col+1]
        base_equity_investments = [float(x) if pd.notna(x) and x != '-' else 0.0 for x in equity_inv_row]

        # 讀取租金 (通常為負值)
        base_rents = [0.0] * num_years
        if rent_row_idx is not None:
            rent_row = df.iloc[rent_row_idx, start_col:last_col+1]
            base_rents = [float(x) if pd.notna(x) and x != '-' else 0.0 for x in rent_row]

        # 讀取減資
        capital_reduction_data = df.iloc[capital_reduction_row_idx, start_col:last_col+1]
        capital_reduction_values = [float(x) if pd.notna(x) and x != '-' else 0.0 for x in capital_reduction_data]

        # --- 5. 應用調整 (Sensitivity Analysis) ---
        
        # 調整設備成本
        adjusted_equip_cost = base_total_equip_cost * equipment_cost_adj
        
        # 計算借款金額
        if loan_amount is None:
            loan_row = df.iloc[loan_row_idx, start_col:last_col+1]
            total_loan = sum([float(x) for x in loan_row if pd.notna(x) and x != '-'])
        else:
            total_loan = loan_amount
            
        # 計算股東投入 (Equity Investment)
        # 股東投入 = 設備總成本 - 借款
        equity_investment = adjusted_equip_cost - total_loan
        
        # 計算折舊差異對淨利的影響
        # 假設折舊年限 20 年，直線法
        base_depreciation = base_total_equip_cost / 20.0 if base_total_equip_cost > 0 else 0
        new_depreciation = adjusted_equip_cost / 20.0 if adjusted_equip_cost > 0 else 0
        depreciation_diff = new_depreciation - base_depreciation # 正值代表折舊增加
        
        # 計算租金差異對淨利的影響
        # 租金為 Outflow (負值)，rent_adj > 1 代表成本增加 (更負)
        # 這裡假設 rent_adj 是乘數，例如 1.1 代表租金變為 1.1 倍
        adjusted_net_profits = []
        adjusted_pre_tax_profits = []
        
        for i in range(num_years):
            base_profit = base_net_profits[i]
            base_pre_tax = base_pre_tax_profits[i]
            base_rent = base_rents[i] # 負值
            
            new_rent = base_rent * rent_adj
            rent_diff = new_rent - base_rent # 負值 - 負值。若成本增加，new_rent 更小，diff 為負
            
            # 淨利調整 = 租金差異 - 折舊差異
            # (忽略稅盾效應以簡化計算，或假設稅率影響已包含在內)
            profit_adj = rent_diff - depreciation_diff
            
            adjusted_net_profits.append(base_profit + profit_adj)
            adjusted_pre_tax_profits.append(base_pre_tax + profit_adj)

        # --- 6. 計算現金流與 IRR ---
        
        # 成本法現金流 (Shareholder Perspective)
        # Year 0: -Equity Investment
        # Year 1+: Dividends + Capital Reduction
        cost_method_cash_flow = [0.0] * num_years
        
        # 假設第一年 (Index 0) 是投資年
        cost_method_cash_flow[0] = -equity_investment
        
        # 股利發放規則：前一年淨利 * 0.9 (若為正)
        dividends = [0.0] * num_years
        for i in range(1, num_years):
            prev_year_profit = adjusted_net_profits[i-1]
            if prev_year_profit > 0:
                dividends[i] = prev_year_profit * 0.9
            else:
                dividends[i] = 0
        
        for i in range(1, num_years):
            # 減資對股東來說是現金流入 (正值)，Excel 中若為負值 (公司流出)，需取絕對值加入
            capital_reduction = abs(capital_reduction_values[i])
            cost_method_cash_flow[i] = dividends[i] + capital_reduction

        try:
            cost_method_irr = npf.irr(cost_method_cash_flow)
        except Exception as e:
            cost_method_irr = f"無法計算: {e}"

        # 權益法現金流 (Equity Method)
        # 根據使用者定義：理財活動-現金增資()負數 + 稅後淨利
        # Year 0: -Equity Investment (Capital Increase)
        # Year 1+: Net Profit (稅後淨利)
        equity_method_cash_flow = [0.0] * num_years
        
        # Year 0: 初始投入 (現金增資)
        equity_method_cash_flow[0] = -equity_investment 
        
        # 如果 Year 0 也有淨利，也要加上
        if adjusted_net_profits[0] != 0:
             equity_method_cash_flow[0] += adjusted_net_profits[0]

        for i in range(1, num_years):
            # 權益法修正：稅後淨利 - 現金增資 + 年底減資 (正值)
            capital_reduction = abs(capital_reduction_values[i])
            capital_increase_val = base_equity_investments[i]
            
            # 使用者要求：稅後淨利 - 理財活動-現金增資 + 年底減資
            equity_method_cash_flow[i] = adjusted_net_profits[i] - capital_increase_val + capital_reduction

        try:
            equity_method_irr = npf.irr(equity_method_cash_flow)
        except Exception as e:
            equity_method_irr = f"無法計算: {e}"
        
        # 專案法現金流 (Project Cash Flow)
        # 直接加總所有 Inflow 和 Outflow (不含融資活動)
        project_cash_flow = [0.0] * num_years
        for i in range(num_years):
            year_sum = 0.0
            current_col = start_col + i
            for label, row_idx in project_irr_rows.items():
                val = df.iloc[row_idx, current_col]
                if pd.notna(val) and val != '-':
                    num_val = float(val)
                    # 應用敏感性分析調整
                    if label == "Outflow : 設備費用":
                        num_val *= equipment_cost_adj
                    elif label == "Outflow : 設備放置區域租金":
                        num_val *= rent_adj
                    year_sum += num_val
            project_cash_flow[i] = year_sum

        try:
            project_irr = npf.irr(project_cash_flow)
        except Exception as e:
            project_irr = f"無法計算: {e}"

        # --- 7. 準備返回結果 ---
        details = {
            "現金增資 (股東投入)": [-equity_investment] + [0]*(num_years-1),
            "理財活動-現金增資 (原始值)": base_equity_investments,
            "稅後淨利 (Net Profit)": adjusted_net_profits,
            "折舊 (Depreciation)": [new_depreciation] * num_years,
            "現金股利 (Dividends)": dividends,
            "年底減資 (Capital Reduction)": capital_reduction_values,
            "成本法現金流 (Cost Method CF)": cost_method_cash_flow,
            "專案法現金流 (Project CF)": project_cash_flow
        }
        
        mod_time = os.path.getmtime(EXCEL_PATH)
        timestamp_str = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')

        return {
            "cost_method_irr": cost_method_irr,
            "equity_method_irr": equity_method_irr,
            "project_irr": project_irr,
            "equity_method_cash_flow": equity_method_cash_flow,
            "project_cash_flow": project_cash_flow,
            "details": details,
            "last_modified": timestamp_str
        }

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return {"error": str(e)}

def get_project_items(excel_path=None, sheet_name=None):
    """
    回傳專案中可進行敏感性分析的項目列表。

    Args:
        excel_path (str, optional): 自訂 Excel 檔案路徑（目前未使用，保留供未來擴展）
        sheet_name (str, optional): 自訂工作表名稱（目前未使用，保留供未來擴展）
    """
    return [
        "Inflow : 電費收入",
        "Inflow : 政府補助",
        "Outflow : 設備費用",
        "Outflow : 設備放置區域租金",
        "Outflow : 維運費",
        "Outflow : 保險費",
        "Outflow : 模組回收費",
        "Outflow : 其它費用(人事薪資…等)",
        "1%管理費",
        "績效獎金",
        "所得稅20%"
    ]

def get_sensitivity_analysis(adjustment_amount=0, adjustment_percentage=None, target_item=None, excel_path=None, sheet_name=None):
    """
    執行敏感性分析：針對每個輸入/輸出項目，模擬變動後的 IRR。
    adjustment_amount: 固定金額調整 (正值=減少金額/改善損益, 負值=增加金額/惡化損益)
    adjustment_percentage: 百分比調整 (例如 0.05 為增加 5%, -0.05 為減少 5%)。若設定此值，將忽略 adjustment_amount。
    若指定 target_item，則只針對該項目進行模擬。
    excel_path (str, optional): 自訂 Excel 檔案路徑。若為 None，使用預設路徑。
    sheet_name (str, optional): 自訂工作表名稱。若為 None，使用預設名稱。
    """
    # 使用提供的路徑或預設路徑
    if excel_path is None:
        excel_path = DEFAULT_EXCEL_PATH
    if sheet_name is None:
        sheet_name = DEFAULT_SHEET_NAME

    if not os.path.exists(excel_path):
        return {"error": f"Excel file not found at {excel_path}"}

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        
        # 1. 定義項目
        row_definitions = {
            "net_profit": "稅後淨利",
            "pre_tax_profit": "稅前淨利",
            "project_irr_items": get_project_items(),
            "capital": {
                "loan": "理財活動-借款",
                "equipment_cost": "投資活動-設備支出",
                "equity_investment": "理財活動-現金增資"
            },
            "reduction": "年底減資",
            "year_row_label": "年份"
        }

        # 2. 尋找列索引 (與 get_irr_calculation_data 類似)
        indices = {}
        # 必須包含 "投資活動-設備支出" 以計算 equity_inv
        items_to_read = row_definitions["project_irr_items"] + [
            row_definitions["net_profit"], 
            row_definitions["pre_tax_profit"], 
            row_definitions["reduction"], 
            row_definitions["capital"]["equity_investment"],
            row_definitions["capital"]["equipment_cost"]
        ]
        
        for label in items_to_read:
            idx = find_row_by_label(df, label)
            if idx is not None:
                indices[label] = idx
        
        year_row_idx = find_row_by_label(df, row_definitions["year_row_label"])
        if year_row_idx is None: return {"error": "找不到年份列"}

        # 3. 確定數據範圍
        row_years = df.iloc[year_row_idx]
        start_col = -1
        for col in range(df.shape[1]):
            val = row_years.iloc[col]
            try:
                if pd.notna(val) and isinstance(val, (int, float)) and val > 1900:
                    start_col = col
                    break
            except: pass
        
        if start_col == -1: return {"error": "無法確定起始年份"}

        last_col = start_col
        for col in range(start_col + 1, df.shape[1]):
            if pd.isna(row_years.iloc[col]): break
            last_col = col
        num_years = last_col - start_col + 1

        # 4. 提取基礎數據 (Base Data)
        base_data = {}
        initial_data = {} # Year 0 data (start_col - 1) for Project IRR
        
        for label, idx in indices.items():
            row_vals = df.iloc[idx, start_col:last_col+1]
            base_data[label] = [float(x) if pd.notna(x) and x != '-' else 0.0 for x in row_vals]
            
            # 嘗試讀取 Year 0 (start_col - 1)
            if start_col > 0:
                val_0 = df.iloc[idx, start_col - 1]
                initial_data[label] = float(val_0) if pd.notna(val_0) and isinstance(val_0, (int, float)) else 0.0
            else:
                initial_data[label] = 0.0

        # 輔助函數：計算 IRR
        def calculate_irrs(current_data, current_initial_data):
            # 專案法
            # Project IRR 需要包含 initial_data (Year -1/Col 2) 作為初始投入
            proj_cf_0 = 0.0
            for item in row_definitions["project_irr_items"]:
                if item in current_initial_data:
                    proj_cf_0 += current_initial_data[item]
            
            # Project CF: [Year 0 (Initial)] + [Year 1-N (Current)]
            proj_cf = [proj_cf_0] + [0.0] * num_years
            for i in range(num_years):
                sum_val = 0.0
                for item in row_definitions["project_irr_items"]:
                    if item in current_data:
                        sum_val += current_data[item][i]
                proj_cf[i+1] = sum_val
            
            try: p_irr = npf.irr(proj_cf)
            except: p_irr = None

            # 成本法 & 權益法
            # 根據之前的驗證，權益法 IRR 在只使用 current_data (Col 3+) 時是正確的 (9.06%)
            # 所以這裡維持只使用 current_data 計算 equity_inv
            
            # 計算 equity_inv
            # 1. 設備費用 (Sum of all years in current_data)
            equip_cost_label = row_definitions["capital"]["equipment_cost"]
            current_equip_cost = 0.0
            
            # Add Year 0 Cost (Initial)
            if equip_cost_label in current_initial_data:
                current_equip_cost += abs(current_initial_data[equip_cost_label])
            
            # Add Years 1-N Cost
            if equip_cost_label in current_data:
                current_equip_cost += sum([abs(x) for x in current_data[equip_cost_label]])
            
            # 2. 借款 (Sum of all years)
            loan_idx = find_row_by_label(df, row_definitions["capital"]["loan"])
            loan_val = 0.0
            if loan_idx:
                loan_row_vals = df.iloc[loan_idx, start_col:last_col+1]
                loan_val = sum([float(x) for x in loan_row_vals if pd.notna(x) and x != '-'])
            
            equity_inv = current_equip_cost - loan_val
            
            # 成本法現金流
            cost_cf = [0.0] * num_years
            # Year 0: -Equity Investment
            cost_cf[0] = -equity_inv
            
            # 權益法現金流
            equity_cf = [0.0] * num_years
            # Year 0: -Equity Investment
            equity_cf[0] = -equity_inv
            # 如果 Year 0 有淨利，加上
            if row_definitions["net_profit"] in current_data:
                 equity_cf[0] += current_data[row_definitions["net_profit"]][0]

            dividends = [0.0] * num_years
            for i in range(num_years):
                # 股利計算
                np = current_data[row_definitions["net_profit"]][i]
                div = np * 0.9 if np > 0 else 0
                dividends[i] = div
                
                # 成本法 (Year 1+)
                if i > 0:
                    red = 0
                    if row_definitions["reduction"] in current_data:
                        red = abs(current_data[row_definitions["reduction"]][i])
                    cost_cf[i] = dividends[i] + red
                
                # 權益法 (Year 1+)
                if i > 0:
                    capital_reduction = 0
                    if row_definitions["reduction"] in current_data:
                        capital_reduction = abs(current_data[row_definitions["reduction"]][i])
                    
                    capital_increase_val = 0
                    if row_definitions["capital"]["equity_investment"] in current_data:
                        capital_increase_val = current_data[row_definitions["capital"]["equity_investment"]][i]

                    if row_definitions["net_profit"] in current_data:
                        # 使用者要求：稅後淨利 - 理財活動-現金增資 + 年底減資
                        equity_cf[i] = current_data[row_definitions["net_profit"]][i] - capital_increase_val + capital_reduction
                    else:
                        equity_cf[i] = 0 - capital_increase_val + capital_reduction

            try: c_irr = npf.irr(cost_cf)
            except: c_irr = None
            try: e_irr = npf.irr(equity_cf)
            except: e_irr = None

            return p_irr, c_irr, e_irr

        # 5. 計算基準 IRR
        base_p, base_c, base_e = calculate_irrs(base_data, initial_data)
        
        results = []
        
        # 決定要模擬的項目列表
        items_to_simulate = row_definitions["project_irr_items"]
        if target_item:
            # 嘗試精確匹配
            if target_item in items_to_simulate:
                items_to_simulate = [target_item]
            else:
                # 嘗試模糊匹配 (只要包含關鍵字)
                matches = [i for i in items_to_simulate if target_item in i]
                if matches:
                    items_to_simulate = matches
                # 若都沒對應到，則維持全部 (或可選擇回傳錯誤，這裡選擇維持全部但印出警告)
                else:
                    print(f"Warning: Target item '{target_item}' not found. Simulating all items.")

        # 6. 模擬每個項目的變動
        for item in items_to_simulate:
            if item not in base_data: continue
            
            # 複製數據
            sim_data = {k: v[:] for k, v in base_data.items()}
            sim_initial = initial_data.copy()
            
            # 準備要調整的項目列表 (處理連動)
            current_items_to_adjust = [item]
            if "設備費用" in item:
                # 連動調整 投資活動-設備支出
                cap_equip = row_definitions["capital"]["equipment_cost"]
                if cap_equip in base_data or cap_equip in initial_data:
                    current_items_to_adjust.append(cap_equip)

            profit_changes = [0.0] * num_years
            
            for adj_item in current_items_to_adjust:
                # 判斷是否為主要項目 (只有主要項目的變動會影響損益表，避免重複計算或誤算)
                # 投資活動-設備支出 不影響損益表 (它是資產負債/現金流項目)，只影響 Equity Investment 計算
                is_main_item = (adj_item == item)

                # 調整 Year 0 (Initial)
                if adj_item in sim_initial:
                    original_val = sim_initial[adj_item]
                    if original_val != 0:
                        if adjustment_percentage is not None:
                            new_val = original_val * (1 + adjustment_percentage)
                            sim_initial[adj_item] = new_val
                        else:
                            if original_val > 0: # Inflow
                                new_val = max(0, original_val - adjustment_amount)
                                sim_initial[adj_item] = new_val
                            else: # Outflow
                                new_val = min(0, original_val + adjustment_amount)
                                sim_initial[adj_item] = new_val
                
                # 調整 Years 1-N
                if adj_item in sim_data:
                    for i in range(num_years):
                        original_val = sim_data[adj_item][i]
                        if original_val == 0: continue 
                        
                        if adjustment_percentage is not None:
                            new_val = original_val * (1 + adjustment_percentage)
                            diff = new_val - original_val
                            sim_data[adj_item][i] = new_val
                            if is_main_item:
                                profit_changes[i] += diff
                        else:
                            if original_val > 0: # Inflow
                                new_val = max(0, original_val - adjustment_amount)
                                diff = new_val - original_val # 負值
                                sim_data[adj_item][i] = new_val
                                if is_main_item:
                                    profit_changes[i] += diff
                            else: # Outflow (Negative)
                                new_val = min(0, original_val + adjustment_amount)
                                diff = new_val - original_val # 正值
                                sim_data[adj_item][i] = new_val
                                if is_main_item:
                                    profit_changes[i] += diff
            
            # 同步調整稅後淨利 (Net Profit) 和 稅前淨利 (Pre-Tax Profit)
            for i in range(num_years):
                sim_data[row_definitions["net_profit"]][i] += profit_changes[i]
                if row_definitions["pre_tax_profit"] in sim_data:
                    sim_data[row_definitions["pre_tax_profit"]][i] += profit_changes[i]

            # 重新計算 IRR
            new_p, new_c, new_e = calculate_irrs(sim_data, sim_initial)
            
            results.append({
                "item": item,
                "p_irr": new_p,
                "c_irr": new_c,
                "e_irr": new_e,
                "p_diff": (new_p - base_p) if new_p and base_p else 0,
                "c_diff": (new_c - base_c) if new_c and base_c else 0,
                "e_diff": (new_e - base_e) if new_e and base_e else 0
            })

        return {
            "base": {"p_irr": base_p, "c_irr": base_c, "e_irr": base_e},
            "simulations": results
        }

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return {"error": str(e)}

def run_iterative_simulation(target_item, step_amount, target_irr_type, target_value, condition_operator='>', excel_path=None, sheet_name=None):
    """
    執行迭代模擬：每次調整 step_amount，直到 IRR 滿足條件。
    step_amount: 負值代表增加金額(變更負或更多)，正值代表減少金額。
    target_irr_type: 'e_irr' (Equity), 'c_irr' (Cost), 'p_irr' (Project)
    target_value: 目標 IRR 值 (例如 0.07)
    condition_operator: '>' 或 '<'
    excel_path (str, optional): 自訂 Excel 檔案路徑。若為 None，使用預設路徑。
    sheet_name (str, optional): 自訂工作表名稱。若為 None，使用預設名稱。
    """
    # 使用提供的路徑或預設路徑
    if excel_path is None:
        excel_path = DEFAULT_EXCEL_PATH
    if sheet_name is None:
        sheet_name = DEFAULT_SHEET_NAME

    if not os.path.exists(excel_path):
        return {"error": f"Excel file not found at {excel_path}"}

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        
        # --- 1. 定義項目 (與 get_sensitivity_analysis 相同) ---
        row_definitions = {
            "net_profit": "稅後淨利",
            "pre_tax_profit": "稅前淨利",
            "project_irr_items": get_project_items(),
            "capital": {
                "loan": "理財活動-借款",
                "equipment_cost": "投資活動-設備支出",
                "equity_investment": "理財活動-現金增資"
            },
            "reduction": "年底減資",
            "year_row_label": "年份"
        }

        # --- 2. 尋找列索引 ---
        indices = {}
        items_to_read = row_definitions["project_irr_items"] + [
            row_definitions["net_profit"], 
            row_definitions["pre_tax_profit"], 
            row_definitions["reduction"], 
            row_definitions["capital"]["equity_investment"],
            row_definitions["capital"]["equipment_cost"]
        ]
        
        for label in items_to_read:
            idx = find_row_by_label(df, label)
            if idx is not None:
                indices[label] = idx
        
        year_row_idx = find_row_by_label(df, row_definitions["year_row_label"])
        if year_row_idx is None: return {"error": "找不到年份列"}

        # --- 3. 確定數據範圍 ---
        row_years = df.iloc[year_row_idx]
        start_col = -1
        for col in range(df.shape[1]):
            val = row_years.iloc[col]
            try:
                if pd.notna(val) and isinstance(val, (int, float)) and val > 1900:
                    start_col = col
                    break
            except: pass
        
        if start_col == -1: return {"error": "無法確定起始年份"}

        last_col = start_col
        for col in range(start_col + 1, df.shape[1]):
            if pd.isna(row_years.iloc[col]): break
            last_col = col
        num_years = last_col - start_col + 1

        # --- 4. 提取基礎數據 ---
        base_data = {}
        initial_data = {} 
        
        for label, idx in indices.items():
            row_vals = df.iloc[idx, start_col:last_col+1]
            base_data[label] = [float(x) if pd.notna(x) and x != '-' else 0.0 for x in row_vals]
            
            if start_col > 0:
                val_0 = df.iloc[idx, start_col - 1]
                initial_data[label] = float(val_0) if pd.notna(val_0) and isinstance(val_0, (int, float)) else 0.0
            else:
                initial_data[label] = 0.0

        # --- 5. 定義計算函數 (內嵌) ---
        def calculate_irrs(current_data, current_initial_data):
            # Project IRR
            proj_cf_0 = 0.0
            for item in row_definitions["project_irr_items"]:
                if item in current_initial_data:
                    proj_cf_0 += current_initial_data[item]
            
            proj_cf = [proj_cf_0] + [0.0] * num_years
            for i in range(num_years):
                sum_val = 0.0
                for item in row_definitions["project_irr_items"]:
                    if item in current_data:
                        sum_val += current_data[item][i]
                proj_cf[i+1] = sum_val
            
            try: p_irr = npf.irr(proj_cf)
            except: p_irr = None

            # Equity Investment Calculation
            equip_cost_label = row_definitions["capital"]["equipment_cost"]
            current_equip_cost = 0.0
            
            # Add Year 0 Cost (Initial)
            if equip_cost_label in current_initial_data:
                current_equip_cost += abs(current_initial_data[equip_cost_label])
            
            # Add Years 1-N Cost
            if equip_cost_label in current_data:
                current_equip_cost += sum([abs(x) for x in current_data[equip_cost_label]])
            
            loan_idx = find_row_by_label(df, row_definitions["capital"]["loan"])
            loan_val = 0.0
            if loan_idx:
                loan_row_vals = df.iloc[loan_idx, start_col:last_col+1]
                loan_val = sum([float(x) for x in loan_row_vals if pd.notna(x) and x != '-'])
            
            equity_inv = current_equip_cost - loan_val
            
            # Cost Method CF
            cost_cf = [0.0] * num_years
            cost_cf[0] = -equity_inv
            
            # Equity Method CF
            equity_cf = [0.0] * num_years
            equity_cf[0] = -equity_inv
            if row_definitions["net_profit"] in current_data:
                 equity_cf[0] += current_data[row_definitions["net_profit"]][0]

            dividends = [0.0] * num_years
            for i in range(num_years):
                np = current_data[row_definitions["net_profit"]][i]
                div = np * 0.9 if np > 0 else 0
                dividends[i] = div
                
                if i > 0:
                    red = 0
                    if row_definitions["reduction"] in current_data:
                        red = abs(current_data[row_definitions["reduction"]][i])
                    cost_cf[i] = dividends[i] + red
                
                if i > 0:
                    capital_reduction = 0
                    if row_definitions["reduction"] in current_data:
                        capital_reduction = abs(current_data[row_definitions["reduction"]][i])
                    
                    capital_increase_val = 0
                    if row_definitions["capital"]["equity_investment"] in current_data:
                        capital_increase_val = current_data[row_definitions["capital"]["equity_investment"]][i]

                    if row_definitions["net_profit"] in current_data:
                        equity_cf[i] = current_data[row_definitions["net_profit"]][i] - capital_increase_val + capital_reduction
                    else:
                        equity_cf[i] = 0 - capital_increase_val + capital_reduction

            try: c_irr = npf.irr(cost_cf)
            except: c_irr = None
            try: e_irr = npf.irr(equity_cf)
            except: e_irr = None

            return p_irr, c_irr, e_irr

        # --- 6. 執行迭代 ---
        
        # 模糊匹配 target_item
        matched_item = None
        for item in row_definitions["project_irr_items"]:
            if target_item in item:
                matched_item = item
                break
        
        if not matched_item:
            return {"error": f"找不到目標項目: {target_item}"}
            
        # 特殊處理：若調整「設備費用」，需同步調整「投資活動-設備支出」以連動權益法 IRR
        items_to_adjust = [matched_item]
        if "設備費用" in matched_item:
            items_to_adjust.append(row_definitions["capital"]["equipment_cost"])

        history = []
        current_adjustment = 0.0
        max_iter = 50 # 增加迭代次數上限
        
        for i in range(max_iter + 1):
            # 複製數據
            sim_data = {k: v[:] for k, v in base_data.items()}
            sim_initial = initial_data.copy()
            
            profit_change_per_year = 0.0
            
            # 應用調整 (支援多個連動項目)
            for item_name in items_to_adjust:
                # 確保項目存在於數據中
                if item_name not in sim_data and item_name not in sim_initial:
                    continue

                # 調整 Year 0
                if item_name in sim_initial:
                    original_val = sim_initial[item_name]
                    if original_val != 0:
                        if original_val > 0: # Inflow
                            new_val = max(0, original_val + current_adjustment)
                        else: # Outflow
                            new_val = min(0, original_val + current_adjustment)
                            sim_initial[item_name] = new_val
                
                # 調整 Years 1-N
                if item_name in sim_data:
                    for y in range(num_years):
                        original_val = sim_data[item_name][y]
                        if original_val == 0: continue
                        
                        if original_val > 0: # Inflow
                            new_val = max(0, original_val + current_adjustment)
                            diff = new_val - original_val
                            sim_data[item_name][y] = new_val
                            # 只有主要項目 (matched_item) 的變動才計入損益表 (避免重複計算)
                            if item_name == matched_item:
                                profit_change_per_year = diff
                        else: # Outflow
                            new_val = min(0, original_val + current_adjustment)
                            diff = new_val - original_val
                            sim_data[item_name][y] = new_val
                            # 只有主要項目 (matched_item) 的變動才計入損益表
                            if item_name == matched_item:
                                profit_change_per_year = diff
            
            # 更新淨利
            for y in range(num_years):
                sim_data[row_definitions["net_profit"]][y] += profit_change_per_year
            
            # 計算 IRR
            p_irr, c_irr, e_irr = calculate_irrs(sim_data, sim_initial)
            
            current_irr_val = locals().get(target_irr_type)
            
            # 記錄
            history.append({
                "iteration": i,
                "adjustment": current_adjustment,
                "irr": current_irr_val,
                "p_irr": p_irr,
                "c_irr": c_irr,
                "e_irr": e_irr
            })
            
            # 檢查條件
            if current_irr_val is not None:
                if condition_operator == '>' and current_irr_val > target_value:
                    return {
                        "success": True, 
                        "message": f"在第 {i} 次迭代滿足條件 (調整 {current_adjustment:,.0f})", 
                        "history": history,
                        "final_adjustment": current_adjustment,
                        "final_irr": current_irr_val
                    }
                elif condition_operator == '<' and current_irr_val < target_value:
                    return {
                        "success": True, 
                        "message": f"在第 {i} 次迭代滿足條件 (調整 {current_adjustment:,.0f})", 
                        "history": history,
                        "final_adjustment": current_adjustment,
                        "final_irr": current_irr_val
                    }
            
            # 準備下一次迭代
            current_adjustment += step_amount
            
        return {
            "success": False, 
            "message": f"達到最大迭代次數 ({max_iter}) 仍未滿足條件", 
            "history": history
        }

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return {"error": str(e)}

if __name__ == "__main__":
    # 用於直接執行此文件進行測試
    result = get_irr_calculation_data()
    import json
    print(json.dumps(result, indent=2, ensure_ascii=False))