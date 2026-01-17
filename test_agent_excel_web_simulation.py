"""
模擬 AGENT 使用 Excel 數據與網頁搜尋值調用工具

此測試展示 AGENT 如何：
1. 從 Excel 讀取原始數據
2. 模擬從網頁獲取參數（如即時電價、利率等）
3. 結合這些數據調用計算工具
4. 將結果寫回 Excel
"""

import json
from tool.tool_manager import ToolManager
from tool.excel_tool import ExcelTool

# ============ 初始化 ============
print("=" * 80)
print("🤖 AGENT 模擬：結合 Excel 數據與網頁資料進行工具調用")
print("=" * 80)
print()

tool_manager = ToolManager()

# ============ 模擬網頁搜尋結果 ============
class WebSearchSimulator:
    """模擬從網頁獲取即時數據"""
    
    @staticmethod
    def get_current_electricity_price():
        """模擬搜尋台電目前躉售電價"""
        return {
            "source": "台灣電力公司官網",
            "date": "2026-01-18",
            "fit_rate": 4.1272,  # 太陽光電躉售費率
            "unit": "元/度",
            "category": "第二型 (高效能)"
        }
    
    @staticmethod
    def get_current_bank_rate():
        """模擬搜尋目前銀行貸款利率"""
        return {
            "source": "中央銀行",
            "date": "2026-01-18",
            "base_rate": 1.875,
            "solar_loan_rate": 2.5,  # 太陽能專案貸款利率
            "unit": "%"
        }
    
    @staticmethod
    def get_equipment_market_price():
        """模擬搜尋市場設備價格"""
        return {
            "source": "太陽光電產業報價",
            "date": "2026-01-18",
            "solar_panel_price": 22000,  # 每 kWp 模組價格
            "inverter_price": 3000,      # 每 kWp 逆變器價格
            "installation_cost": 2000,   # 每 kWp 安裝費
            "total_avg_price": 27000,    # 市場平均總價
            "unit": "元/kWp"
        }

    @staticmethod
    def get_target_irr_requirement():
        """模擬獲取投資人目標 IRR 要求"""
        return {
            "source": "投資評估準則",
            "min_project_irr": 6.0,
            "target_project_irr": 7.5,
            "min_equity_irr": 8.0,
            "unit": "%"
        }


# ============ 場景 1: 讀取 Excel 數據 + 網頁電價 ============
print("📌 場景 1: AGENT 讀取 Excel 基礎數據並結合網頁電價資訊")
print("-" * 80)
print()

print("👤 用戶: 「幫我讀取目前專案的基本資料，並查詢最新的電價」")
print()

print("🤖 AGENT 思考:")
print("   1. 需要從 Excel 讀取專案基礎數據")
print("   2. 需要從網頁搜尋最新電價資訊")
print("   3. 比對兩者差異並提供建議")
print()

# AGENT 執行動作 1: 讀取 Excel 數據
print("🔧 AGENT 調用工具: read_excel_cell")
excel_results = {}

# 讀取專案名稱
result = tool_manager.execute_tool('read_excel_cell', {'cell': 'B2'})
excel_results['project_name'] = result.get('value', '未知專案')
print(f"   讀取 B2 (專案名稱): {excel_results['project_name']}")

# 讀取建置容量
result = tool_manager.execute_tool('read_excel_cell', {'cell': 'B4'})
excel_results['capacity'] = result.get('value', 0)
print(f"   讀取 B4 (建置容量): {excel_results['capacity']} kWp")

# 讀取目前設備成本
result = tool_manager.execute_tool('read_excel_cell', {'cell': 'C16'})
excel_results['equipment_cost'] = result.get('value', 0)
print(f"   讀取 C16 (設備成本): {excel_results['equipment_cost']} 元/kWp")

# 讀取目前躉售費率
result = tool_manager.execute_tool('read_excel_cell', {'cell': 'C13'})
excel_results['current_fit'] = result.get('value', 0)
print(f"   讀取 C13 (躉售費率): {excel_results['current_fit']} 元/度")

print()

# AGENT 執行動作 2: 搜尋網頁資料
print("🌐 AGENT 搜尋網頁: 查詢最新電價資訊")
web_price = WebSearchSimulator.get_current_electricity_price()
print(f"   來源: {web_price['source']}")
print(f"   日期: {web_price['date']}")
print(f"   最新躉售費率: {web_price['fit_rate']} 元/度")
print()

# AGENT 整合資訊回覆
print("🤖 AGENT 回應:")
print(f"""
📊 **專案基礎資料概覽**

| 項目 | Excel 數據 | 最新網頁資料 |
|------|-----------|-------------|
| 專案名稱 | {excel_results['project_name']} | - |
| 建置容量 | {excel_results['capacity']} kWp | - |
| 設備成本 | {excel_results['equipment_cost']} 元/kWp | - |
| 躉售費率 | {excel_results['current_fit']} 元/度 | {web_price['fit_rate']} 元/度 |

💡 **建議**: 
- 目前 Excel 中的躉售費率為 {excel_results['current_fit']}，最新公告費率為 {web_price['fit_rate']}
- 建議更新 Excel 中的費率以反映最新數據
""")

print()
print()

# ============ 場景 2: 網頁參數 + Excel 基礎 進行 IRR 計算 ============
print("📌 場景 2: AGENT 使用網頁市場價格進行價金滾算分析")
print("-" * 80)
print()

print("👤 用戶: 「查一下現在市場設備價格，然後分析如果用市場價來算 IRR 會是多少」")
print()

print("🤖 AGENT 思考:")
print("   1. 搜尋目前市場設備價格")
print("   2. 使用 calculate_price_rolling 工具進行分析")
print("   3. 比較不同價格區間的 IRR")
print()

# AGENT 搜尋網頁
print("🌐 AGENT 搜尋網頁: 查詢市場設備價格")
market_price = WebSearchSimulator.get_equipment_market_price()
print(f"   來源: {market_price['source']}")
print(f"   市場平均總價: {market_price['total_avg_price']} 元/kWp")
print(f"   模組價格: {market_price['solar_panel_price']} 元/kWp")
print(f"   逆變器: {market_price['inverter_price']} 元/kWp")
print(f"   安裝費: {market_price['installation_cost']} 元/kWp")
print()

# 使用市場價格作為邊界進行滾算
print("🔧 AGENT 調用工具: calculate_price_rolling")
print(f"   參數: mode='CashMode', boundary={market_price['total_avg_price']}, step=500")
print()

result = tool_manager.execute_tool(
    'calculate_price_rolling',
    {
        'mode': 'CashMode',
        'boundary': market_price['total_avg_price'],
        'step': 500,
        'profit_rate': 0.20
    }
)

if result.get('success'):
    summary = result.get('results_summary', {})
    data_rows = summary.get('data', [])
    
    print("🤖 AGENT 回應:")
    print(f"""
📊 **價金滾算分析結果**（基於市場價格 {market_price['total_avg_price']} 元/kWp）

從 Excel 初始設備成本 {excel_results['equipment_cost']} 元/kWp 到市場價 {market_price['total_avg_price']} 元/kWp：
""")
    
    if len(data_rows) >= 2:
        print(f"""
| 設備成本 | 含利潤價格 | 專案法IRR | 成本法IRR | 權益法IRR |
|----------|-----------|----------|----------|----------|""")
        
        # 顯示首尾對比
        first = data_rows[0]
        last = data_rows[-1]
        print(f"| {first[0]:,} | {first[2]:,} | {first[3]:.2f}% | {first[4]:.2f}% | {first[5]:.2f}% |")
        if len(data_rows) > 2:
            mid = data_rows[len(data_rows)//2]
            print(f"| {mid[0]:,} | {mid[2]:,} | {mid[3]:.2f}% | {mid[4]:.2f}% | {mid[5]:.2f}% |")
        print(f"| {last[0]:,} | {last[2]:,} | {last[3]:.2f}% | {last[4]:.2f}% | {last[5]:.2f}% |")
        
        print(f"""
💡 **分析結論**:
- 設備成本從 {first[0]:,} 降至 {last[0]:,} 元/kWp
- 專案法 IRR 可從 {first[3]:.2f}% 提升至 {last[3]:.2f}%（提升 {last[3]-first[3]:.2f}%）
- 若採用市場均價 {market_price['total_avg_price']:,} 元/kWp，IRR 表現可觀
""")

print()
print()

# ============ 場景 3: 結合目標 IRR 進行反推 ============
print("📌 場景 3: AGENT 根據投資人目標 IRR 反推合理價金")
print("-" * 80)
print()

print("👤 用戶: 「投資人要求專案法 IRR 至少 7%，設備成本要壓到多少？」")
print()

print("🤖 AGENT 思考:")
print("   1. 搜尋投資人目標 IRR 標準")
print("   2. 使用工具進行滾算")
print("   3. 找出符合條件的價金範圍")
print()

# AGENT 搜尋目標
print("🌐 AGENT 搜尋網頁: 查詢投資評估標準")
irr_requirement = WebSearchSimulator.get_target_irr_requirement()
print(f"   來源: {irr_requirement['source']}")
print(f"   最低專案法 IRR: {irr_requirement['min_project_irr']}%")
print(f"   目標專案法 IRR: {irr_requirement['target_project_irr']}%")
print(f"   最低權益法 IRR: {irr_requirement['min_equity_irr']}%")
print()

# 進行滾算尋找目標
print("🔧 AGENT 調用工具: calculate_price_rolling (大範圍搜尋)")
result = tool_manager.execute_tool(
    'calculate_price_rolling',
    {
        'mode': 'CashMode',
        'boundary': 18000,  # 搜尋更低的價格
        'step': 500,
        'profit_rate': 0.20
    }
)

if result.get('success'):
    data_rows = result.get('results_summary', {}).get('data', [])
    
    # 尋找符合條件的價金
    target_irr = irr_requirement['min_project_irr']  # 7%
    qualified_prices = []
    
    for row in data_rows:
        if row[3] >= target_irr:
            qualified_prices.append(row)
    
    print("🤖 AGENT 回應:")
    print(f"""
📊 **目標 IRR 反推分析**

用戶要求: 專案法 IRR ≥ {target_irr}%
""")
    
    if qualified_prices:
        first_qualified = qualified_prices[0]
        print(f"""
✅ **符合條件的設備成本範圍**:

| 設備成本 | 含利潤價格 | 專案法IRR | 成本法IRR | 權益法IRR | 狀態 |
|----------|-----------|----------|----------|----------|------|""")
        
        for i, row in enumerate(qualified_prices[:5]):  # 只顯示前5個
            status = "⭐ 推薦" if i == 0 else "✓ 符合"
            print(f"| {row[0]:,} | {row[2]:,} | {row[3]:.2f}% | {row[4]:.2f}% | {row[5]:.2f}% | {status} |")
        
        print(f"""
💡 **建議**:
- 設備成本需降至 **{first_qualified[0]:,} 元/kWp** 以下才能達到 {target_irr}% IRR
- 最佳推薦價格: {first_qualified[0]:,} 元/kWp（IRR: {first_qualified[3]:.2f}%）
- 與目前成本 {excel_results['equipment_cost']} 相比，需降價 {excel_results['equipment_cost'] - first_qualified[0]:,} 元/kWp
""")
    else:
        print(f"""
❌ 在當前設定下，無法達到 {target_irr}% 的專案法 IRR

💡 **建議**:
- 考慮調整利潤率
- 或與供應商議價取得更低的設備成本
""")

print()
print()

# ============ 場景 4: 完整流程 - 讀取 Excel + 網頁參數 + 計算 + 寫回 ============
print("📌 場景 4: AGENT 執行完整流程並保存記錄")
print("-" * 80)
print()

print("👤 用戶: 「用剛才分析的結果執行滾算，並把記錄保存下來」")
print()

print("🤖 AGENT 思考:")
print("   1. 使用先前分析的最佳參數")
print("   2. 調用 execute_price_rolling 執行完整滾算")
print("   3. 工具會自動保存 Excel 記錄")
print()

# 使用找到的最佳參數
if qualified_prices:
    target_price = qualified_prices[0][0]
else:
    target_price = 22000

print(f"🔧 AGENT 調用工具: execute_price_rolling")
print(f"   參數: mode='cash', boundary={target_price}, step=500, profit_rate=0.20")
print()

result = tool_manager.execute_tool(
    'execute_price_rolling',
    {
        'mode': 'cash',
        'boundary': target_price,
        'step': 500,
        'profit_rate': 0.20
    }
)

if result.get('success'):
    summary = result.get('summary', {})
    output_file = result.get('output_file', '')
    
    print("🤖 AGENT 回應:")
    print(f"""
✅ **價金滾算完成**

📁 **記錄檔案**: {output_file}

📊 **滾算摘要**:
| 項目 | 數值 |
|------|------|
| 初始設備成本 | {summary.get('initial_cost', 'N/A')} 元/kWp |
| 最終設備成本 | {summary.get('final_cost', 'N/A')} 元/kWp |
| 調整次數 | {summary.get('adjustment_count', 'N/A')} 次 |
| 總降幅 | {summary.get('total_reduction', 'N/A')} 元/kWp |

🎯 **整合網頁資料**:
- 市場參考價: {market_price['total_avg_price']:,} 元/kWp
- 投資人最低 IRR 要求: {irr_requirement['min_project_irr']}%
- 最新躉售費率: {web_price['fit_rate']} 元/度

所有滾算結果已保存，可隨時查閱！
""")
else:
    print(f"❌ 滾算執行失敗: {result.get('message', '未知錯誤')}")

print()
print()

# ============ 場景 5: 銀行利率查詢 + 貸款方案分析 ============
print("📌 場景 5: AGENT 結合即時利率進行財務分析")
print("-" * 80)
print()

print("👤 用戶: 「查一下目前銀行貸款利率，看看對 IRR 有什麼影響」")
print()

print("🤖 AGENT 思考:")
print("   1. 搜尋最新銀行利率")
print("   2. 讀取 Excel 中目前設定的利率")
print("   3. 比較差異並提供建議")
print()

# 搜尋利率
print("🌐 AGENT 搜尋網頁: 查詢銀行貸款利率")
bank_rate = WebSearchSimulator.get_current_bank_rate()
print(f"   來源: {bank_rate['source']}")
print(f"   基準利率: {bank_rate['base_rate']}%")
print(f"   太陽能專案利率: {bank_rate['solar_loan_rate']}%")
print()

# 讀取 Excel 利率
print("🔧 AGENT 調用工具: read_excel_cell")
result = tool_manager.execute_tool('read_excel_cell', {'cell': 'C25'})
excel_rate = result.get('value', 0)
print(f"   讀取 C25 (銀行利率): {excel_rate}%")
print()

print("🤖 AGENT 回應:")
print(f"""
📊 **利率比較分析**

| 項目 | Excel 設定 | 最新網頁資料 | 差異 |
|------|-----------|-------------|------|
| 銀行利率 | {excel_rate}% | {bank_rate['solar_loan_rate']}% | {bank_rate['solar_loan_rate'] - (excel_rate if excel_rate else 0):.2f}% |

💡 **影響分析**:
- 利率每增加 0.5%，年利息支出約增加設備總額的 0.5%
- 以 {excel_results['capacity']} kWp、{excel_results['equipment_cost']} 元/kWp 計算
- 總設備費用: {excel_results['capacity'] * excel_results['equipment_cost']:,.0f} 元
- 利率差異影響: 約 {excel_results['capacity'] * excel_results['equipment_cost'] * 0.005:,.0f} 元/年

📝 **建議**:
- 若利率有變動，建議更新 Excel 中的利率設定後重新計算 IRR
""")

print()
print()

# ============ 測試總結 ============
print("=" * 80)
print("📋 測試總結")
print("=" * 80)
print()
print("✅ AGENT 成功展示的能力:")
print()
print("   1. 📑 **Excel 數據讀取**")
print("      - 使用 read_excel_cell 工具讀取專案基礎數據")
print("      - 支援單一儲存格和範圍讀取")
print()
print("   2. 🌐 **網頁資料整合**")
print("      - 模擬搜尋即時電價、利率、市場價格")
print("      - 與 Excel 數據進行比對分析")
print()
print("   3. 🔧 **工具智能調用**")
print("      - calculate_price_rolling: 純計算分析")
print("      - execute_price_rolling: 執行並保存記錄")
print("      - 根據場景選擇合適的工具")
print()
print("   4. 📊 **數據分析與建議**")
print("      - 目標 IRR 反推設備成本")
print("      - 比較不同情境的影響")
print("      - 提供具體可行的建議")
print()
print("   5. 💾 **結果保存**")
print("      - 自動輸出 Excel 滾算記錄")
print("      - 整合所有分析數據")
print()
print("=" * 80)
print("🎉 模擬測試完成！")
print("=" * 80)
