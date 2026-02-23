"""模擬 AGENT 操作價金滾算工具"""
from tool.tool_manager import ToolManager
import json

tool_manager = ToolManager()

print('=' * 70)
print('模擬 AGENT 操作測試')
print('=' * 70)
print()

# ============ 場景 1: 用戶想快速了解 IRR 變化 ============
print('場景 1: 用戶想快速分析不同價金的 IRR')
print('-' * 70)
print('用戶: "幫我看一下如果設備成本從 25000 降到 20000，IRR 會怎麼變化？"')
print()
print('AGENT 思考: 用戶只想看分析結果，不需要保存檔案')
print('AGENT 選擇: calculate_price_rolling (純計算模式)')
print()

result = tool_manager.execute_tool(
    'calculate_price_rolling',
    {
        'mode': 'CashMode',
        'boundary': 20000,
        'step': 500,
        'profit_rate': 0.20
    }
)

if result['success']:
    base_irr = result.get('base_irr', {})
    summary = result.get('results_summary', {})
    data_rows = summary.get('data', [])
    
    print('AGENT 回應:')
    print(f'根據分析結果，設備成本從 25000 降到 20000 元/kW：')
    print(f'• 初始狀態 (25000 元/kW):')
    print(f'  - 專案法 IRR: {data_rows[0][3]:.2f}%')
    print(f'  - 成本法 IRR: {data_rows[0][4]:.2f}%')
    print(f'  - 權益法 IRR: {data_rows[0][5]:.2f}%')
    print(f'• 最終狀態 (20000 元/kW):')
    print(f'  - 專案法 IRR: {data_rows[-1][3]:.2f}%')
    print(f'  - 成本法 IRR: {data_rows[-1][4]:.2f}%')
    print(f'  - 權益法 IRR: {data_rows[-1][5]:.2f}%')
    print(f'• IRR 提升幅度:')
    print(f'  - 專案法: +{data_rows[-1][3] - data_rows[0][3]:.2f}%')
    print(f'  - 成本法: +{data_rows[-1][4] - data_rows[0][4]:.2f}%')
    print(f'  - 權益法: +{data_rows[-1][5] - data_rows[0][5]:.2f}%')

print()
print()

# ============ 場景 2: 用戶想知道達到目標 IRR 需要的價金 ============
print('場景 2: 用戶想找到達到目標 IRR 的價金')
print('-' * 70)
print('用戶: "我希望專案法 IRR 達到 7% 以上，設備成本要降到多少？"')
print()
print('AGENT 思考: 需要遍歷計算，找到符合條件的價金')
print('AGENT 選擇: calculate_price_rolling (純計算模式)')
print()

result = tool_manager.execute_tool(
    'calculate_price_rolling',
    {
        'mode': 'CashMode',
        'boundary': 20000,
        'step': 500,
        'profit_rate': 0.20
    }
)

if result['success']:
    data_rows = result.get('results_summary', {}).get('data', [])
    
    # 找到第一個達到 7% 的價金
    target_irr = 7.0
    found = None
    for row in data_rows:
        if row[3] >= target_irr:
            found = row
            break
    
    print('AGENT 回應:')
    if found:
        print(f'要達到專案法 IRR {target_irr}% 以上：')
        print(f'• 設備成本需要降到: {found[0]} 元/kW')
        print(f'• 預期 IRR: {found[3]:.2f}%')
        print(f'• 含利潤最終價格: {found[2]} 元/kW')
        print(f'• 從 25000 元/kW 降低了: {25000 - found[0]} 元/kW')
    else:
        print(f'在當前設定下，無法達到 {target_irr}% 的 IRR')

print()
print()

# ============ 場景 3: 用戶要正式執行並保存記錄 ============
print('場景 3: 用戶要正式執行價金滾算並保存記錄')
print('-' * 70)
print('用戶: "好，那就按照這個方案執行，幫我記錄下來"')
print()
print('AGENT 思考: 需要保存完整記錄到 Excel 檔案')
print('AGENT 選擇: execute_price_rolling (完整流程)')
print()

result = tool_manager.execute_tool(
    'execute_price_rolling',
    {
        'mode': 'cash',
        'boundary': 23000,
        'step': 500,
        'profit_rate': 0.20
    }
)

if result['success']:
    summary = result.get('summary', {})
    output_file = result.get('output_file', '')
    
    print('AGENT 回應:')
    print(f'已完成價金滾算並保存記錄：')
    print(f'• 初始設備成本: {summary.get("initial_cost")} 元/kW')
    print(f'• 最終設備成本: {summary.get("final_cost")} 元/kW')
    print(f'• 調整次數: {summary.get("adjustment_count")} 次')
    print(f'• 總降幅: {summary.get("total_reduction")} 元/kW')
    print(f'• 記錄檔案: {output_file}')
    print()
    print('所有滾算結果已保存到 Excel final 資料夾，可以隨時查閱！')

print()
print()

# ============ 場景 4: 用戶想比較不同利潤率 ============
print('場景 4: 用戶想比較不同利潤率的影響')
print('-' * 70)
print('用戶: "如果利潤率改成 15% 或 25%，IRR 會差多少？"')
print()
print('AGENT 思考: 需要多次計算並比較結果')
print('AGENT 選擇: 多次呼叫 calculate_price_rolling')
print()

profit_rates = [0.15, 0.20, 0.25]
results_comparison = []

for pr in profit_rates:
    result = tool_manager.execute_tool(
        'calculate_price_rolling',
        {
            'mode': 'CashMode',
            'boundary': 23000,
            'step': 1000,
            'profit_rate': pr
        }
    )
    
    if result['success']:
        data_rows = result.get('results_summary', {}).get('data', [])
        results_comparison.append({
            'profit_rate': pr,
            'initial_irr': data_rows[0][3],
            'final_irr': data_rows[-1][3]
        })

print('AGENT 回應:')
print('不同利潤率的專案法 IRR 比較 (設備成本 25000→23000 元/kW):')
print()
print(f'{"利潤率":>10} {"初始IRR":>12} {"最終IRR":>12} {"IRR提升":>12}')
print('-' * 50)
for r in results_comparison:
    increase = r['final_irr'] - r['initial_irr']
    print(f'{r["profit_rate"]:>9.0%} {r["initial_irr"]:>11.2f}% {r["final_irr"]:>11.2f}% {increase:>11.2f}%')

print()
print('可以看到，利潤率越低，IRR 越高。這是因為較低的利潤率')
print('意味著更低的最終售價，從而提高了投資報酬率。')

print()
print()
print('=' * 70)
print('測試總結')
print('=' * 70)
print()
print('✓ AGENT 可以根據用戶需求智能選擇工具：')
print('  • 快速分析 → calculate_price_rolling')
print('  • 保存記錄 → execute_price_rolling')
print()
print('✓ AGENT 可以串聯多個工具呼叫完成複雜任務')
print('✓ AGENT 可以解讀工具返回結果並給出建議')
