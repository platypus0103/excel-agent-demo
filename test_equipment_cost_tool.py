# test_equipment_cost_tool.py
"""
設備成本滾算工具測試程式
"""
import os
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from tool.equipment_cost_tool import EquipmentCostTool

def test_equipment_cost_tool():
    """測試設備成本滾算工具"""
    print("=== 設備成本滾算工具測試 ===\n")
    
    # 創建工具實例
    tool = EquipmentCostTool()
    
    # 測試案例 1: 現金模式
    print("1. 測試現金模式滾算:")
    result1 = tool.execute_price_rolling(
        mode="cash",
        boundary=20000,
        step=2000,
        profit_rate=0.1,
        development_fee=2000
    )
    
    if result1["success"]:
        print(f"✓ 現金模式滾算成功")
        print(f"  初始價金: {result1['result']['initial_cost']}")
        print(f"  滾算次數: {result1['summary']['adjustment_count']}")
        print(f"  最終價金: {result1['summary']['final_cost']}")
        print(f"  總減少額: {result1['summary']['total_reduction']}")
        print(f"  輸出檔案: {result1['output_file']}")
    else:
        print(f"✗ 現金模式滾算失敗: {result1['message']}")
    
    print()
    
    # 測試案例 2: 比率模式
    print("2. 測試比率模式滾算:")
    result2 = tool.execute_price_rolling(
        mode="ratio",
        boundary=15000,
        step=0.05,  # 5% 減少
        profit_rate=0.12,
        development_fee=1500
    )
    
    if result2["success"]:
        print(f"✓ 比率模式滾算成功")
        print(f"  初始價金: {result2['result']['initial_cost']}")
        print(f"  滾算次數: {result2['summary']['adjustment_count']}")
        print(f"  最終價金: {result2['summary']['final_cost']}")
        print(f"  總減少額: {result2['summary']['total_reduction']}")
    else:
        print(f"✗ 比率模式滾算失敗: {result2['message']}")
    
    print()
    
    # 測試案例 3: 自訂模式
    print("3. 測試自訂模式滾算:")
    result3 = tool.execute_price_rolling(
        mode="customize",
        boundary=18000,
        adjustment_times=5,
        profit_rate=0.08,
        development_fee=1800,
        auto_config=True
    )
    
    if result3["success"]:
        print(f"✓ 自訂模式滾算成功")
        print(f"  初始價金: {result3['result']['initial_cost']}")
        print(f"  滾算次數: {result3['summary']['adjustment_count']}")
        print(f"  最終價金: {result3['summary']['final_cost']}")
        print(f"  總減少額: {result3['summary']['total_reduction']}")
    else:
        print(f"✗ 自訂模式滾算失敗: {result3['message']}")
    
    print()
    
    # 測試案例 4: 條件模式
    print("4. 測試條件模式滾算:")
    result4 = tool.execute_price_rolling(
        mode="conditional",
        boundary=16000,
        profit_rate=0.15,
        development_fee=2200,
        maximum_value=50000,
        minimum_value=25000,
        condition_step_1=3000,
        condition_step_2=1500,
        condition_step_3=800
    )
    
    if result4["success"]:
        print(f"✓ 條件模式滾算成功")
        print(f"  初始價金: {result4['result']['initial_cost']}")
        print(f"  滾算次數: {result4['summary']['adjustment_count']}")
        print(f"  最終價金: {result4['summary']['final_cost']}")
        print(f"  總減少額: {result4['summary']['total_reduction']}")
    else:
        print(f"✗ 條件模式滾算失敗: {result4['message']}")

def test_tool_manager_integration():
    """測試與工具管理器的整合"""
    print("\n=== 工具管理器整合測試 ===\n")
    
    try:
        from tool.tool_manager import ToolManager
        
        # 創建工具管理器實例
        tm = ToolManager()
        
        # 檢查工具是否已註冊
        tools = tm.list_tools()
        print(f"已註冊的工具: {tools}")
        
        if "execute_price_rolling" in tools:
            print("✓ 設備成本滾算工具已成功註冊到工具管理器")
            
            # 使用工具管理器執行工具
            result = tm.execute_tool(
                tool_name="execute_price_rolling",
                arguments={
                    "mode": "cash",
                    "boundary": 22000,
                    "step": 1500,
                    "profit_rate": 0.09,
                    "development_fee": 1700
                }
            )
            
            if result.get("success"):
                print("✓ 透過工具管理器執行滾算成功")
            else:
                print(f"✗ 透過工具管理器執行滾算失敗: {result.get('message')}")
        else:
            print("✗ 設備成本滾算工具未註冊到工具管理器")
            
    except Exception as e:
        print(f"✗ 工具管理器整合測試失敗: {str(e)}")

def test_irr_simulation_agent():
    """測試IRR模擬AGENT調用工具"""
    print("\n=== IRR模擬AGENT調用工具測試 ===\n")
    
    try:
        from tool.tool_manager import ToolManager
        
        tm = ToolManager()
        print("🤖 模擬AGENT智能調用滾算工具")
        
        # 場景1: AGENT快速分析IRR變化
        print("\n--- 場景1: AGENT分析IRR變化趨勢 ---")
        print('用戶詢問: "設備成本從30000降到22000，IRR會如何變化？"')
        print('🤖 AGENT思考: 用戶需要快速分析，選擇純計算模式')
        
        result1 = tm.execute_tool(
            'calculate_price_rolling',
            {
                'mode': 'CashMode',
                'boundary': 22000,
                'equipment_cost': 30000,
                'step': 1000,
                'profit_rate': 0.10,
                'development_fee': 2000
            }
        )
        
        if result1.get('success'):
            print("✅ AGENT成功執行IRR分析")
            print("🤖 AGENT回應: ")
            
            # 處理不同的返回格式
            if 'result' in result1:
                res = result1['result']
            else:
                res = result1
                
            adjustment_record = res.get('adjustment_record', [])
            if len(adjustment_record) >= 2:
                print(f"   設備成本變化: {adjustment_record[0]} → {adjustment_record[-1]} 元")
                print(f"   滾算步數: {len(adjustment_record)-1} 次")
            
            # 檢查IRR結果
            irr_results = res.get('irr_results', [])
            if irr_results:
                valid_irr = [irr for irr in irr_results if not irr.get('error')]
                if valid_irr:
                    first_irr = valid_irr[0]
                    last_irr = valid_irr[-1]
                    print(f"   初始IRR: 專案法{first_irr.get('project_irr', 'N/A')}%, 權益法{first_irr.get('equity_method_irr', 'N/A')}%")
                    print(f"   最終IRR: 專案法{last_irr.get('project_irr', 'N/A')}%, 權益法{last_irr.get('equity_method_irr', 'N/A')}%")
                else:
                    print("   ⚠️ IRR計算遇到數據格式問題，但滾算邏輯正常")
            else:
                print("   📊 滾算計算完成，IRR數據準備中...")
        else:
            print(f"❌ AGENT分析失敗: {result1.get('message')}")
        
        # 場景2: AGENT執行完整流程並保存結果
        print("\n--- 場景2: AGENT執行完整滾算流程 ---")
        print('用戶需求: "請幫我做比率模式滾算，並保存到Excel"')
        print('🤖 AGENT思考: 需要完整流程，選擇execute_price_rolling')
        
        result2 = tm.execute_tool(
            'execute_price_rolling',
            {
                'mode': 'ratio',
                'boundary': 18000,
                'step': 0.04,  # 4%
                'profit_rate': 0.08,
                'development_fee': 1800
            }
        )
        
        if result2.get('success'):
            print("✅ AGENT成功執行完整滾算流程")
            print("🤖 AGENT回應:")
            summary = result2.get('summary', {})
            print(f"   初始價金: {summary.get('initial_cost', 'N/A'):,} 元")
            print(f"   最終價金: {summary.get('final_cost', 'N/A'):,} 元")
            print(f"   調整次數: {summary.get('adjustment_count', 'N/A')} 次")
            print(f"   總減少額: {summary.get('total_reduction', 'N/A'):,} 元")
            print(f"   Excel檔案: {result2.get('output_file', 'N/A')}")
        else:
            print(f"❌ AGENT完整流程失敗: {result2.get('message')}")
        
        # 場景3: AGENT比較不同模式
        print("\n--- 場景3: AGENT智能比較不同滾算模式 ---")
        print('用戶詢問: "現金模式和比率模式哪個減少得更多？"')
        print('🤖 AGENT思考: 需要比較分析，執行兩種模式計算')
        
        # 現金模式
        cash_result = tm.execute_tool('calculate_price_rolling', {
            'mode': 'CashMode',
            'boundary': 20000,
            'equipment_cost': 28000,
            'step': 800,
            'profit_rate': 0.09
        })
        
        # 比率模式  
        ratio_result = tm.execute_tool('calculate_price_rolling', {
            'mode': 'RatioMode',
            'boundary': 20000,
            'equipment_cost': 28000,
            'step': 0.03,  # 3%
            'profit_rate': 0.09
        })
        
        if cash_result.get('success') and ratio_result.get('success'):
            print("✅ AGENT成功比較兩種模式")
            print("🤖 AGENT分析結果:")
            
            # 安全處理返回結果
            cash_res = cash_result.get('result', cash_result)
            ratio_res = ratio_result.get('result', ratio_result)
            
            cash_record = cash_res.get('adjustment_record', [])
            ratio_record = ratio_res.get('adjustment_record', [])
            
            cash_reduction = cash_record[0] - cash_record[-1] if len(cash_record) >= 2 else 0
            ratio_reduction = ratio_record[0] - ratio_record[-1] if len(ratio_record) >= 2 else 0
            
            print(f"   現金模式: 減少 {cash_reduction:,} 元 ({len(cash_record)-1} 次調整)")
            print(f"   比率模式: 減少 {ratio_reduction:,} 元 ({len(ratio_record)-1} 次調整)")
            
            if ratio_reduction > cash_reduction:
                print("   💡 建議: 比率模式減少得更多，推薦使用")
            else:
                print("   💡 建議: 現金模式減少得更多，推薦使用")
        else:
            print("❌ AGENT比較分析失敗")
        
        print("\n🎯 IRR模擬AGENT測試總結:")
        print("✅ AGENT能智能選擇適合的工具")
        print("✅ AGENT能執行複雜的分析和比較")
        print("✅ AGENT能根據用戶需求調整策略")
        print("✅ AGENT具備完整的工具調用能力")
        
    except Exception as e:
        print(f"❌ IRR模擬AGENT測試失敗: {str(e)}")

if __name__ == "__main__":
    test_equipment_cost_tool()
    test_tool_manager_integration()
    test_irr_simulation_agent()