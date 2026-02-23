"""
測試網頁輸入值讀取
模擬前端發送的請求數據
"""
import json

# 模擬不同模式的請求數據
test_cases = {
    "CashMode": {
        "mode": "CashMode",
        "equipment_cost": 32000,
        "profit_rate": 0.22,
        "development_fee": 1800,
        "boundary": 18000,
        "step": 2500,
        "case_name": "測試案場",
        "original_filename": "公版.xlsx",
        "sheet_name": "輸入公版"
    },
    "RatioMode": {
        "mode": "RatioMode",
        "equipment_cost": 35000,
        "profit_rate": 0.25,
        "development_fee": 2200,
        "boundary": 20000,
        "step": 0.06,
        "case_name": "測試案場",
        "original_filename": "公版.xlsx"
    },
    "ConditionalMode": {
        "mode": "ConditionalMode",
        "equipment_cost": 40000,
        "profit_rate": 0.20,
        "development_fee": 1500,
        "boundary": 22000,
        "max_value": 45000,
        "min_value": 28000,
        "step1": 1800,
        "step2": 900,
        "step3": 400,
        "case_name": "測試案場",
        "original_filename": "公版.xlsx"
    },
    "CustomizeMode": {
        "mode": "CustomizeMode",
        "equipment_cost": 38000,
        "profit_rate": 0.23,
        "development_fee": 1700,
        "boundary": 19000,
        "adjust_times": 6,
        "steps": [3000, 2500, 2000, 1500, 1000, 500],
        "case_name": "測試案場",
        "original_filename": "公版.xlsx"
    },
    "開發費為空": {
        "mode": "CashMode",
        "equipment_cost": 30000,
        "profit_rate": 0.20,
        # development_fee 故意不提供
        "boundary": 20000,
        "step": 2000,
        "case_name": "測試案場",
        "original_filename": "公版.xlsx"
    }
}

def test_parameter_reading(mode_name, request_data):
    """
    測試參數讀取邏輯
    """
    print(f"\n{'='*60}")
    print(f"測試模式: {mode_name}")
    print(f"{'='*60}")
    
    # 基本參數
    mode = request_data.get('mode')
    equipment_cost = request_data.get('equipment_cost')
    profit_rate = request_data.get('profit_rate')
    development_fee = request_data.get('development_fee')
    boundary = request_data.get('boundary')
    
    print(f"📥 接收到的基本參數:")
    print(f"   - mode: {mode}")
    print(f"   - equipment_cost: {equipment_cost} (類型: {type(equipment_cost).__name__})")
    print(f"   - profit_rate: {profit_rate} (類型: {type(profit_rate).__name__})")
    print(f"   - development_fee: {development_fee} (類型: {type(development_fee).__name__ if development_fee is not None else 'None'})")
    print(f"   - boundary: {boundary} (類型: {type(boundary).__name__})")
    
    # 開發費處理邏輯
    if development_fee is None:
        print(f"   ⚠️  開發費為 None，應從 Excel 讀取預設值")
    else:
        print(f"   ✅ 使用網頁輸入的開發費: {development_fee}")
    
    # 案場資訊
    case_name = request_data.get('case_name', '')
    original_filename = request_data.get('original_filename', '')
    sheet_name = request_data.get('sheet_name', None)
    
    print(f"\n📁 案場資訊:")
    print(f"   - case_name: {case_name}")
    print(f"   - original_filename: {original_filename}")
    print(f"   - sheet_name: {sheet_name}")
    
    # 模式特定參數
    print(f"\n🔧 模式特定參數:")
    
    if mode == "CashMode":
        step = request_data.get('step')
        print(f"   - step: {step} (類型: {type(step).__name__ if step is not None else 'None'})")
        
    elif mode == "RatioMode":
        step = request_data.get('step')
        print(f"   - step: {step} (類型: {type(step).__name__ if step is not None else 'None'})")
        
    elif mode == "ConditionalMode":
        max_value = request_data.get('max_value')
        min_value = request_data.get('min_value')
        step1 = request_data.get('step1')
        step2 = request_data.get('step2')
        step3 = request_data.get('step3')
        
        print(f"   - max_value: {max_value} (類型: {type(max_value).__name__ if max_value is not None else 'None'})")
        print(f"   - min_value: {min_value} (類型: {type(min_value).__name__ if min_value is not None else 'None'})")
        print(f"   - step1: {step1} (類型: {type(step1).__name__ if step1 is not None else 'None'})")
        print(f"   - step2: {step2} (類型: {type(step2).__name__ if step2 is not None else 'None'})")
        print(f"   - step3: {step3} (類型: {type(step3).__name__ if step3 is not None else 'None'})")
        
    elif mode == "CustomizeMode":
        adjust_times = request_data.get('adjust_times')
        steps = request_data.get('steps')
        
        print(f"   - adjust_times: {adjust_times} (類型: {type(adjust_times).__name__ if adjust_times is not None else 'None'})")
        print(f"   - steps: {steps} (類型: {type(steps).__name__ if steps is not None else 'None'})")
        if steps:
            print(f"     步伐列表: {steps}")
    
    # 構建工具參數
    print(f"\n🛠️  構建的工具參數:")
    tool_args = {
        "mode": mode,
        "boundary": boundary
    }
    
    if equipment_cost is not None:
        tool_args["equipment_cost"] = equipment_cost
    if profit_rate is not None:
        tool_args["profit_rate"] = profit_rate
    if development_fee is not None:
        tool_args["development_fee"] = development_fee
    if sheet_name is not None:
        tool_args["sheet_name"] = sheet_name
    
    # 添加模式特定參數
    if mode == "CashMode":
        step = request_data.get('step')
        if step is not None:
            tool_args["step"] = step
            
    elif mode == "RatioMode":
        step = request_data.get('step')
        if step is not None:
            tool_args["step"] = step
            
    elif mode == "ConditionalMode":
        for key in ['max_value', 'min_value', 'step1', 'step2', 'step3']:
            val = request_data.get(key)
            if val is not None:
                tool_args[key] = val
                
    elif mode == "CustomizeMode":
        adjust_times = request_data.get('adjust_times')
        steps = request_data.get('steps')
        if adjust_times is not None:
            tool_args["adjust_times"] = adjust_times
        if steps is not None:
            tool_args["steps"] = steps
    
    print(json.dumps(tool_args, indent=2, ensure_ascii=False))
    
    # 檢查是否有遺漏的參數
    print(f"\n✅ 參數檢查:")
    missing_params = []
    
    # 必要參數
    if mode is None:
        missing_params.append("mode")
    if boundary is None:
        missing_params.append("boundary")
    
    # 模式特定必要參數
    if mode == "CashMode" and request_data.get('step') is None:
        missing_params.append("step (CashMode)")
    elif mode == "RatioMode" and request_data.get('step') is None:
        missing_params.append("step (RatioMode)")
    elif mode == "ConditionalMode":
        for key in ['max_value', 'min_value', 'step1', 'step2', 'step3']:
            if request_data.get(key) is None:
                missing_params.append(key)
    elif mode == "CustomizeMode" and request_data.get('adjust_times') is None:
        missing_params.append("adjust_times")
    
    if missing_params:
        print(f"   ❌ 缺少參數: {', '.join(missing_params)}")
    else:
        print(f"   ✅ 所有必要參數都已提供")
    
    # 檢查可選參數
    optional_params = ['equipment_cost', 'profit_rate', 'development_fee', 'sheet_name']
    provided_optional = [p for p in optional_params if request_data.get(p) is not None]
    missing_optional = [p for p in optional_params if request_data.get(p) is None]
    
    if provided_optional:
        print(f"   ✅ 提供的可選參數: {', '.join(provided_optional)}")
    if missing_optional:
        print(f"   ℹ️  未提供的可選參數: {', '.join(missing_optional)}")

if __name__ == "__main__":
    print("\n" + "="*60)
    print("網頁輸入值讀取測試")
    print("="*60)
    
    for mode_name, request_data in test_cases.items():
        test_parameter_reading(mode_name, request_data)
    
    print(f"\n{'='*60}")
    print("測試完成")
    print(f"{'='*60}\n")
