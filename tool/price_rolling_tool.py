# tool/price_rolling_tool.py
from typing import List, Dict, Union, Optional, Any
from tool.equipment_cost_services import (
    CashMode,
    RatioMode,
    CustomizeMode,
    CostStructureService
)
from tool.finance_tool import FinanceTool

class PriceRollingTool:
    """價金滾算工具類"""

    def __init__(self):
        self.finance_tool = FinanceTool()

    def set_excel_file(self, excel_file: str, sheet_name: str = None):
        """設定財務工具使用的 Excel 檔案"""
        self.finance_tool.set_excel_file(excel_file, sheet_name)

    def calculate_price_rolling(
        self,
        mode: str,
        boundary: int,
        equipment_cost: Optional[int] = None,
        profit_rate: Optional[float] = None,
        development_fee: Optional[int] = None,
        sheet_name: Optional[str] = None,
        step: Optional[Union[int, float]] = None,
        max_value: Optional[int] = None,
        min_value: Optional[int] = None,
        step1: Optional[int] = None,
        step2: Optional[int] = None,
        step3: Optional[int] = None,
        adjust_times: Optional[int] = None,
        steps: Optional[List[int]] = None
    ) -> Dict:
        """
        執行價金滾算

        Args:
            mode: 模式 ("CashMode", "RatioMode", "CustomizeMode")
            boundary: 邊界價格
            equipment_cost: 初始價金 / kW (若未提供則嘗試從 Excel 讀取)
            profit_rate: 信邦利潤率 (例如 0.05) (若未提供則嘗試從 Excel 讀取)
            development_fee: 開發費用 (預設為 None，若未提供則嘗試從 Excel 讀取)
            sheet_name: Excel 工作表名稱 (用於計算 IRR)
            step: 調整 Step (CashMode) 或 調整比例 (RatioMode)
            adjust_times: 調整次數 (CustomizeMode)
            steps: 自訂 Step 列表 (CustomizeMode)
        """
        try:
            # 嘗試讀取 Excel 數據以獲取預設值
            excel_data = {}
            try:
                # 確保有設定 Excel 檔案
                if not self.finance_tool.excel_file:
                    self.finance_tool._load_excel_data() # 這會觸發自動搜尋
                
                # 讀取數據
                excel_data = self.finance_tool._load_excel_data()
            except Exception as e:
                print(f"讀取 Excel 數據時發生錯誤: {e}")

            # 如果沒有提供開發費用，嘗試從 Excel 讀取
            if development_fee is None:
                dev_cost = excel_data.get('development_cost')
                if dev_cost is not None:
                    development_fee = int(dev_cost)
                    print(f"自動讀取開發費用: {development_fee}")
                else:
                    development_fee = 0
                    print("Excel 中未設定開發費用 (development_cost 為 None)，使用預設值 0")

            # 如果沒有提供利潤率，嘗試從 Excel 讀取
            if profit_rate is None:
                p_rate = excel_data.get('profit_rate')
                if p_rate is not None:
                    profit_rate = float(p_rate)
                    print(f"自動讀取利潤率: {profit_rate}")
                else:
                    return {"success": False, "message": "未提供利潤率 (profit_rate)，且無法從 Excel 讀取"}

            # 如果沒有提供初始價金，嘗試從 Excel 讀取
            if equipment_cost is None:
                eq_cost = excel_data.get('equipment_cost')
                if eq_cost is not None:
                    equipment_cost = int(eq_cost)
                    print(f"自動讀取初始價金: {equipment_cost}")
                else:
                    return {"success": False, "message": "未提供初始價金 (equipment_cost)，且無法從 Excel 讀取"}

            # 初始化服務
            cost_service = CostStructureService(
                profit_rate=profit_rate,
                development_fee=development_fee
            )
            
            # 1. 先計算原本 Excel 的 IRR (Base Case)
            base_irr_result = {}
            try:
                base_res = self.finance_tool.calculate_irr(sheet_name)
                if base_res.get("success"):
                    base_data = base_res.get("data", {})
                    base_irr_result = {
                        "equipment_cost": base_data.get("raw_equipment_cost"),
                        "profit_rate": base_data.get("profit_rate"),
                        "project_irr": base_data.get("project_irr"),
                        "cost_method_irr": base_data.get("cost_method_irr"),
                        "equity_method_irr": base_data.get("equity_method_irr")
                    }
            except Exception as e:
                print(f"計算基礎 IRR 失敗: {e}")

            adjustment_record = []
            
            if mode == "CashMode":
                if step is None:
                    return {"success": False, "message": "CashMode 需要 step 參數"}
                mode_instance = CashMode(boundary, int(step))
                adjustment_record = mode_instance.calculation(equipment_cost)
                
            elif mode == "RatioMode":
                if step is None:
                    return {"success": False, "message": "RatioMode 需要 step 參數"}
                mode_instance = RatioMode(boundary, float(step))
                adjustment_record = mode_instance.calculation(equipment_cost)
                
            elif mode == "CustomizeMode":
                if not adjust_times:
                     return {"success": False, "message": "CustomizeMode 需要 adjust_times 參數"}
                
                mode_instance = CustomizeMode(boundary, adjust_times)
                
                if not steps:
                    # 自動生成
                    steps = mode_instance.automatic_configuration(equipment_cost)
                
                adjustment_record = mode_instance.calculation(steps, equipment_cost)
                
            else:
                return {"success": False, "message": f"未知的模式: {mode}"}

            profit_list = cost_service.get_profit(adjustment_record)
            final_cost_list = cost_service.equipment_cost_calculation(adjustment_record)
            
            # 格式化結果 (使用更緊湊的格式以節省 Token)
            columns = ["price_per_kw", "profit_per_kw", "final_price_per_kw", "project_irr", "cost_method_irr", "equity_method_irr"]
            data_rows = []
            
            print(f"開始計算 {len(adjustment_record)} 個價金的 IRR...")
            
            for i, (price, profit, final_cost) in enumerate(zip(adjustment_record, profit_list, final_cost_list)):
                # 計算該價金下的 IRR
                irr_result = self.finance_tool.calculate_scenario_irr(
                    equipment_cost=price, 
                    profit_rate=profit_rate,
                    development_cost=development_fee,
                    sheet_name=sheet_name
                )
                
                # 檢查是否有 IRR 計算失敗
                if (irr_result.get('project_irr') is None or
                        irr_result.get('cost_method_irr') is None or
                        irr_result.get('equity_method_irr') is None):
                    return {"success": False, "message": "公版數值異常導致IRR計算錯誤"}

                row = [
                    price,
                    profit,
                    final_cost,
                    irr_result.get('project_irr'),
                    irr_result.get('cost_method_irr'),
                    irr_result.get('equity_method_irr')
                ]
                data_rows.append(row)
                
            return {
                "success": True,
                "mode": mode,
                "base_irr": base_irr_result,
                "used_parameters": {
                    "equipment_cost": equipment_cost,
                    "profit_rate": profit_rate,
                    "development_fee": development_fee,
                    "boundary": boundary,
                    "step": step
                },
                "results_summary": {
                    "columns": columns,
                    "data": data_rows,
                    "count": len(data_rows)
                },
                "note": "Results are in 'results_summary'. 'data' contains rows corresponding to 'columns'."
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"計算失敗: {str(e)}"
            }

# 定義工具的 schema
PRICE_ROLLING_TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "calculate_price_rolling",
            "description": "【純計算分析】執行價金滾算分析，計算不同價金下的 IRR 變化。此工具僅進行計算和分析，不會寫入任何檔案。支援多種模式：CashMode(固定金額調整), RatioMode(比例調整), CustomizeMode(自訂調整)。回傳結果包含原始 Excel IRR (base_irr) 以及滾算後的專案法 IRR、成本法 IRR 和權益法 IRR。如需將結果寫入 Excel 檔案，請使用 execute_price_rolling 工具。",
            "parameters": {
                "type": "object",
                "properties": {
                    "mode": {
                        "type": "string",
                        "enum": ["CashMode", "RatioMode", "CustomizeMode"],
                        "description": "滾算模式"
                    },
                    "equipment_cost": {
                        "type": "integer",
                        "description": "初始價金 / kW (若未提供則嘗試從 Excel 讀取)"
                    },
                    "profit_rate": {
                        "type": "number",
                        "description": "信邦利潤率 (例如 0.05 表示 5%) (若未提供則嘗試從 Excel 讀取)"
                    },
                    "boundary": {
                        "type": "integer",
                        "description": "邊界價格"
                    },
                    "development_fee": {
                        "type": "integer",
                        "description": "開發費用 (預設為 None，自動從 Excel 讀取)"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Excel 工作表名稱 (用於計算 IRR)"
                    },
                    "step": {
                        "type": "number",
                        "description": "調整 Step (CashMode 用整數) 或 調整比例 (RatioMode 用小數 0~1)"
                    },
                    "adjust_times": {
                        "type": "integer",
                        "description": "調整次數 (CustomizeMode)"
                    },
                    "steps": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "自訂 Step 列表 (CustomizeMode)，若不提供則自動生成"
                    }
                },
                "required": ["mode", "boundary"]
            }
        }
    }
]
