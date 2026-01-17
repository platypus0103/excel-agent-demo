# tool/equipment_cost_tool.py
import os
import shutil
import time
import datetime
from openpyxl import load_workbook, Workbook
from typing import Dict, List, Optional, Any, Union
from tool.equipment_cost_services import (
    CostStructureService, 
    CashMode, 
    RatioMode, 
    ConditionalMode, 
    CustomizeMode
)
from tool.finance_tool import FinanceTool

class EquipmentCostTool:
    """設備成本滾算工具類"""

    def __init__(self):
        """初始化設備成本工具"""
        self.excel_folder = "Excel"
        self.excel_final_folder = "Excel final"
        self.output_file_name = "滾算後記錄.xlsx"
        self.finance_tool = None  # 財務工具實例

    def _find_excel_file(self, directory: str) -> Optional[str]:
        """在指定目錄中找到第一個 Excel 檔案"""
        if not os.path.exists(directory):
            return None
        
        for file in os.listdir(directory):
            if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                return os.path.join(directory, file)
        return None

    def _safe_float_convert(self, value, default: float) -> float:
        """安全地將值轉換為浮點數"""
        try:
            if value is None:
                return default
            if isinstance(value, str) and (value.strip() == '' or '參數數值' in value):
                return default
            return float(value)
        except (ValueError, TypeError):
            return default
    
    def _safe_int_convert(self, value, default: int) -> int:
        """安全地將值轉換為整數"""
        try:
            if value is None:
                return default
            if isinstance(value, str) and (value.strip() == '' or '參數數值' in value):
                return default
            return int(float(value))
        except (ValueError, TypeError):
            return default

    def _read_excel_data(self, file_path: str) -> Dict[str, Any]:
        """讀取 Excel 檔案數據"""
        try:
            wb = load_workbook(file_path, data_only=True)
            sheet = wb.active
            
            # 根據提供的 Excel 結構讀取數據
            data = {
                "project_name": sheet["B2"].value or "淡江大學-財務模擬案場",                "plant_lifetime": self._safe_int_convert(sheet["B3"].value, 20),  # 電站壽命                "capacity": self._safe_float_convert(sheet["B4"].value, 436.1),#建置量(kWp)
                "start_year": self._safe_int_convert(sheet["B5"].value, 2020),#起始年度
                "end_year": self._safe_int_convert(sheet["D5"].value, 2039),#結束年度
                "annual_generation": self._safe_float_convert(sheet["B10"].value, 0),#年度AC併網總發電量
                "annual_ac_generation": self._safe_float_convert(sheet["C10"].value, 0),  # 年度AC併網總發電量
                "electricity_revenue_years": self._safe_int_convert(sheet["D10"].value, 0),  # 電費收入年限
                "first_year_decline_rate": self._safe_float_convert(sheet["C11"].value, 0),  # 首年衰退率
                "fit_rate": self._safe_float_convert(sheet["C12"].value, 0),#次年衰退率               
                "fit_price_c13": self._safe_float_convert(sheet["C13"].value, 0),  # 躉售費率(FIT)從C13讀取
                "equipment_cost": self._safe_int_convert(sheet["C16"].value, 0),  # 每kWp設備成本
                "equipment_amortization_years": self._safe_int_convert(sheet["D16"].value, 0),  # 設備費用(價金)支提年限
                "loan_interest": self._safe_float_convert(sheet["C26"].value, 0),#所得稅率
                "development_fee": self._safe_int_convert(sheet["C18"].value, 0),  # 開發費
                "rent_calculation_method": self._safe_float_convert(sheet["C19"].value, 0),  # 租金費用租金計算方式
                "rent_calculation_amortization_years": self._safe_int_convert(sheet["D19"].value, 0),  # 租金費用租金計算方式支提年限
                "rent_method1_total": self._safe_float_convert(sheet["C20"].value, 0),  # 租金費用【方式1參數】租金總額
                "rent_method1_amortization_years": self._safe_int_convert(sheet["D20"].value, 0),  # 租金費用【方式1參數】租金總額支提年限
                "rent_method2_ratio": self._safe_float_convert(sheet["C21"].value, 0),  # 租金費用【方式2參數】抽成比率
                "rent_method2_amortization_years": self._safe_int_convert(sheet["D21"].value, 0),  # 租金費用【方式2參數】抽成比率支提年限
                "loan_amount": self._safe_float_convert(sheet["C24"].value, 0),#【方式1參數】租金總額
                "maintenance_cost": self._safe_float_convert(sheet["C22"].value, 0),#運維費用每kw運維成本
                "maintenance_amortization_years": self._safe_int_convert(sheet["D22"].value, 0),  # 運維費用支提年限
                "insurance_cost": self._safe_float_convert(sheet["C23"].value, 0),#保險費用占整體設備費用比例
                "insurance_amortization_years": self._safe_int_convert(sheet["D23"].value, 0),  # 保險費用支提年限
                "recycle_cost": self._safe_float_convert(sheet["C24"].value, 0),  # 從C24讀取回收費用
                "recycle_amortization_years": self._safe_int_convert(sheet["D24"].value, 0),  # 模組回收費用支提年限
                "profit_rate": self._safe_float_convert(sheet["C17"].value, 0),  # 信邦利潤率
                "bank_interest_rate": self._safe_float_convert(sheet["C25"].value, 0),  # 利息費用銀行利率
                "interest_amortization_years": self._safe_int_convert(sheet["D25"].value, 0),  # 利息費用支提年限
                "tax_rate": self._safe_float_convert(sheet["C26"].value, 0),  # 所得稅率
                "tax_amortization_years": self._safe_int_convert(sheet["D26"].value, 0),  # 所得稅支提年限
                # 新增 C31~C34 數據讀取
                "c31_value": self._safe_float_convert(sheet["C31"].value, 0),  # 貸款成數占總設備費用比例
                "c32_value": self._safe_float_convert(sheet["C32"].value, 0),  # 貸款還款攤還期數
                "c33_value": self._safe_float_convert(sheet["C33"].value, 0),  # 現金股利股利比率
                "c34_value": self._safe_float_convert(sheet["C34"].value, 0)   # 年底減資攤還期數
            }
            
            wb.close()
            return data
            
        except Exception as e:
            raise Exception(f"讀取 Excel 檔案失敗: {str(e)}")

    def _calculate_irr_for_prices(self, excel_file: str, adjustment_record: List[int], profit_rate: float, development_fee: int, sheet_name: str = None) -> List[Dict[str, float]]:
        """為每個價金計算IRR（參考price_rolling_tool實作）"""
        try:
            # 初始化財務工具
            if self.finance_tool is None:
                self.finance_tool = FinanceTool()
            
            # 設定Excel檔案
            self.finance_tool.set_excel_file(excel_file, sheet_name)
            
            irr_results = []
            
            for price in adjustment_record:
                try:
                    # 使用財務工具計算IRR（參考price_rolling_tool的方式）
                    irr_result = self.finance_tool.calculate_scenario_irr(
                        equipment_cost=price,
                        profit_rate=profit_rate,
                        development_cost=development_fee,
                        sheet_name=sheet_name
                    )
                    irr_results.append(irr_result)
                except Exception as e:
                    print(f"計算 IRR 失敗 (價金 {price}): {str(e)}")
                    # 如果計算失敗，填入空值
                    irr_results.append({
                        "project_irr": None,
                        "cost_method_irr": None,
                        "equity_method_irr": None
                    })
                    
            return irr_results
            
        except Exception as e:
            print(f"初始化財務工具失敗: {str(e)}")
            # 如果初始化失敗，為所有價金返回空值
            return [{
                "project_irr": None,
                "cost_method_irr": None, 
                "equity_method_irr": None
            } for _ in adjustment_record]

    def _prepare_output_file(self, source_file_path: str) -> str:
        """準備輸出檔案，先複製滾算後記錄.xlsx檔案"""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = "滾算後記錄"
        output_path = os.path.join(self.excel_final_folder, f"{base_name}_{timestamp}.xlsx")
        
        # 確保目錄存在
        if not os.path.exists(self.excel_final_folder):
            os.makedirs(self.excel_final_folder)
        
        # 要複製的目標檔案路徑
        target_file = os.path.join(self.excel_final_folder, "滾算後記錄.xlsx")
        
        # 複製滾算後記錄.xlsx檔案到輸出位置
        if os.path.exists(target_file):
            try:
                shutil.copy2(target_file, output_path)
                print(f"已複製滾算後記錄.xlsx 到 {output_path}")
            except PermissionError:
                # 如果複製失敗，從原始檔案複製
                print(f"無法複製滾算後記錄.xlsx(可能被鎖定)，將從原始檔案複製: {source_file_path}")
                if os.path.exists(source_file_path):
                    try:
                        shutil.copy2(source_file_path, output_path)
                        print(f"已從原始檔案複製到 {output_path}")
                    except PermissionError:
                        print(f"原始檔案也被鎖定，將創建新檔案: {output_path}")
                        wb = Workbook()
                        sheet = wb.active
                        sheet.title = "滾算結果"
                        wb.save(output_path)
                        wb.close()
                else:
                    wb = Workbook()
                    sheet = wb.active
                    sheet.title = "滾算結果"
                    wb.save(output_path)
                    wb.close()
        else:
            # 如果滾算後記錄.xlsx不存在，從原始檔案複製
            print(f"滾算後記錄.xlsx 不存在，將從原始檔案複製: {source_file_path}")
            if os.path.exists(source_file_path):
                try:
                    shutil.copy2(source_file_path, output_path)
                    print(f"已從原始檔案複製到 {output_path}")
                except PermissionError:
                    print(f"原始檔案被鎖定，將創建新檔案: {output_path}")
                    wb = Workbook()
                    sheet = wb.active
                    sheet.title = "滾算結果"
                    wb.save(output_path)
                    wb.close()
            else:
                # 創建基本模板
                wb = Workbook()
                sheet = wb.active
                sheet.title = "滾算結果"
                wb.save(output_path)
                wb.close()
        
        return output_path

    def _get_next_sheet_number(self, wb) -> int:
        """取得下一個可用的工作表編號"""
        existing_numbers = []
        for sheet_name in wb.sheetnames:
            if sheet_name.startswith("開始滾算紀錄"):
                # 提取編號，支援"開始滾算紀錄1", "開始滾算紀錄2"等格式
                number_part = sheet_name.replace("開始滾算紀錄", "")
                try:
                    existing_numbers.append(int(number_part))
                except ValueError:
                    pass
        
        # 返回下一個可用編號
        if not existing_numbers:
            return 1  # 如果沒有任何開始滾算紀錄工作表，從1開始
        else:
            return max(existing_numbers) + 1

    def _fill_copied_sheet_with_data(self, new_sheet, result: Dict[str, Any], profit_rate: float, capacity: float):
        """在複製的工作表中填入滾算結果數據，為後續從輸入公版取值做準備"""
        try:
            # 基本滾算資訊
            adjustment_record = result.get("adjustment_record", [])
            irr_results = result.get("irr_results", [])
            mode = result.get("mode", "")
            boundary = result.get("boundary", "")
            
            # 找一個不會影響原有格式的位置來放滾算結果數據
            # 在工作表的右側或下方空白區域填入數據，避免覆蓋原有內容
            
            # 先找到工作表的最大使用範圍
            max_row = new_sheet.max_row
            max_col = new_sheet.max_column
            
            # 在右側空白區域(例如從列Z開始)填入滾算摘要資訊
            summary_col = max_col + 2  # 留兩列間距
            
            new_sheet.cell(row=1, column=summary_col, value="滾算結果摘要")
            new_sheet.cell(row=2, column=summary_col, value=f"模式: {mode}")
            new_sheet.cell(row=3, column=summary_col, value=f"邊界值: {boundary}")  
            new_sheet.cell(row=4, column=summary_col, value=f"利潤率: {profit_rate}")
            new_sheet.cell(row=5, column=summary_col, value=f"建置容量: {capacity}")
            
            # 在下方填入價金和IRR序列
            if adjustment_record:
                new_sheet.cell(row=7, column=summary_col, value="滾算價金:")
                for i, price in enumerate(adjustment_record[:5]):  # 只顯示前5個
                    new_sheet.cell(row=8+i, column=summary_col, value=price)
            
            if irr_results:
                new_sheet.cell(row=7, column=summary_col+1, value="專案IRR:")
                for i, irr_data in enumerate(irr_results[:5]):  # 只顯示前5個
                    project_irr = irr_data.get("project_irr")
                    if project_irr:
                        new_sheet.cell(row=8+i, column=summary_col+1, value=f"{project_irr:.2f}%")
            
            print(f"已在 {new_sheet.title} 工作表右側空白區域填入滾算結果摘要")
            
        except Exception as e:
            print(f"填入複製工作表數據時發生錯誤: {str(e)}")

    def _fill_copied_sheet_with_rolling_data(self, new_sheet, result, rolling_idx):
        """為特定的滾算次數在複製工作表中填入數據，並從輸入公版填入基礎資料"""
        try:
            adjustment_record = result.get("adjustment_record", [])
            profit_record = result.get("profit_record", [])
            irr_results = result.get("irr_results", [])
            mode = result.get("mode", "")
            step = result.get("step", "")
            boundary = result.get("boundary", "")
            original_data = result.get("original_data", {})
            # 獲取用戶輸入的開發費（優先使用用戶輸入值）
            user_development_fee = result.get("development_fee", original_data.get("development_fee", 0))
            # 獲取用戶輸入的開發費（優先使用用戶輸入值）
            user_development_fee = result.get("development_fee", original_data.get("development_fee", 0))
            
            # === 1. 從輸入公版填入基礎資料到對應位置 ===
            # 輸入公版的B2(項目名稱)填到C2
            project_name = original_data.get("project_name", "")
            if project_name:
                new_sheet.cell(row=2, column=3, value=project_name)  # C2
            
            # 輸入公版的B3(電站壽命)填到C3
            plant_lifetime = original_data.get("plant_lifetime", 20)
            new_sheet.cell(row=3, column=3, value=plant_lifetime)  # C3
            
            # 輸入公版的B4(裝置容量)填到C4 
            capacity = original_data.get("capacity", 436.1)
            new_sheet.cell(row=4, column=3, value=capacity)  # C4
            
            # 輸入公版的B5(起始年度)填到C5
            start_year = original_data.get("start_year", 2020)
            new_sheet.cell(row=5, column=3, value=start_year)  # C5
            
            # 輸入公版的D5(結束年度)填到E5
            end_year = original_data.get("end_year", 2039)
            new_sheet.cell(row=5, column=5, value=end_year)  # E5
            
            # 輸入公版的C10(年度AC併網總發電量)填到D10
            annual_ac_generation = original_data.get("annual_ac_generation", 547002)
            new_sheet.cell(row=10, column=4, value=annual_ac_generation)  # D10
            
            # 輸入公版的D10(電費收入年限)填到E10
            electricity_revenue_years = original_data.get("electricity_revenue_years", 0)
            new_sheet.cell(row=10, column=5, value=electricity_revenue_years)  # E10
            
            # 輸入公版的C11(首年衰退率)填到D11
            first_year_decline_rate = original_data.get("first_year_decline_rate", 0)
            new_sheet.cell(row=11, column=4, value=first_year_decline_rate)  # D11
            
            # 輸入公版的C12(次年衰退率)填到D12
            second_year_decline_rate = original_data.get("fit_rate", 0)
            new_sheet.cell(row=12, column=4, value=second_year_decline_rate)  # D12
            
            # 輸入公版的C13(躉售費率FIT)填到D13
            fit_price_c13 = original_data.get("fit_price_c13", 3.7152)
            new_sheet.cell(row=13, column=4, value=fit_price_c13)  # D13
            
            # 輸入公版的C17~C26填到D17~D26 (使用原有字段名稱)
            new_sheet.cell(row=17, column=4, value=original_data.get("profit_rate", 0))  # D17信邦利潤率
            new_sheet.cell(row=18, column=4, value=user_development_fee)  # D18開發費（使用用戶輸入值）
            new_sheet.cell(row=19, column=4, value=original_data.get("rent_calculation_method", 0))  # D19租金費用租金計算方式
            new_sheet.cell(row=19, column=5, value=original_data.get("rent_calculation_amortization_years", 0))  # E19租金費用租金計算方式支提年限
            new_sheet.cell(row=20, column=4, value=original_data.get("rent_method1_total", 0))  # D20租金費用【方式1參數】租金總額
            new_sheet.cell(row=20, column=5, value=original_data.get("rent_method1_amortization_years", 0))  # E20租金費用【方式1參數】租金總額支提年限
            new_sheet.cell(row=21, column=4, value=original_data.get("rent_method2_ratio", 0))  # D21租金費用【方式2參數】拽成比率
            new_sheet.cell(row=21, column=5, value=original_data.get("rent_method2_amortization_years", 0))  # E21租金費用【方式2參數】拽成比率支提年限
            new_sheet.cell(row=22, column=4, value=original_data.get("maintenance_cost", 0))  # D22運維費用
            new_sheet.cell(row=22, column=5, value=original_data.get("maintenance_amortization_years", 0))  # E22運維費用支提年限
            new_sheet.cell(row=23, column=4, value=original_data.get("insurance_cost", 0))  # D23保險費用
            new_sheet.cell(row=23, column=5, value=original_data.get("insurance_amortization_years", 0))  # E23保險費用支提年限
            new_sheet.cell(row=24, column=4, value=original_data.get("loan_amount", 0))  # D24租金總額
            new_sheet.cell(row=24, column=5, value=original_data.get("recycle_amortization_years", 0))  # E24模組回收費用支提年限
            new_sheet.cell(row=25, column=4, value=original_data.get("bank_interest_rate", 0))  # D25利息費用銀行利率
            new_sheet.cell(row=25, column=5, value=original_data.get("interest_amortization_years", 0))  # E25利息費用支提年限
            new_sheet.cell(row=26, column=4, value=original_data.get("tax_rate", 0))  # D26所得稅率
            new_sheet.cell(row=26, column=5, value=original_data.get("tax_amortization_years", 0))  # E26所得稅支提年限
            
            # 新增：輸入公版C31~C34填到D31~D34
            new_sheet.cell(row=31, column=4, value=original_data.get("c31_value", 0))  # D31: 貸款成數占總設備費用比例
            new_sheet.cell(row=32, column=4, value=original_data.get("c32_value", 0))  # D32: 貸款還款攤還期數
            new_sheet.cell(row=33, column=4, value=original_data.get("c33_value", 0))  # D33: 現金股利股利比率
            new_sheet.cell(row=34, column=4, value=original_data.get("c34_value", 0))  # D34: 年底減資攤還期數
            
            # 獲取特定滾算次數的數據
            if rolling_idx < len(adjustment_record):
                current_price = adjustment_record[rolling_idx]
                current_profit = profit_record[rolling_idx] if rolling_idx < len(profit_record) else 0
                irr_data = irr_results[rolling_idx] if rolling_idx < len(irr_results) else {"project_irr": None, "equity_method_irr": None}
                
                # 獲取利潤率和容量（優先使用用戶輸入值）
                current_profit_rate = original_data.get("profit_rate", 0.2)
                if hasattr(self, 'current_profit_rate') and self.current_profit_rate is not None:
                    current_profit_rate = self.current_profit_rate
                
                # === 2. 填入滾算計算結果到設備成本位置 ===
                # 將當前滾算的設備成本填入對應位置（D16）
                new_sheet.cell(row=16, column=4, value=current_price)  # D16設備成本
                
                # 輸入公版的D16(設備費用攤提年限)填到E16
                equipment_amortization_years = original_data.get("equipment_amortization_years", 20)
                new_sheet.cell(row=16, column=5, value=equipment_amortization_years)  # E16
                
                # 輸入公版的D16(設備費用支提年限)填到E16
                equipment_amortization_years = original_data.get("equipment_amortization_years", 0)
                new_sheet.cell(row=16, column=5, value=equipment_amortization_years)  # E16
                
                # === 3. 填入利潤率（優先使用用戶輸入值） ===
                # 填入到D17信邦利潤率位置
                new_sheet.cell(row=17, column=4, value=current_profit_rate)  # D17信邦利潤率
                
                # === 4. 在右側空白區域填入滾算摘要（保持原有功能） ===
                summary_col = 13  # 列M
                new_sheet.cell(row=5, column=summary_col, value=f"第{rolling_idx}次滾算摘要")
                new_sheet.cell(row=6, column=summary_col, value=f"價金: {current_price}")
                new_sheet.cell(row=7, column=summary_col, value=f"專案IRR: {irr_data.get('project_irr', 'N/A')}")
                new_sheet.cell(row=8, column=summary_col, value=f"權益法IRR: {irr_data.get('equity_method_irr', 'N/A')}")
                new_sheet.cell(row=9, column=summary_col, value=f"利潤率: {current_profit_rate*100:.2f}%")
                
                # 生成備註說明
                mode_desc = {
                    "cash": f"現金模式，每次減少 {step} 元",
                    "ratio": f"比率模式，每次減少 {step*100}%",
                    "conditional": "條件模式，依價金範圍調整步幅",
                    "customize": "自訂模式，自動或手動配置步幅"
                }.get(mode, f"{mode}模式")
                
                if rolling_idx == 0:
                    remark = f"初始價金 - {mode_desc}"
                else:
                    reduction = adjustment_record[rolling_idx-1] - current_price
                    remark = f"第{rolling_idx}次滾算，減少 {reduction} 元 - {mode_desc}"
                
                new_sheet.cell(row=10, column=summary_col, value=remark)
                
                print(f"已在 {new_sheet.title} 工作表填入輸入公版基礎資料和第{rolling_idx}次滾算結果")
            
        except Exception as e:
            print(f"填入複製工作表滾算數據時發生錯誤: {str(e)}")

    def _fill_original_sheet_with_initial_data(self, wb, result):
        """為滾算紀錄1工作表填入初始價金的數據"""
        try:
            if "滾算紀錄1" not in wb.sheetnames:
                print("警告：找不到滾算紀錄1工作表")
                return
                
            sheet = wb["滾算紀錄1"]
            adjustment_record = result.get("adjustment_record", [])
            irr_results = result.get("irr_results", [])
            original_data = result.get("original_data", {})
            # 獲取用戶輸入的開發費（優先使用用戶輸入值）
            user_development_fee = result.get("development_fee", original_data.get("development_fee", 0))
            
            # === 1. 從輸入公版填入基礎資料到對應位置 ===
            # 輸入公版的B2(項目名稱)填到C2
            project_name = original_data.get("project_name", "")
            if project_name:
                sheet.cell(row=2, column=3, value=project_name)  # C2
            
            # 輸入公版的B3(電站壽命)填到C3
            plant_lifetime = original_data.get("plant_lifetime", 20)
            sheet.cell(row=3, column=3, value=plant_lifetime)  # C3
            
            # 輸入公版的B4(裝置容量)填到C4 
            capacity = original_data.get("capacity", 436.1)
            sheet.cell(row=4, column=3, value=capacity)  # C4
            
            # 輸入公版的B5(起始年度)填到C5
            start_year = original_data.get("start_year", 2020)
            sheet.cell(row=5, column=3, value=start_year)  # C5
            
            # 輸入公版的D5(結束年度)填到E5
            end_year = original_data.get("end_year", 2039)
            sheet.cell(row=5, column=5, value=end_year)  # E5
            
            # 輸入公版的C10(年度AC併網總發電量)填到D10
            annual_ac_generation = original_data.get("annual_ac_generation", 547002)
            sheet.cell(row=10, column=4, value=annual_ac_generation)  # D10
            
            # 輸入公版的D10(電費收入年限)填到E10
            electricity_revenue_years = original_data.get("electricity_revenue_years", 0)
            sheet.cell(row=10, column=5, value=electricity_revenue_years)  # E10
            
            # 輸入公版的C11(首年衰退率)填到D11
            first_year_decline_rate = original_data.get("first_year_decline_rate", 0)
            sheet.cell(row=11, column=4, value=first_year_decline_rate)  # D11
            
            # 輸入公版的C12(次年衰退率)填到D12
            second_year_decline_rate = original_data.get("fit_rate", 0)
            sheet.cell(row=12, column=4, value=second_year_decline_rate)  # D12
            
            # 輸入公版的C13(躉售費率FIT)填到D13
            fit_price_c13 = original_data.get("fit_price_c13", 3.7152)
            sheet.cell(row=13, column=4, value=fit_price_c13)  # D13
            
            # 輸入公版的C17~C26填到D17~D26 (使用原有字段名稱)
            sheet.cell(row=17, column=4, value=original_data.get("profit_rate", 0))  # D17信邦利潤率
            sheet.cell(row=18, column=4, value=user_development_fee)  # D18開發費（使用用戶輸入值）
            sheet.cell(row=19, column=4, value=original_data.get("rent_calculation_method", 0))  # D19租金費用租金計算方式
            sheet.cell(row=19, column=5, value=original_data.get("rent_calculation_amortization_years", 0))  # E19租金費用租金計算方式支提年限
            sheet.cell(row=20, column=4, value=original_data.get("rent_method1_total", 0))  # D20租金費用【方式1參數】租金總額
            sheet.cell(row=20, column=5, value=original_data.get("rent_method1_amortization_years", 0))  # E20租金費用【方式1參數】租金總額支提年限
            sheet.cell(row=21, column=4, value=original_data.get("rent_method2_ratio", 0))  # D21租金費用【方式2參數】拽成比率
            sheet.cell(row=21, column=5, value=original_data.get("rent_method2_amortization_years", 0))  # E21租金費用【方式2參數】拽成比率支提年限
            sheet.cell(row=22, column=4, value=original_data.get("maintenance_cost", 0))  # D22運維費用
            sheet.cell(row=22, column=5, value=original_data.get("maintenance_amortization_years", 0))  # E22運維費用支提年限
            sheet.cell(row=23, column=4, value=original_data.get("insurance_cost", 0))  # D23保險費用
            sheet.cell(row=23, column=5, value=original_data.get("insurance_amortization_years", 0))  # E23保險費用支提年限
            sheet.cell(row=24, column=4, value=original_data.get("loan_amount", 0))  # D24租金總額
            sheet.cell(row=24, column=5, value=original_data.get("recycle_amortization_years", 0))  # E24模組回收費用支提年限
            sheet.cell(row=25, column=4, value=original_data.get("bank_interest_rate", 0))  # D25利息費用銀行利率
            sheet.cell(row=25, column=5, value=original_data.get("interest_amortization_years", 0))  # E25利息費用支提年限
            sheet.cell(row=26, column=4, value=original_data.get("tax_rate", 0))  # D26所得稅率
            sheet.cell(row=26, column=5, value=original_data.get("tax_amortization_years", 0))  # E26所得稅支提年限
            
            # 輸入公版的C10(年度AC併網總發電量)填到D10
            annual_ac_generation = original_data.get("annual_ac_generation", 547002)
            sheet.cell(row=10, column=4, value=annual_ac_generation)  # D10
            
            # 新增：輸入公版C31~C34填到D31~D34
            sheet.cell(row=31, column=4, value=original_data.get("c31_value", 0))  # D31: 貸款成數占總設備費用比例
            sheet.cell(row=32, column=4, value=original_data.get("c32_value", 0))  # D32: 貸款還款攤還期數
            sheet.cell(row=33, column=4, value=original_data.get("c33_value", 0))  # D33: 現金股利股利比率
            sheet.cell(row=34, column=4, value=original_data.get("c34_value", 0))  # D34: 年底減資攤還期數
            
            # === 2. 填入初始價金數據（第0次，即初始價金） ===
            if adjustment_record and irr_results:
                initial_price = adjustment_record[0]
                initial_irr = irr_results[0] if len(irr_results) > 0 else {"project_irr": None, "equity_method_irr": None}
                
                # 獲取利潤率（優先使用用戶輸入值）
                current_profit_rate = original_data.get("profit_rate", 0.2)
                if hasattr(self, 'current_profit_rate') and self.current_profit_rate is not None:
                    current_profit_rate = self.current_profit_rate
                
                # === 3. 填入初始設備成本到D16 ===
                sheet.cell(row=16, column=4, value=initial_price)  # D16設備成本
                
                # 輸入公版的D16(設備費用攤提年限)填到E16
                equipment_amortization_years = original_data.get("equipment_amortization_years", 20)
                sheet.cell(row=16, column=5, value=equipment_amortization_years)  # E16
                
                # 輸入公版的D16(設備費用支提年限)填到E16
                equipment_amortization_years = original_data.get("equipment_amortization_years", 0)
                sheet.cell(row=16, column=5, value=equipment_amortization_years)  # E16
                
                # === 4. 填入利潤率到D17 ===
                sheet.cell(row=17, column=4, value=current_profit_rate)  # D17信邦利潤率
                
                # === 5. 在右側空白區域填入初始價金摘要 ===
                summary_col = 13  # 列M
                sheet.cell(row=5, column=summary_col, value="初始價金摘要")
                sheet.cell(row=6, column=summary_col, value=f"初始價金: {initial_price}")
                sheet.cell(row=7, column=summary_col, value=f"專案IRR: {initial_irr.get('project_irr', 'N/A')}")
                sheet.cell(row=8, column=summary_col, value=f"權益法IRR: {initial_irr.get('equity_method_irr', 'N/A')}")
                sheet.cell(row=9, column=summary_col, value=f"利潤率: {current_profit_rate*100:.2f}%")
                sheet.cell(row=10, column=summary_col, value="基準價金 - 滾算起始點")
                
                print(f"已在 滾算紀錄1 工作表填入輸入公版基礎資料和初始價金數據")
        
        except Exception as e:
            print(f"填入滾算紀錄1工作表數據時發生錯誤: {str(e)}")

    def _write_results_to_excel(self, output_path: str, results: List[Dict[str, Any]]):
        """將結果寫入滾算紀錄單(總紀錄)工作表，並複製滾算紀錄工作表"""
        max_retries = 3
        retry_delay = 1
        
        for attempt in range(max_retries):
            try:
                wb = load_workbook(output_path)
                
                # === 1. 寫入滾算紀錄單(總紀錄)工作表 - 保持原有功能 ===
                if "滾算紀錄單(總紀錄)" in wb.sheetnames:
                    sheet = wb["滾算紀錄單(總紀錄)"]
                else:
                    # 如果沒有滾算紀錄單(總紀錄)，創建一個
                    sheet = wb.create_sheet("滾算紀錄單(總紀錄)")
                    # 設定標題行
                    headers = ["價金每KW", "專案法IRR", "權益法IRR", "信邦利潤率", "信邦利潤", "每kw信邦利潤", "開發費", "備註"]
                    for col, header in enumerate(headers, 1):
                        sheet.cell(row=1, column=col, value=header)
                
                # 找到下一個空行開始寫入
                start_row = 2
                while sheet.cell(row=start_row, column=1).value is not None:
                    start_row += 1
                
                # 寫入滾算結果，並為每次滾算創建對應的工作表
                current_row = start_row
                for result_index, result in enumerate(results):
                    adjustment_record = result.get("adjustment_record", [])
                    profit_record = result.get("profit_record", [])
                    irr_results = result.get("irr_results", [])
                    mode = result.get("mode", "")
                    step = result.get("step", "")
                    boundary = result.get("boundary", "")
                    # 使用用戶輸入的開發費（result層級），而非Excel原始數據
                    development_fee = result.get("development_fee", 0)
                    
                    # === 2. 為滾算紀錄1工作表填入初始價金數據 ===
                    self._fill_original_sheet_with_initial_data(wb, result)
                    
                    # === 3. 新增功能：為每次滾算複製一個工作表（除了第0次初始價金） ===
                    for rolling_idx in range(len(adjustment_record)):
                        if rolling_idx == 0:
                            # 第0次是初始價金，使用現有的滾算紀錄1，不需要複製
                            continue
                            
                        # 為每次滾算創建工作表，工作表編號從1開始（對應第1次滾算）
                        rolling_sheet_number = rolling_idx  # rolling_idx=1時對應開始滾算紀錄1
                        new_sheet_name = f"開始滾算紀錄{rolling_sheet_number}"
                        
                        # 複製滾算紀錄工作表
                        if "滾算紀錄1" in wb.sheetnames:
                            source_sheet = wb["滾算紀錄1"]
                            # 複製工作表
                            new_sheet = wb.copy_worksheet(source_sheet)
                            new_sheet.title = new_sheet_name
                            print(f"已複製滾算紀錄1工作表為: {new_sheet_name}")
                            
                            # 為這個特定的滾算填入數據
                            self._fill_copied_sheet_with_rolling_data(new_sheet, result, rolling_idx)
                        else:
                            print("警告：找不到滾算紀錄1工作表，無法複製")
                    profit_record = result.get("profit_record", [])
                    irr_results = result.get("irr_results", [])
                    mode = result.get("mode", "")
                    step = result.get("step", "")
                    boundary = result.get("boundary", "")
                    # 使用用戶輸入的開發費（result層級），而非Excel原始數據
                    development_fee = result.get("development_fee", 0)
                    
                    # 生成備註說明
                    mode_desc = {
                        "cash": f"現金模式，每次減少 {step} 元",
                        "ratio": f"比率模式，每次減少 {step*100}%",
                        "conditional": "條件模式，依價金範圍調整步幅",
                        "customize": "自訂模式，自動或手動配置步幅"
                    }.get(mode, f"{mode}模式")
                    
                    # === 寫入滾算紀錄單(總紀錄)的結果 ===
                    for i, (price, profit) in enumerate(zip(adjustment_record, profit_record)):
                        # 取得對應的IRR結果
                        irr_data = irr_results[i] if i < len(irr_results) else {"project_irr": None, "equity_method_irr": None}
                        
                        if i == 0:
                            remark = f"初始價金 - {mode_desc}，邊界值 {boundary} 元"
                        else:
                            reduction = adjustment_record[i-1] - price
                            remark = f"第{i}次滾算，減少 {reduction} 元 - {mode_desc}"
                        
                        # 獲取利潤率（優先使用當次執行傳入的利潤率，其次使用Excel原始數據）
                        # 從result中獲取動態利潤率，而不是使用固定的original_data
                        current_profit_rate = result.get("original_data", {}).get("profit_rate", 0.2)
                        # 檢查是否有外部傳入的利潤率參數覆蓋
                        if hasattr(self, 'current_profit_rate') and self.current_profit_rate is not None:
                            current_profit_rate = self.current_profit_rate
                        
                        # 獲取Excel建置容量和攤提年限
                        capacity = result.get("original_data", {}).get("capacity", 436.1)
                        equipment_amortization_years = result.get("original_data", {}).get("equipment_amortization_years", 1)
                        
                        # 計算每kw信邦利潤：基礎利潤 - 每年開發費
                        # 基礎利潤 = equipment_cost / (1 - profit_rate) - equipment_cost
                        # development_fee是每年開發費，總開發費 = development_fee * equipment_amortization_years
                        base_profit_per_kw = price / (1 - current_profit_rate) - price
                        annual_development_cost = development_fee  # 每年開發費
                        profit_per_kw = base_profit_per_kw - annual_development_cost  # 信邦每KW利潤(年)
                        
                        # 計算信邦利潤總額：信邦每KW利潤(年) × 建置容量
                        calculated_profit = profit_per_kw * capacity
                        
                        # 按照滾算紀錄單格式填入
                        sheet.cell(row=current_row, column=1, value=price)  # 價金每KW
                        sheet.cell(row=current_row, column=2, value=irr_data.get("project_irr"))  # 專案法IRR
                        sheet.cell(row=current_row, column=3, value=irr_data.get("equity_method_irr"))  # 權益法IRR
                        
                        # 設置利潤率，直接使用小數值並設置百分比格式（例如0.15顯示為15.00%）
                        profit_rate_cell = sheet.cell(row=current_row, column=4, value=current_profit_rate) # 信邦利潤率
                        profit_rate_cell.number_format = '0.00%'  # 設置為百分比格式，保留兩位小數
                        
                        sheet.cell(row=current_row, column=5, value=calculated_profit)  # 信邦利潤(總利潤)
                        sheet.cell(row=current_row, column=6, value=profit_per_kw)      # 每kw信邦利潤
                        sheet.cell(row=current_row, column=7, value=development_fee)  # 開發費
                        sheet.cell(row=current_row, column=8, value=remark) # 備註
                        current_row += 1
                    
                    # 添加空行分隔不同的滾算結果
                    current_row += 1
                
                wb.save(output_path)
                wb.close()
                return  # 成功完成，退出重試循環
                
            except PermissionError as e:
                if attempt < max_retries - 1:
                    print(f"檔案被鎖定，等待 {retry_delay} 秒後重試 (第 {attempt + 1}/{max_retries} 次)")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # 指數退避
                else:
                    raise Exception(f"檔案被鎖定，無法寫入。請確保 Excel 檔案未被其他程式開啟: {str(e)}")
            except Exception as e:
                raise Exception(f"寫入 Excel 檔案失敗: {str(e)}")

    def execute_price_rolling(self, 
                              mode: str = "cash",
                              boundary: int = 20000,
                              step: Union[int, float] = 1000,
                              profit_rate: float = None,
                              development_fee: int = None,
                              adjustment_times: int = 10,
                              sheet_name: str = None,
                              **kwargs) -> Dict[str, Any]:
        """
        執行價金滾算

        Args:
            mode: 滾算模式 ("cash", "ratio", "conditional", "customize")
            boundary: 價金調整邊界
            step: 調整步幅（現金模式為整數，比率模式為小數）
            profit_rate: 信邦利潤率（如未指定，將從Excel檔案讀取）
            development_fee: 開發費用（如未指定，將從Excel檔案讀取）
            adjustment_times: 調整次數（自訂模式使用）
            **kwargs: 其他模式特定參數

        Returns:
            執行結果
        """
        try:
            # 1. 讀取 Excel 資料夾中的檔案
            excel_file = self._find_excel_file(self.excel_folder)
            if not excel_file:
                return {
                    "success": False,
                    "message": f"在 {self.excel_folder} 資料夾中找不到 Excel 檔案"
                }

            # 2. 讀取數據
            data = self._read_excel_data(excel_file)
            initial_cost = int(data.get("equipment_cost", 30000))
            
            # 3. 使用Excel中的預設值，除非用戶明確指定
            if profit_rate is None:
                profit_rate = float(data.get("profit_rate", 0.1))
                print(f"使用Excel預設利潤率: {profit_rate}")
            else:
                print(f"使用用戶指定利潤率: {profit_rate}")
            
            # 設定當前利潤率，供Excel寫入時使用
            self.current_profit_rate = profit_rate
            
            if development_fee is None:
                development_fee = int(data.get("development_fee", 2000))
                print(f"使用Excel預設開發費: {development_fee}")
            else:
                print(f"使用用戶指定開發費: {development_fee}")

            # 4. 執行滾算計算
            results = []
            
            if mode == "cash":
                calculator = CashMode(boundary=boundary, step=int(step))
                adjustment_record = calculator.calculation(initial_cost)
                
            elif mode == "ratio":
                calculator = RatioMode(boundary=boundary, step=float(step))
                adjustment_record = calculator.calculation(initial_cost)
                
            elif mode == "conditional":
                maximum_value = kwargs.get("maximum_value", 50000)
                minimum_value = kwargs.get("minimum_value", 30000)
                condition_step_1 = kwargs.get("condition_step_1", 2000)
                condition_step_2 = kwargs.get("condition_step_2", 1000)
                condition_step_3 = kwargs.get("condition_step_3", 500)
                
                calculator = ConditionalMode(
                    boundary=boundary,
                    maximum_value=maximum_value,
                    minimum_value=minimum_value,
                    condition_step_1=condition_step_1,
                    condition_step_2=condition_step_2,
                    condition_step_3=condition_step_3
                )
                adjustment_record = calculator.calculation(initial_cost)
                
            elif mode == "customize":
                calculator = CustomizeMode(boundary=boundary, adjustment_times=adjustment_times)
                if kwargs.get("auto_config", True):
                    steps = calculator.automatic_configuration(initial_cost)
                else:
                    steps = kwargs.get("custom_steps", [1000] * adjustment_times)
                adjustment_record = calculator.calculation(steps, initial_cost)
                
            else:
                return {
                    "success": False,
                    "message": f"不支援的滾算模式: {mode}"
                }

            # 4. 計算成本結構調整
            cost_service = CostStructureService(profit_rate=profit_rate, development_fee=development_fee)
            cost_structure_adjusted = cost_service.equipment_cost_calculation(adjustment_record)
            profit_record = cost_service.get_profit(adjustment_record)
            
            # 5. 計算IRR
            print("正在計算IRR...")
            irr_results = self._calculate_irr_for_prices(excel_file, adjustment_record, profit_rate, development_fee, sheet_name)
            
            # 6. 準備結果數據
            result = {
                "mode": mode,
                "initial_cost": initial_cost,
                "boundary": boundary,
                "step": step,
                "adjustment_record": adjustment_record,
                "cost_structure_adjusted": cost_structure_adjusted,
                "profit_record": profit_record,
                "irr_results": irr_results,
                "development_fee": development_fee,
                "original_data": data
            }
            results.append(result)

            # 7. 準備輸出檔案並寫入結果
            output_path = self._prepare_output_file(excel_file)
            self._write_results_to_excel(output_path, results)

            return {
                "success": True,
                "message": "價金滾算執行完成",
                "result": result,
                "output_file": output_path,
                "summary": {
                    "initial_cost": initial_cost,
                    "final_cost": adjustment_record[-1] if adjustment_record else initial_cost,
                    "adjustment_count": len(adjustment_record) - 1,
                    "total_reduction": initial_cost - (adjustment_record[-1] if adjustment_record else initial_cost)
                }
            }

        except Exception as e:
            return {
                "success": False,
                "message": f"執行價金滾算時發生錯誤: {str(e)}"
            }


# 定義工具 Schema
EQUIPMENT_COST_TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "execute_price_rolling",
            "description": "【完整流程】執行設備成本價金滾算，包含完整流程：1) 從 Excel 資料夾讀取數據, 2) 執行價金滾算計算, 3) 計算每個價金的 IRR, 4) 將結果寫入 'Excel final/滾算後記錄_時間戳.xlsx'。此工具會自動保存結果檔案。如只需要計算分析不寫入檔案，請使用 calculate_price_rolling 工具。",
            "parameters": {
                "type": "object",
                "properties": {
                    "mode": {
                        "type": "string",
                        "enum": ["cash", "ratio", "conditional", "customize"],
                        "description": "滾算模式：cash(現金模式), ratio(比率模式), conditional(條件模式), customize(自訂模式)"
                    },
                    "boundary": {
                        "type": "integer",
                        "description": "價金調整的邊界值"
                    },
                    "step": {
                        "type": "number",
                        "description": "調整步幅，現金模式使用整數，比率模式使用 0-1 之間的小數"
                    },
                    "profit_rate": {
                        "type": "number",
                        "description": "信邦利潤率 (0-1 之間的小數)，如未指定將從Excel檔案讀取預設值"
                    },
                    "development_fee": {
                        "type": "integer",
                        "description": "開發費用，如未指定將從Excel檔案讀取預設值"
                    },
                    "adjustment_times": {
                        "type": "integer",
                        "description": "調整次數（自訂模式使用）"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "用於IRR計算的Excel工作表名稱（可選）"
                    },
                    "maximum_value": {
                        "type": "integer",
                        "description": "條件模式：判斷價金範圍的最大值"
                    },
                    "minimum_value": {
                        "type": "integer",
                        "description": "條件模式：判斷價金範圍的最小值"
                    },
                    "condition_step_1": {
                        "type": "integer",
                        "description": "條件模式：價金 > maximum_value 時的 step"
                    },
                    "condition_step_2": {
                        "type": "integer",
                        "description": "條件模式：minimum_value <= 價金 <= maximum_value 時的 step"
                    },
                    "condition_step_3": {
                        "type": "integer",
                        "description": "條件模式：價金 < minimum_value 時的 step"
                    },
                    "auto_config": {
                        "type": "boolean",
                        "description": "自訂模式：是否自動配置 steps"
                    },
                    "custom_steps": {
                        "type": "array",
                        "items": {
                            "type": "integer"
                        },
                        "description": "自訂模式：手動指定的 steps 列表"
                    }
                },
                "required": ["mode"]
            }
        }
    }
]