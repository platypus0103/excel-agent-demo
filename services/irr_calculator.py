import numpy as np
import numpy_financial as npf
from typing import List, Tuple, Dict, Any
from models.irr_models_v2 import (
    EquipmentCostParams, IncomeData, ExpenseData, InterestData,
    IRRCalculationRequest, IRRCalculationResponse, CashFlowItem,
    RangeData, KWBasedData, BankLoanData, DecayData, AreaBasedData,
    CashFlowStatementParams, CashFlowStatementItem, IRRAnalysis
)

"""
IRR 計算服務
將原始 JavaScript 邏輯移植到 Python，使用 numpy-financial 優化計算
"""
class IRRCalculatorService:
    """IRR 計算服務類"""

    @staticmethod
    def calculate_equipment_cost(params: EquipmentCostParams) -> float:
        """
        計算設備費用總額
        公式：報價總金額 = {[每KW價格 ÷ (1 - 利潤率%)] + 開發費} × 建置容量
        """
        try:
            # 調整後的每KW價格 = 原價格 ÷ (1 - 利潤率%)
            adjusted_price_per_kw = params.price_per_kw / (1 - params.profit_rate / 100)

            # 總費用 = (調整後每KW價格 + 開發費) × 建置容量
            total_cost = (adjusted_price_per_kw + params.development_fee) * params.capacity

            return total_cost
        except Exception as e:
            raise ValueError(f"設備費用計算錯誤: {str(e)}")

    @staticmethod
    def calculate_interest_values(interest_data: InterestData, equipment_cost: float, target_length: int, years: List[int]) -> List[float]:
        """
        計算利息費用
        支援無利息和銀行貸款兩種模式
        """
        if interest_data.no_interest:
            # 無利息模式，返回全零列表
            return [0.0] * target_length

        # 銀行貸款模式
        bank_loan = interest_data.bank_loan_data
        if not bank_loan:
            raise ValueError("銀行貸款模式下必須提供貸款參數")

        # 計算貸款金額
        loan_amount = equipment_cost * (bank_loan.loan_ratio / 100)

        # 計算每年本金攤還
        annual_principal = loan_amount / bank_loan.repayment_period

        # 計算每年利息
        interest_values = []
        for i in range(target_length):
            year_index = i + 1  # 第幾年 (1, 2, 3...)

            if year_index <= bank_loan.repayment_period:
                # 剩餘本金 = 貸款金額 - (已攤還年數 × 每年攤還本金)
                remaining_principal = loan_amount - (year_index - 1) * annual_principal
                # 年利息 = 剩餘本金 × 利率
                annual_interest = remaining_principal * (bank_loan.bank_rate / 100)
                interest_values.append(annual_interest)
            else:
                # 攤還期數結束後，無利息
                interest_values.append(0.0)

        return interest_values

    @staticmethod
    def calculate_cash_flow_statement(
        cash_flow_items: List[CashFlowItem],
        equipment_cost: float,
        years: List[int],
        cash_flow_params: CashFlowStatementParams,
        interest_data: InterestData
    ) -> Tuple[List[CashFlowStatementItem], IRRAnalysis]:
        """
        計算現金流量表
        """
        year_count = len(years)
        cash_flow_statement = []

        # 計算基本參數
        if interest_data.no_interest:
            loan_amount = 0
            annual_repayment = 0
            repayment_period = 0
        else:
            loan_amount = equipment_cost * (interest_data.bank_loan_data.loan_ratio / 100)
            repayment_period = interest_data.bank_loan_data.repayment_period
            annual_repayment = loan_amount / repayment_period if repayment_period > 0 else 0

        cash_capital_increase = equipment_cost - loan_amount
        annual_capital_reduction = cash_capital_increase / cash_flow_params.capital_reduction_period

        # IRR現金流收集
        cost_method_cash_flows = []
        equity_method_cash_flows = []

        # 期初現金流
        opening_cash_flow = 0.0

        for i in range(year_count):
            cf = cash_flow_items[i]
            year = years[i]

            # === 營運活動 ===
            aftertax_net_profit = cf.after_tax_cash_flow
            equipment_depreciation = cf.equipment_depreciation
            operating_cash_flow = aftertax_net_profit + equipment_depreciation

            # === 投資活動 ===
            equipment_expenditure = -equipment_cost if i == 0 else 0.0

            # === 理財活動 ===
            loan_financing = loan_amount if i == 0 else 0.0
            # 還款從第二年開始
            loan_repayment = -annual_repayment if i > 0 and i <= repayment_period else 0.0
            cash_capital_increase_flow = cash_capital_increase if i == 0 else 0.0

            # 現金股利 = 前一年稅後淨利 × 股利比率（第一年不支出）
            if i > 0:
                prev_aftertax_profit = cash_flow_items[i-1].after_tax_cash_flow
                cash_dividend = -(prev_aftertax_profit * cash_flow_params.dividend_ratio / 100)
            else:
                cash_dividend = 0.0

            # 年底減資：在最後N年分攤
            capital_reduction_start = year_count - cash_flow_params.capital_reduction_period
            if i >= capital_reduction_start:
                capital_reduction = -annual_capital_reduction
            else:
                capital_reduction = 0.0

            # === 現金流匯總 ===
            net_cash_inflow = (operating_cash_flow + equipment_expenditure + loan_financing +
                             loan_repayment + cash_capital_increase_flow + cash_dividend + capital_reduction)

            closing_cash_flow = opening_cash_flow + net_cash_inflow

            # === IRR分析 ===
            # 成本法現金流
            if i == 0:
                cost_method_cash_flow = -cash_capital_increase  # 初始投資
            else:
                cost_method_cash_flow = -cash_dividend - capital_reduction  # 轉為正數

            # 權益法現金流
            if i == 0:
                equity_method_cash_flow = aftertax_net_profit - cash_capital_increase  # 初始投資
            else:
                equity_method_cash_flow = aftertax_net_profit + abs(capital_reduction)  # 稅後淨利 + 年底減資的絕對值

            # === 借款餘額 ===
            if i == 0:
                loan_balance = loan_amount
            else:
                prev_balance = cash_flow_statement[i-1].loan_balance if i > 0 else loan_amount
                loan_balance = prev_balance + loan_repayment  # loan_repayment已是負數

            # 收集IRR現金流
            cost_method_cash_flows.append(cost_method_cash_flow)
            equity_method_cash_flows.append(equity_method_cash_flow)

            # 創建現金流量表項目
            stmt_item = CashFlowStatementItem(
                year=year,
                aftertax_net_profit=aftertax_net_profit,
                equipment_depreciation=equipment_depreciation,
                operating_cash_flow=operating_cash_flow,
                equipment_expenditure=equipment_expenditure,
                loan_financing=loan_financing,
                loan_repayment=loan_repayment,
                cash_capital_increase=cash_capital_increase_flow,
                cash_dividend=cash_dividend,
                capital_reduction=capital_reduction,
                net_cash_inflow=net_cash_inflow,
                opening_cash_flow=opening_cash_flow,
                closing_cash_flow=closing_cash_flow,
                cost_method_cash_flow=cost_method_cash_flow,
                equity_method_cash_flow=equity_method_cash_flow,
                loan_balance=loan_balance
            )

            cash_flow_statement.append(stmt_item)
            opening_cash_flow = closing_cash_flow

        # 計算額外一年的成本法現金流
        if year_count > 0:
            # 成本法最後一年現金流 = 前一年的期末淨現金流
            last_closing_cash_flow = cash_flow_statement[-1].closing_cash_flow
            additional_cost_method_flow = last_closing_cash_flow

            # 添加到IRR現金流
            cost_method_cash_flows.append(additional_cost_method_flow)
            # 權益法不需要額外一年，所以不添加

        # 計算IRR
        cost_method_success, cost_method_irr, _ = IRRCalculatorService.calculate_irr_numpy(cost_method_cash_flows)
        equity_method_success, equity_method_irr, _ = IRRCalculatorService.calculate_irr_numpy(equity_method_cash_flows)

        irr_analysis = IRRAnalysis(
            cost_method_irr=cost_method_irr if cost_method_success else None,
            equity_method_irr=equity_method_irr if equity_method_success else None,
            cost_method_cash_flows=cost_method_cash_flows,
            equity_method_cash_flows=equity_method_cash_flows
        )

        return cash_flow_statement, irr_analysis

    @staticmethod
    def process_data_input(
    data: IncomeData | ExpenseData, 
    target_length: int, 
    years: List[int], 
    equipment_capacity: float = 0,
    kwh_values: List[float] = None,   
    final_tariff: float = 0           
) -> List[float]:
        """
        處理輸入數據，統一轉換為指定長度的數值列表
        支援多種模式：range, kw_based, decay, area_based, new_project
        """
        if data.mode == "range":
            # 年份範圍攤平模式
            annual_amount = data.range_data.annual_amount
            start_year = data.range_data.start_year
            end_year = data.range_data.end_year

            # 創建全零數組
            result = [0.0] * target_length

            # 在指定年份範圍內填入年金額
            for i, year in enumerate(years):
                if start_year <= year <= end_year:
                    result[i] = annual_amount

            return result

        elif data.mode == "kw_based":
            # 基於KW計算模式
            price_per_kw = data.kw_based_data.price_per_kw
            start_year = data.kw_based_data.start_year
            end_year = data.kw_based_data.end_year

            # 計算總金額 = 每KW價格 × 建置容量
            total_amount = price_per_kw * equipment_capacity

            # 創建全零數組
            result = [0.0] * target_length

            # 計算攤平年數和每年金額
            range_years = end_year - start_year + 1
            amount_per_year = total_amount / range_years if range_years > 0 else 0

            # 在指定年份範圍內填入攤平金額
            for i, year in enumerate(years):
                if start_year <= year <= end_year:
                    result[i] = amount_per_year

            return result

        elif data.mode == "decay":
            # 首年發電推算模式
            first_year_amount = data.decay_data.first_year_amount
            decay_rate = data.decay_data.decay_rate / 100  # 轉換為小數
            start_year = data.decay_data.start_year
            end_year = data.decay_data.end_year

            result = [0.0] * target_length
            current_amount = first_year_amount

            for i, year in enumerate(years):
                if start_year <= year <= end_year:
                    if year == start_year:
                        result[i] = first_year_amount
                        current_amount = first_year_amount * (1 - decay_rate)
                    else:
                        result[i] = current_amount
                        current_amount *= (1 - decay_rate)
            return result

        elif data.mode == "area_based":
            # 基於面積計算模式
            area_data = data.area_based_data
            
            # 單位換算因子 (全部換算成公頃)
            conversion_factors = {
                "hectare": 1.0,
                "jia": 0.9699,      # 1 甲 約 0.9699 公頃
                "ping": 0.000330579, # 1 坪 約 0.00033 公頃
                "sq_meter": 0.0001   # 1 平方公尺 = 0.0001 公頃
            }
            
            # 如果價格是基於公頃，則將面積轉換為公頃
            area_in_hectares = area_data.area * conversion_factors[area_data.unit]
            
            # 年租金 = 面積(公頃) * 每公頃年租金
            annual_rent = area_in_hectares * area_data.price_per_area
            
            result = [0.0] * target_length
            for i, year in enumerate(years):
                if area_data.start_year <= year <= area_data.end_year:
                    result[i] = annual_rent
            return result

        
        elif data.mode == "new_project":

            if not hasattr(data, 'new_project_data') or data.new_project_data is None:
                raise ValueError("Income mode is 'new_project', but 'new_project_data' is missing.")

            gen_data = data.new_project_data

            first_year_est = float(gen_data.first_year_est)
            # 將 % 轉換為小數
            first_year_decay_rate = float(gen_data.first_year_decay) / 100.0
            remaining_decay_rate = float(gen_data.remaining_years_decay) / 100.0

            gen_start_year = int(gen_data.start_year)
            gen_end_year = int(gen_data.end_year)

            result = [0.0] * target_length

            # 用於追蹤計算出的kwh，以便下一年可以參考
            # 鍵為年份，值為kwh
            annual_kwh_map = {}

            for i, year in enumerate(years):
                # 檢查今年是否在發電區間內
                if gen_start_year <= year <= gen_end_year:

                    # 這是發電的第幾年 (0 = 第1年, 1 = 第2年, ...)
                    generation_year_index = year - gen_start_year

                    current_kwh = 0.0
                    if generation_year_index == 0:
                        # 第 1 年: 首年預估 × (1 - 首年衰退率)
                        current_kwh = first_year_est * (1 - first_year_decay_rate)
                    else:
                        # 第 2 年起: 前一年 × (1 - 次年衰退率)
                        prev_kwh = annual_kwh_map.get(year - 1, 0)
                        current_kwh = prev_kwh * (1 - remaining_decay_rate)

                    annual_kwh_map[year] = current_kwh
                    result[i] = round(current_kwh, 2)

            return result
        elif data.mode == "electricity_based":
            # 基於電費收入百分比計算模式
            if not hasattr(data, 'electricity_based_data') or data.electricity_based_data is None:
                raise ValueError("Expense mode is 'electricity_based', but 'electricity_based_data' is missing.")

            # 此模式依賴於 calculate_irr_full 函數中傳遞的 kwh_values 和 final_tariff
            if kwh_values is None or final_tariff == 0 or len(kwh_values) != target_length:
                 raise ValueError("Electricity-based mode requires valid KWH values and final tariff for calculation.")

            elec_data = data.electricity_based_data

            # 獲取設定值
            revenue_percent = elec_data.revenue_percent / 100.0
            start_year = elec_data.start_year
            end_year = elec_data.end_year

            result = [0.0] * target_length

            for i, year in enumerate(years):
                if start_year <= year <= end_year:
                    # 計算當年度售電收入
                    annual_revenue = kwh_values[i] * final_tariff
                    # 維護費用 = 售電收入 × 設定比例
                    maintenance_fee = annual_revenue * revenue_percent
                    result[i] = maintenance_fee

            return result

        else:
            raise ValueError(f"不支援的模式: {data.mode}")

    @staticmethod
    def calculate_irr_numpy(cash_flows: List[float]) -> Tuple[bool, float, str]:
        """
        使用 numpy-financial 計算 IRR
        """
        try:
            # 數據驗證
            if len(cash_flows) < 2:
                return False, 0, "現金流數據不足，至少需要2個數據點"

            has_positive = any(cf > 0 for cf in cash_flows)
            has_negative = any(cf < 0 for cf in cash_flows)

            if not has_positive or not has_negative:
                return False, 0, "現金流必須包含正值和負值"

            # 使用 numpy-financial 計算 IRR
            cash_flows_array = np.array(cash_flows)
            irr_result = npf.irr(cash_flows_array)

            # 檢查結果是否有效
            if np.isnan(irr_result) or np.isinf(irr_result):
                return False, 0, "無法計算有效的 IRR 值"

            # 轉換為百分比
            irr_percentage = irr_result * 100

            return True, irr_percentage, None

        except Exception as e:
            return False, 0, f"IRR 計算錯誤: {str(e)}"

    @staticmethod
    def calculate_irr_full(request: IRRCalculationRequest) -> IRRCalculationResponse:
        """
        完整的 IRR 計算流程
        """
        try:
            # 1. 計算設備費用
            equipment_cost = IRRCalculatorService.calculate_equipment_cost(request.equipment_params)

            # 2. 確定計算年數
            years = list(range(request.start_year, request.end_year + 1))
            year_count = len(years)

            # 檢查年數是否有效
            if year_count <= 0:
                raise ValueError("計算年數必須大於零")

            # 3. 處理所有輸入數據（傳入設備容量以支援KW計算）
            equipment_capacity = request.equipment_params.capacity

            # 先取得發電度數（KWH）
            kwh_values = IRRCalculatorService.process_data_input(request.income, year_count, years, equipment_capacity)
            # 乘上最終躉售費率
            final_tariff = request.income.final_tariff
            income_values = [kwh * final_tariff for kwh in kwh_values]
            interest_values = IRRCalculatorService.calculate_interest_values(request.interest, equipment_cost, year_count, years)
            rent_values = IRRCalculatorService.process_data_input(request.rent, year_count, years, equipment_capacity)
            maintenance_values = IRRCalculatorService.process_data_input(
                request.maintenance, 
                year_count, 
                years, 
                equipment_capacity,
                kwh_values=kwh_values,       # 傳遞發電量數組
                final_tariff=final_tariff    # 傳遞售電費率
            )
            insurance_values = IRRCalculatorService.process_data_input(request.insurance, year_count, years, equipment_capacity)
            recycling_values = IRRCalculatorService.process_data_input(request.recycling, year_count, years, equipment_capacity)
            custom_values = IRRCalculatorService.process_data_input(request.custom, year_count, years, equipment_capacity)

            # 4. 計算現金流
            cash_flows = [-equipment_cost]  # 初始投資為負值（重資打包賣价）
            cash_flow_items = []

            # 計算設備折舊（重資打包賣價平攤到每一年）
            equipment_depreciation_per_year = equipment_cost / year_count

            for i in range(year_count):
                # === 表格顯示用的計算 ===
                # 稅前淨利 = 電費收入 - 設備折舊 - 利息 - 租金 - 運維 - 保險 - 回收費 - 自定義
                pretax_net_profit = (income_values[i] - equipment_depreciation_per_year - interest_values[i] -
                                       rent_values[i] - maintenance_values[i] - insurance_values[i] - recycling_values[i] - custom_values[i])

                # 表格顯示用所得稅 = 稅前淨利 × 稅率（利息可抵稅）
                display_tax_amount = max(0, pretax_net_profit * request.tax_rate / 100) if pretax_net_profit > 0 else 0

                # 稅後淨利 = 稅前淨利 - 所得稅
                aftertax_net_profit = pretax_net_profit - display_tax_amount

                # === IRR計算用的邏輯（利息不可抵稅）===
                # 所得稅計算時，首年租金為 0
                rent_for_tax = 0 if i == 0 else rent_values[i]

                # IRR稅前淨利（不含自定義欄位）= 電費收入 - 設備折舊 - 利息 - 租金(首年為0) - 運維 - 保險 - 回收費
                pretax_net_profit_for_irr = (income_values[i] - equipment_depreciation_per_year - interest_values[i] -
                                            rent_for_tax - maintenance_values[i] - insurance_values[i] - recycling_values[i])

                # IRR所得稅基數 = 稅前淨利 + 利息費用（利息不可抵稅，故加回）
                irr_tax_base = pretax_net_profit_for_irr + interest_values[i]

                # IRR所得稅 = 稅基 × 稅率
                irr_tax_amount = max(0, irr_tax_base * request.tax_rate / 100) if irr_tax_base > 0 else 0

                # IRR現金流 = 電費收入 - 租金(每年都算) - 運維 - 保險 - 回收費 - IRR所得稅
                # 注意：不包含設備折舊（已在初始投資中）、不包含利息、不包含自定義欄位
                irr_cash_flow = (income_values[i] - rent_values[i] - maintenance_values[i] -
                               insurance_values[i] - recycling_values[i] - irr_tax_amount)

                cash_flows.append(irr_cash_flow)

                # 記錄現金流項目（用於表格顯示）
                cash_flow_item = CashFlowItem(
                    year=years[i],
                    income=income_values[i],
                    equipment_depreciation=equipment_depreciation_per_year,
                    interest=interest_values[i],
                    rent=rent_values[i],
                    maintenance=maintenance_values[i],
                    insurance=insurance_values[i],
                    recycling=recycling_values[i],
                    custom=custom_values[i],
                    net_cash_flow=pretax_net_profit,  # 稅前淨利
                    tax_amount=display_tax_amount,     # 表格顯示用所得稅
                    after_tax_cash_flow=aftertax_net_profit  # 稅後淨利
                )
                cash_flow_items.append(cash_flow_item)

            # 5. 計算 IRR
            success, irr_value, error = IRRCalculatorService.calculate_irr_numpy(cash_flows)

            # 6. 計算現金流量表
            cash_flow_statement, irr_analysis = IRRCalculatorService.calculate_cash_flow_statement(
                cash_flow_items, equipment_cost, years, request.cash_flow_params, request.interest
            )

            # 7. 建立回應
            response = IRRCalculationResponse(
                success=success,
                irr=irr_value if success else None,
                error=error,
                equipment_cost=equipment_cost,
                cash_flows=cash_flow_items,
                cash_flow_statement=cash_flow_statement,
                irr_analysis=irr_analysis,
                years=years
            )

            return response

        except Exception as e:
            return IRRCalculationResponse(
                success=False,
                irr=None,
                error=f"計算過程發生錯誤: {str(e)}",
                equipment_cost=0,
                cash_flows=[],
                cash_flow_statement=[],
                irr_analysis=IRRAnalysis(),
                years=[]
            )