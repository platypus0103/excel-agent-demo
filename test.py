import numpy as np
import openpyxl
import numpy_financial as npf
from decimal import Decimal, ROUND_HALF_UP

wb = openpyxl.load_workbook('LLM輸入公版(1).xlsx', data_only=True)
s1=wb['輸入公版(輸入模擬)']
project_name=s1['B2'].value # 專案名稱
year_start=int(s1['B5'].value) if s1['B5'].value else None # 起始年份
year_end=int(s1['D5'].value) if s1['D5'].value else None # 結束年份
capacities=Decimal(str(s1['B4'].value)) if s1['B4'].value else None # 建置量
# 收入相關
electricity_generation=Decimal(str(s1['C10'].value)) if s1['C10'].value else None # 年發電量
first_year_decade=Decimal(str(s1['C11'].value)) if s1['C11'].value else None # 首年衰退率
every_year_decade=Decimal(str(s1['C12'].value)) if s1['C12'].value else None # 次年衰退率
rate_of_sell=Decimal(str(s1['C13'].value)) if s1['C13'].value else None # 躉售費率
income_years=int(s1['D10'].value) if s1['D10'].value else None # 收入年限
# 支出
equipment_cost=Decimal(str(s1['C16'].value)) if s1['C16'].value else None # 設備成本
equipment_years=int(s1['D16'].value) if s1['D16'].value else None # 設備年限
profit_rate=Decimal(str(s1['C17'].value)) if s1['C17'].value else None # 利潤率
development_cost=Decimal(str(s1['C18'].value)) if s1['C18'].value else None # 開發費
rent=Decimal(str(s1['C20'].value)) if s1['C20'].value else Decimal('0') # 租金
rent_years=int(s1['D20'].value) if s1['D20'].value else 0 # 租金年限
maintenance_cost=Decimal(str(s1['C22'].value)) if s1['C22'].value else Decimal('0') # 運維費
maintenance_years=int(s1['D22'].value) if s1['D22'].value else 0 # 運維年限
insurance_cost=Decimal(str(s1['C23'].value)) if s1['C23'].value else Decimal('0') # 保險費
insurance_years=int(s1['D23'].value) if s1['D23'].value else 0 # 保險年限
recycle_cost=Decimal(str(s1['C24'].value)) if s1['C24'].value else Decimal('0') # 回收費
recycle_years=int(s1['D24'].value) if s1['D24'].value else 0 # 回收年限
interest_rate=Decimal(str(s1['C25'].value)) if s1['C25'].value else None # 銀行利率
inconme_tax=Decimal(str(s1['C26'].value)) if s1['C26'].value else None # 所得稅
# 現金流量表
loan_rate=Decimal(str(s1['C31'].value)) if s1['C31'].value else None # 貸款比例
loan_amoritization_years=int(s1['C32'].value) if s1['C32'].value else None # 貸款攤還年限
dividend_rate=Decimal(str(s1['C33'].value)) if s1['C33'].value else None # 股利
reduction_amoritization_years=int(s1['C34'].value) if s1['C34'].value else None # 減資攤還年限

# Debug: 顯示讀取的原始資料
print("=" * 60)
print("原始資料檢查:")
print("=" * 60)
print(f"專案名稱: {project_name}")
print(f"年份: {year_start} - {year_end}")
print(f"容量: {capacities}")
print(f"年發電量: {electricity_generation}")
print(f"躉售費率: {rate_of_sell}")
print(f"收入年限: {income_years}")
print(f"設備成本: {equipment_cost}")
print(f"利潤率: {profit_rate}")
print(f"開發成本: {development_cost}")
print(f"租金: {rent}, 年限: {rent_years}")
print(f"運維: {maintenance_cost}, 年限: {maintenance_years}")
print(f"保險: {insurance_cost}, 年限: {insurance_years}")
print(f"回收: {recycle_cost}, 年限: {recycle_years}")
print(f"貸款比例: {loan_rate}")
print(f"貸款攤還年限: {loan_amoritization_years}")
print(f"銀行利率: {interest_rate}")
print("=" * 60)

# 檢查是否有 None 值
if None in [year_start, year_end, capacities, electricity_generation, rate_of_sell, 
            income_years, equipment_cost, profit_rate, development_cost]:
    print(" 錯誤：有關鍵資料為空值 (None)")
    exit(1)

total_equipment_cost = ((equipment_cost/(Decimal('1')-profit_rate))+development_cost)*capacities # 總設備成本
year_equipment_cost = total_equipment_cost / Decimal(str(equipment_years)) # 每年設備成本
# 計算專案總年限
total_years = int(year_end - year_start + 1)

# 計算每年的收入和支出（平均分攤）
# 總發電收入
first_year_electricity_generation= electricity_generation * (Decimal('1')-first_year_decade)
first_year_electricity_generation_income = first_year_electricity_generation * rate_of_sell # 首年發電收入
temp = first_year_electricity_generation # 暫存變數，用於計算後續年份發電量
electricity_generation_income_list=[]
electricity_generation_income_list.append(first_year_electricity_generation_income)

# 每年的支出（每個項目 ÷ 年限）
annual_rent = rent/Decimal(str(rent_years)) if rent_years and rent_years > 0 else Decimal('0') # 每年租金
annual_maintenance = maintenance_cost*capacities if maintenance_years and maintenance_years > 0 else Decimal('0') # 每年運維費
annual_insurance = equipment_cost*capacities*insurance_cost if insurance_years and insurance_years > 0 else Decimal('0') # 每年保險費
annual_recycle = recycle_cost*capacities/Decimal(str(recycle_years)) if recycle_years and recycle_years > 0 else Decimal('0') # 每年回收費

# 創建支出 list（長度 = total_years）
rent_list = []
maintenance_list = []
insurance_list = []
recycle_list = []
depreciation_list = []

for year in range(1, total_years + 1):
    # 租金（首年為0，次年起到 rent_years 為止）
    if year == 1:
        rent_list.append(Decimal('0'))
    elif year <= rent_years:
        rent_list.append(annual_rent)
    else:
        rent_list.append(Decimal('0'))

    # 運維費（根據年限）
    if maintenance_years and year <= maintenance_years:
        maintenance_list.append(annual_maintenance)
    else:
        maintenance_list.append(Decimal('0'))

    # 保險費（根據年限）
    if insurance_years and year <= insurance_years:
        insurance_list.append(annual_insurance)
    else:
        insurance_list.append(Decimal('0'))

    # 回收費（根據年限）
    if recycle_years and year <= recycle_years:
        recycle_list.append(annual_recycle)
    else:
        recycle_list.append(Decimal('0'))

    # 折舊（根據設備年限）
    if equipment_years and year <= equipment_years:
        depreciation_list.append(year_equipment_cost)
    else:
        depreciation_list.append(Decimal('0'))

# 創建利息 list 和完整的發電收入 list
interest_list = []
remaining_loan = loan_amount = total_equipment_cost * loan_rate # 貸款金額
pay_back = loan_amount / Decimal(str(loan_amoritization_years)) if loan_amoritization_years and loan_amoritization_years > 0 else Decimal('0') # 每年還款金額

for year in range(1, total_years + 1):
    # 計算該年利息
    interest = remaining_loan * interest_rate if remaining_loan > Decimal('0') else Decimal('0')
    interest_list.append(interest)

    # 更新剩餘貸款
    remaining_loan -= pay_back
    if remaining_loan < Decimal('0'):
        remaining_loan = Decimal('0')

    # 計算後續年份的發電收入（首年已經在前面加入了）
    if year > 1:
        temp = temp - (electricity_generation * every_year_decade)
        electricity_generation_income_list.append(temp * rate_of_sell)


tax_for_irr=[]
remaining_loan = loan_amount  # 重置用於後續計算
remaining_loan_list = loan_amount
first_year_interest = interest_list[0]  # 首年利息
#first_cost = annual_maintenance + annual_insurance + annual_recycle + first_year_interest # 首年總支出（不含租金）
total_cost = annual_rent + annual_maintenance + annual_insurance + annual_recycle # 次年起的基本支出（含租金）

# 初始化現金流陣列 (第0年是初始投資)
first_year_cash_flow = first_year_electricity_generation_income - total_cost # 首年現金流
first_tax = (first_year_cash_flow - year_equipment_cost + annual_rent) * inconme_tax # 首年所得稅
tax_for_irr.append(first_tax)
first_year_cash_flow -= first_tax  # 扣除所得稅

cash_flows = [-total_equipment_cost]  # 初始投資為負值
cash_flows.append(first_year_cash_flow)

# 更新剩餘貸款
remaining_loan -= pay_back

# 計算後續年份
for year in range(2, total_years + 1):
    # 使用 list 中的值
    following_years_electricity_generation_income = electricity_generation_income_list[year-1]
    #interest_payment = interest_list[year-1]

    annual_cash_flow = following_years_electricity_generation_income - rent_list[year-1] - maintenance_list[year-1] - insurance_list[year-1] - recycle_list[year-1]  # 次年及後續年份現金流
    annual_tax = (annual_cash_flow - year_equipment_cost) * inconme_tax # 次年及後續年份所得稅
    annual_cash_flow -= annual_tax # 扣除所得稅
    tax_for_irr.append(annual_tax)
    cash_flows.append(annual_cash_flow)

    # 更新剩餘貸款（維持原有邏輯）
    remaining_loan -= pay_back
    if remaining_loan < 0:
        remaining_loan = 0
net_profit_after_tax=[]
for year in range(1, total_years + 1):
    # 使用 list 中的值
    incom = electricity_generation_income_list[year-1]
    rent_expense = rent_list[year-1]
    maintenance_expense = maintenance_list[year-1]
    insurance_expense = insurance_list[year-1]
    recycle_expense = recycle_list[year-1]
    interest_expense = interest_list[year-1]
    depreciation_expense = depreciation_list[year-1]

    if year == 1:
        # 首年總支出（不含租金）
        net_profit_after_tax_year = incom - maintenance_expense - insurance_expense - recycle_expense - interest_expense - depreciation_expense

    else:
        # 次年起（含租金）
        net_profit_after_tax_year = incom - rent_expense - maintenance_expense - insurance_expense - recycle_expense - interest_expense - depreciation_expense

    tax_flow = net_profit_after_tax_year * inconme_tax
    net_profit_after_tax_year -= tax_flow
    net_profit_after_tax.append(net_profit_after_tax_year)


# 計算IRR（需要轉換為 float 以便 numpy 計算）
irr = npf.irr([float(cf) for cf in cash_flows])
#成本法IRR
cost_method_initial_investment = (total_equipment_cost*(Decimal('1') - loan_rate))
cost_method_cash_flows =[]
cost_method_cash_flows.append(-cost_method_initial_investment)
reduction_at_the_end=(cost_method_initial_investment/Decimal(str(reduction_amoritization_years)))
dividend_list=[]
for year in range(1, total_years ):
    dividend=net_profit_after_tax[year-1]*dividend_rate
    cost_method_cash_flows.append(dividend)
for year in range(total_years - reduction_amoritization_years, total_years):
    cost_method_cash_flows[year]+=reduction_at_the_end
    
last_cost_method_cash_flows=sum(net_profit_after_tax)-(sum(net_profit_after_tax)-net_profit_after_tax[-1])*dividend_rate
cost_method_cash_flows.append(last_cost_method_cash_flows)
cost_method_irr = npf.irr([float(cf) for cf in cost_method_cash_flows])

#權益法IRR
equity_method_initial_investment = -(total_equipment_cost * (Decimal('1') - loan_rate))+net_profit_after_tax[0]
equity_method_cash_flows =[]
equity_method_cash_flows.append(equity_method_initial_investment)
for year in range(2, total_years + 1):
    equity_method_cash_flows.append(net_profit_after_tax[year-1])
    if (year-1)>=(total_years- reduction_amoritization_years):
        equity_method_cash_flows[year-1]+=reduction_at_the_end
equity_method_irr = npf.irr([float(cf) for cf in equity_method_cash_flows])

# 四捨五入函數
def round_decimal(value):
    """將 Decimal 四捨五入到整數"""
    if isinstance(value, Decimal):
        return int(value.quantize(Decimal('1'), rounding=ROUND_HALF_UP))
    return int(value)

def round_irr(value):
    """將 IRR 四捨五入到小數點後兩位"""
    if np.isnan(value):
        return value
    return round(value * 100, 2)

# 輸出結果
print(f"\n專案名稱: {project_name}")
print(f"專案期間: {int(year_start)} - {int(year_end)} ({total_years}年)")
print(f"初始投資: ${round_decimal(total_equipment_cost):,}")
print(f"貸款金額: ${round_decimal(loan_amount):,}")
print(f"每年還款: ${round_decimal(pay_back):,}")
print(f"\n收入計算:")
print(f"首年發電收入: ${round_decimal(first_year_electricity_generation_income):,}")
print(f"\n支出計算:")
print(f"每年租金: ${round_decimal(annual_rent):,} (共{int(rent_years) if rent_years else 0}年)")
print(f"每年運維: ${round_decimal(annual_maintenance):,} (共{int(maintenance_years) if maintenance_years else 0}年)")
print(f"每年保險: ${round_decimal(annual_insurance):,} (共{int(insurance_years) if insurance_years else 0}年)")
print(f"每年回收: ${round_decimal(annual_recycle):,} (共{int(recycle_years) if recycle_years else 0}年)")
print(f"首年利息: ${round_decimal(first_year_interest):,}")
print(f"\n年度現金流:")
print(f"第0年 (初始投資): ${round_decimal(cash_flows[0]):,}")
for i in range(1, min(6, len(cash_flows))):  # 只顯示前5年
    print(f"第{i}年: ${round_decimal(cash_flows[i]):,}")
if len(cash_flows) > 6:
    print(f"... (省略中間年份)")
    for i in range(len(cash_flows)-2, len(cash_flows)):
        print(f"第{i}年: ${round_decimal(cash_flows[i]):,}")

# 檢查現金流是否合理
positive_flows = [cf for cf in cash_flows if cf > Decimal('0')]
negative_flows = [cf for cf in cash_flows if cf < Decimal('0')]
print(f"\n現金流統計:")
print(f"正現金流數量: {len(positive_flows)}")
print(f"負現金流數量: {len(negative_flows)}")
print(f"現金流總和: ${round_decimal(sum(cash_flows)):,}")

if len(positive_flows) == 0:
    print("\n沒有正現金流，專案無法回本")
elif len(negative_flows) == 0:
    print("\n沒有負現金流，資料可能有誤")

print(f"\n內部報酬率 (IRR): {round_irr(irr):.2f}%" if not np.isnan(irr) else "\n內部報酬率 (IRR): 無法計算 (NaN)")
print(f"成本法內部報酬率 (Cost Method IRR): {round_irr(cost_method_irr):.2f}%" if not np.isnan(cost_method_irr) else "成本法內部報酬率 (Cost Method IRR): 無法計算 (NaN)")
print(f"權益法內部報酬率 (Equity Method IRR): {round_irr(equity_method_irr):.2f}%" if not np.isnan(equity_method_irr) else "權益法內部報酬率 (Equity Method IRR): 無法計算 (NaN)")
# 如果 IRR 是 NaN，提供更多資訊
if np.isnan(irr):
    print("\nIRR 無法計算的可能原因：")
    print("   1. 現金流全部為負值（專案無法回本）")
    print("   2. 現金流全部為正值（沒有初始投資）")
    print("   3. 現金流模式導致無解（多個 IRR 或無 IRR）")
    print("   4. 請檢查 Excel 資料是否正確")
print("\n" + "=" * 60)
print("各年度詳細數據:")
print("=" * 60)
for year in range(total_years):
    print(f"\n第{year+1}年:")
    print(f"  發電收入: ${round_decimal(electricity_generation_income_list[year]):,}")
    print(f"  租金: ${round_decimal(rent_list[year]):,}")
    print(f"  運維費: ${round_decimal(maintenance_list[year]):,}")
    print(f"  保險費: ${round_decimal(insurance_list[year]):,}")
    print(f"  回收費: ${round_decimal(recycle_list[year]):,}")
    print(f"  利息: ${round_decimal(interest_list[year]):,}")
    print(f"  折舊: ${round_decimal(depreciation_list[year]):,}")
    print(f"  現金流: ${round_decimal(cash_flows[year+1]):,}")
    print(f"  稅後淨利: ${round_decimal(net_profit_after_tax[year]):,}")
    print(f"  成本法現金流: ${round_decimal(cost_method_cash_flows[year+1]):,}")