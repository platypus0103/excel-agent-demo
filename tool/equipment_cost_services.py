"""
簡介：此程式為價金速算的藍圖，目的是為了將程式快速套入專案中，可視情況修改調整成符合設計的內容。
規格描述：
此功能的目的在於價金的快速滾算作業，
從初始情境的價金開始，不斷往下調(減少支出)，
先行假設價金為正數，若與設計輸入有差異，可以自行調整邏輯。

1. CostStructureService：將滾算完的價金(設備費用)，進行成本結構變動計算(信邦利潤率profit_rate、開發費用development_fee)
2. CashMod-現金模式：Step為設定現金的計算模式
3. RatioMode-比率模式：Step為設定比例的計算模式
4. ConditionalMode-條件模式：Step採取條件是設定的計算模式(可參考，未必會實作)
5. CustomizeMode-自訂模式：User可自行決定滾算次數、每一次的Step值(或採取隨機自動配置)
"""
import random
from typing import List

class CostStructureService:
    """
    規格描述：此功能的目的在於價金的成本結構更動，函數說明如下
    __init__()：類別建構子
    get_profit()：計算利率
    equipment_cost_calculation()：計算成本結構變更後的設備費用 / 價金(設備費用 = 原始價金 / (1 - 信邦利潤率) + 開發費)
    """
    def __init__(self, profit_rate:float, development_fee:int) -> None:
        """
        建構子參數：
        profit_rate：信邦利潤率
        development_fee：開發費用
        adjustment_record：調整過後的價金
        """
        self.profit_rate = profit_rate
        self.development_fee = development_fee

    def get_profit(self, adjustment_record:List[int]) -> List[int]:
        """
        規格描述：此功能用來滾算設備費用並記錄每次滾算的結果。
        adjustment_record：調整過後的價金
        profit_record：紀錄計算的信邦利潤
        profit：單筆計算信邦利潤紀錄
        """
        profit_record = []
        for equipment_cost in adjustment_record:
            profit = round(equipment_cost / (1 - self.profit_rate) - equipment_cost)
            profit_record.append(profit)
        return profit_record
    
    def equipment_cost_calculation(self, adjustment_record:List[int]) -> List[int]:
        """
        規格描述：計算成本結構變更後的設備費用
        調整公式: equipment_cost / (1 - profit_rate) + development_fee
        adjustment_record：滾算後的設備價金列表
        return：成本結構變更後的設備費用列表
        """
        cost_record = []
        for equipment_cost in adjustment_record:
            cost = round(equipment_cost / (1 - self.profit_rate) + self.development_fee)
            cost_record.append(cost)
        return cost_record
            


class CashMode:
    """
    規格描述：現金模式下的價金滾算功能，Step為現金，函數說明如下
    __init__()：類別建構子
    calculation()：滾算函數(公式：價金 - 調整步幅)
    """
    def __init__(self, boundary:int, step:int) -> None:
        """
        建構子參數：
        boundary：價金調整的邊界
        step：每一次的調整步幅
        """
        self.boundary = boundary
        self.step = step


    def calculation(self, equipment_cost:int) -> List[int]:
        """
        規格描述：此功能用來滾算設備費用並記錄每次滾算的結果。
        equipment_cost：初始價金(初始設備費用)
        adjustable_range：判斷是否還可繼續調整，若不行則停止調整
        adjustment_record：紀錄調整過後的價金變化
        """        
        adjustment_record = [equipment_cost]
        adjustable_range = equipment_cost - self.boundary
        
        while equipment_cost > self.boundary and adjustable_range >= self.step:
            equipment_cost -= self.step # 將價金減一個Step
            adjustable_range = equipment_cost - self.boundary # 確定本次調整後的剩餘可調整範圍
            adjustment_record.append(equipment_cost) # 將滾算後的價金記錄下來
        
        return adjustment_record


class RatioMode:
    """
    規格描述：比率模式下的價金滾算功能，Step為比率(如0.01，範圍0 < Step < 1)，函數說明如下
    __init__()：類別建構子
    calculation()：滾算函數(公式：價金 * (1 - 調整步幅))
    """
    def __init__(self, boundary:int, step:float) -> None:
        """
        建構子參數：
        boundary：價金調整的邊界
        step：每一次的調整步幅
        """
        self.boundary = boundary
        self.step = step

    def calculation(self, equipment_cost:int) -> List[int]:
        """
        規格描述：此功能用來滾算設備費用並記錄每次滾算的結果。
        equipment_cost：初始價金(初始設備費用)
        adjustment_record：紀錄調整過後的價金變化
        adjsuted_cost：調整過後的價金
        """
        adjustment_record = [round(equipment_cost)]
        while True:
            adjusted_cost = equipment_cost * (1 - self.step) # 計算調整過後的價金
            
            if adjusted_cost < self.boundary: # 判斷調整後的價金是否已經超出邊界
                break
            
            equipment_cost = adjusted_cost 
            adjustment_record.append(round(equipment_cost))

        return adjustment_record 


class ConditionalMode:
    """
    規格描述：條件模式下的價金滾算功能，Step為現金，但由User設定的最大最小值來決定，最多會有三種，函數說明如下
    __init__()：類別建構子
    calculation()：滾算函數(公式：價金 - 調整步幅)
    """
    def __init__(self, boundary:int, maximum_value:int, minimum_value:int, condition_step_1:int, condition_step_2:int, condition_step_3:int) -> None:
        """
        建構子參數：
        boundary：價金調整的邊界
        maximum_value：判斷價金範圍決定step的最大值
        minimum_value：判斷價金範圍決定step的最小值
        condition_step_1：價金 > maximum_value的step
        condition_step_2：minimum < 價金 < maximum_value的step
        condition_step_3：價金 < minimum_value的step
        """
        self.boundary = boundary
        self.maximum_value = maximum_value
        self.minimum_value = minimum_value
        self.condition_step_1 = condition_step_1
        self.condition_step_2 = condition_step_2
        self.condition_step_3 = condition_step_3

    def calculation(self, equipment_cost:int) -> List[int]:
        """
        規格描述：此功能用來滾算設備費用並記錄每次滾算的結果。
        equipment_cost：初始價金(初始設備費用)
        adjustable_range：判斷是否還可繼續調整，若不行則停止調整
        adjustment_record：紀錄調整過後的價金變化

        請注意!!minimum_value不可以小於boundary
        """   
        adjustment_record = [equipment_cost]

        while True:
            # Step設定判斷式
            if equipment_cost > self.maximum_value:
                step = self.condition_step_1 # 價金 > maximum_value的step
            elif self.minimum_value <= equipment_cost <= self.maximum_value:
                step = self.condition_step_2 # minimum < 價金 < maximum_value的step
            else:
                step = self.condition_step_3 # 價金 < minimum_value的step

            # 檢查是否達到邊界
            adjustable_range = equipment_cost - self.boundary
            if equipment_cost <= self.boundary or adjustable_range < step:
                break

            equipment_cost -= step # 將價金減一個Step
            adjustment_record.append(equipment_cost) # 將滾算後的價金記錄下來

        return adjustment_record
    

class CustomizeMode:
    """
    規格描述：自訂模式下的價金滾算功能，使用者可以自行決定滾算次數與每次滾算的step，或者user可以採取隨機的方式自動配置step，函數說明如下。
    __init__()：類別建構子
    automatic_configuration()：自動產生每一次滾算的step
    calculation()：滾算函數(公式：價金 - 調整步幅)
    """
    def __init__(self, boundary:int, adjustment_times:int=10) -> None:
        """
        建構子參數：
        boundary：價金調整的邊界
        adjustment_times：使用者決定的調整次數
        """
        self.boundary = boundary
        self.adjustment_times = adjustment_times

    def automatic_configuration(self, equipment_cost:int) -> List[int]:
        """
        規格描述：自動且合理的配置每一次的調整步幅，方便使用者快速設定。
        equipment_cost：初始價金(初始設備費用)
        adjustable_range：判斷是否還可繼續調整，若不行則停止調整
        adjustment_times：使用者決定的調整次數
        average_step：計算每一步的step，作為隨機生成的依據
        steps：存放每一部自動配置的step
        """
        steps = []
        adjustable_range = equipment_cost - self.boundary # 計算可調整範圍

        adjustment_times = min(self.adjustment_times, 10) # 取得User設定的滾算次數，最多預設為10次

        average_step = adjustable_range // adjustment_times # 計算每一步的step，作為隨機生成的依據

        for _ in range(adjustment_times):
            # 在平均值 ±20% 範圍內生成隨機 step
            step = random.randint(int(average_step * 0.8), int(average_step * 1.2))

            # 確保這一次的step不會超過可調整範圍
            step = min(step, adjustable_range)

            # 將每一次的step加入列表
            steps.append(step)

            # 扣對這一次產生的step，更新可調整範圍
            adjustable_range -= step

            # 若可調整範圍已經到達極限，停止生成
            if adjustable_range <= 0:
                break
        return steps
    
    def calculation(self, steps:List[int], equipment_cost:int) -> List[int]:
        """
        規格描述：此功能用來滾算設備費用並記錄每次滾算的結果。
        equipment_cost：初始價金(初始設備費用)
        adjustment_record：紀錄調整過後的價金變化
        adjsuted_cost：調整過後的價金
        steps：使用者設定的每一次step
        """

        adjustment_record = [equipment_cost]
        for step in steps:
            equipment_cost -= step
            adjustment_record.append(equipment_cost)

        return adjustment_record

        
            









    
            

    


