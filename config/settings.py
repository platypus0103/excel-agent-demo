# config/settings.py
from dataclasses import dataclass
from typing import Optional

@dataclass
class AgentConfig:

    
    # 模型設定
    model_name: str = "qwen3:4b"
    
    # 生成參數
    temperature: float = 0.1        # 創意度 (0.0-2.0)
    top_p: float = 0.1            # 多樣性 (0.0-1.0)
    top_k: int = 100                # 候選詞數量
    
    # 對話管理
    max_history_length: int = 20   # 最多保留幾輪對話
    context_window: int = 4096     # 上下文窗口大小（token數）
    
    # 系統設定
    system_prompt: str = (
        "你是優秀的財務 AI 助手，協助分析太陽能電站財務數據。所有回應使用繁體中文。\n\n"
        "【核心原則】\n"
        "- 禁止自行編造數值，必須依賴工具回傳的數據。\n"
        "- 若輸入不清楚、過短或無法判斷意圖，直接詢問使用者，禁止猜測。\n"
        "- 若提到不存在的工作表，告知使用者可用的工作表名稱，不要猜測或編造。\n\n"
        "【工具使用】\n"
        "- IRR 計算 → calculate_irr\n"
        "- 價金滾算（cashmode/ratiomode/customizemode 等關鍵字）→ calculate_price_rolling\n"
        "  若指令含明確數值則傳入，否則留空讓工具從 Excel 讀取。\n"
        "- 修改工作表 → edit_sheet_by_field\n"
        "  說「支出/費用/成本」→ is_expense=true；說「收入」→ is_expense=false\n"
        "  說「公版」→ section_type=公版；說「綜合損益表」→ section_type=綜合損益表；說「現金流量表」→ section_type=現金流量表\n\n"
        "【回應格式】\n"
        "- 滾算結果用 Markdown 表格呈現，先列出原始 IRR 作為基準。\n"
        "- 修改後回報：工作表、欄位、年份範圍、新數值。"
    )
    
    # 錯誤處理
    retry_times: int = 3           # 失敗重試次數
    timeout: int = 30              # 超時時間（秒）
    
    def __post_init__(self):
        """驗證參數合法性"""
        if not 0.0 <= self.temperature <= 2.0:
            raise ValueError("temperature 必須在 0.0-2.0 之間")
        if not 0.0 <= self.top_p <= 1.0:
            raise ValueError("top_p 必須在 0.0-1.0 之間")
        if self.max_history_length < 1:
            raise ValueError("max_history_length 必須 >= 1")


# 預設配置
DEFAULT_CONFIG = AgentConfig()