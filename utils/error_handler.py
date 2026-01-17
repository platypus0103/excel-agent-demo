# utils/error_handler.py
import time
from functools import wraps
from typing import Callable, Any

class AgentError(Exception):
    """Agent 自定義錯誤"""
    pass

class ConnectionError(AgentError):
    """連接錯誤"""
    pass

class ModelError(AgentError):
    """模型錯誤"""
    pass

class TimeoutError(AgentError):
    """超時錯誤"""
    pass


def retry_on_error(max_retries: int = 3, delay: float = 1.0):
    """
    重試裝飾器
    
    解釋：
    想像你打電話給朋友，如果沒接通：
    1. 等一下再打（delay）
    2. 最多打3次（max_retries）
    3. 還是不通就放棄，告訴你「打不通」
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            last_exception = None
            
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    if attempt < max_retries - 1:
                        print(f"嘗試 {attempt + 1}/{max_retries} 失敗: {str(e)}")
                        print(f"等待 {delay} 秒後重試...")
                        time.sleep(delay)
                    else:
                        print(f"已達最大重試次數")
            
            raise last_exception
        
        return wrapper
    return decorator


def safe_execute(func: Callable, default_return: Any = None) -> Any:
    """
    安全執行函數
    
    用費曼技巧解釋：
    就像給函數加上「安全氣囊」
    出錯時不會讓整個程式崩潰，而是回傳預設值
    """
    try:
        return func()
    except Exception as e:
        print(f"執行錯誤: {str(e)}")
        return default_return


# 測試程式
if __name__ == "__main__":
    @retry_on_error(max_retries=3, delay=1.0)
    def unstable_function():
        """模擬不穩定的函數"""
        import random
        if random.random() < 0.7:  # 70% 機率失敗
            raise Exception("隨機錯誤")
        return "成功！"
    
    print("測試重試機制...")
    try:
        result = unstable_function()
        print(f"{result}")
    except Exception as e:
        print(f"最終失敗: {e}")