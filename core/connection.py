import ollama
from typing import Dict, List, Optional
from utils.error_handler import retry_on_error, ConnectionError

class OllamaConnection:
    def __init__(self, model_name: str = "qwen3:4b"):
        self.model_name = model_name
        self._test_connection()

    @retry_on_error(max_retries=3, delay=1.0)
    def _test_connection(self):
        try:
            ollama.list()
            print(f"使用模型:{self.model_name}")
        except Exception as e:
            print(f"連接失敗{e}")
            raise

    @retry_on_error(max_retries=2, delay=0.5)
    def send_message(self, messages: List[Dict], temperature: float = 0.7, top_p: float = 0.9, top_k: int = 40) -> str:
        try:
            response = ollama.chat(
                model=self.model_name,
                messages=messages,
                options={
                    'temperature': temperature,
                    'top_p': top_p,
                    'top_k': top_k
                }
            )
            return response['message']['content']
        except Exception as e:
            raise ConnectionError(f"發送訊息失敗: {e}")

    @retry_on_error(max_retries=2, delay=0.5)
    def send_message_with_tools(
        self,
        messages: List[Dict],
        tools: List[Dict],
        temperature: float = 0.7,
        top_p: float = 0.9,
        top_k: int = 40,
        think: bool = False
    ) -> Dict:
        """
        發送訊息並支援工具調用（function calling）

        Args:
            messages: 對話歷史
            tools: 工具 schema 列表
            temperature: 創意度
            top_p: 多樣性
            top_k: 候選詞數量
            think: 是否啟用思考模式（qwen3 支援，32B 建議開啟）

        Returns:
            完整的回應字典（包含可能的工具調用）
        """
        try:
            response = ollama.chat(
                model=self.model_name,
                messages=messages,
                tools=tools,
                think=think,
                options={
                    'temperature': temperature,
                    'top_p': top_p,
                    'top_k': top_k
                }
            )
            return response
        except Exception as e:
            raise ConnectionError(f"發送訊息失敗: {e}")
    


