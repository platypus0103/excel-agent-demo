from typing import Optional
from core.connection import OllamaConnection
from core.conversation import ConversationManager
from config.settings import AgentConfig, DEFAULT_CONFIG
from utils.error_handler import AgentError
from tool.tool_manager import ToolManager
import json

class AIAgent:
    def __init__(self, config: Optional[AgentConfig] = None):
        self.config = config or DEFAULT_CONFIG
        self.connection = OllamaConnection(self.config.model_name)
        print(f"   模型: {self.config.model_name}")
        print(f"   Temperature: {self.config.temperature}")
        print(f"   Top-P: {self.config.top_p}")
        print(f"   最大歷史: {self.config.max_history_length} 輪")
        self.conversation = ConversationManager(
            system_prompt=self.config.system_prompt,
            max_history_length=self.config.max_history_length
        )
        # 初始化工具管理器
        self.tool_manager = ToolManager()
        # 追蹤最近使用的工具
        self.last_used_tools = []
    
    def chat(self, user_input: str) -> str:
        try:
            # 重置工具使用追蹤
            self.last_used_tools = []

            # 1. 將使用者訊息加入歷史
            self.conversation.add_user_message(user_input)

            # 2. 獲取完整對話歷史
            messages = self.conversation.get_messages()

            # 強制注入系統提醒（防止長對話導致 System Prompt 被截斷）
            if messages and messages[-1]['role'] == 'user':
                reminder = "\n\n[System Reminder]\n1. Respond in Traditional Chinese (繁體中文).\n2. Use tool data strictly. Do not hallucinate.\n3. If this is a rolling calculation, show the Base IRR first."
                messages[-1]['content'] += reminder

            # 3. 獲取工具 schema
            tools = self.tool_manager.get_tools_schema()

            # 4. 發送給 AI（包含工具資訊）
            response = self.connection.send_message_with_tools(
                messages=messages,
                tools=tools,
                temperature=self.config.temperature,
                top_p=self.config.top_p,
                top_k=self.config.top_k,
                think=self.config.thinking_mode
            )

            # 5. 處理回應（可能包含工具調用）
            final_response = self._handle_response(response, messages)

            # 6. 將最終 AI 回應加入歷史
            self.conversation.add_assistant_message(final_response)

            return final_response
        except AgentError as e:
            error_msg = f"對話過程中發生錯誤: {e}"
            print(error_msg)
            return error_msg

    def _strip_thinking(self, content: str) -> str:
        """移除 qwen3 的思考過程標籤 <think>...</think>"""
        import re
        content = re.sub(r'<think>.*?</think>', '', content, flags=re.DOTALL)
        return content.strip()

    def _handle_response(self, response: dict, messages: list) -> str:
        """
        處理 AI 回應，包括工具調用

        Args:
            response: AI 的原始回應
            messages: 當前的對話歷史

        Returns:
            最終的文字回應
        """
        message = response.get('message', {})

        # 檢查是否有工具調用
        tool_calls = message.get('tool_calls')

        if not tool_calls:
            # 沒有工具調用，直接返回文字回應（移除思考過程標籤）
            return self._strip_thinking(message.get('content', ''))

        # 有工具調用，執行工具
        print("\n🔧 AI 正在使用工具...")

        for tool_call in tool_calls:
            function_name = tool_call['function']['name']
            arguments = tool_call['function']['arguments']

            print(f"   調用工具: {function_name}")
            print(f"   參數: {json.dumps(arguments, ensure_ascii=False)}")

            # 執行工具
            tool_result = self.tool_manager.execute_tool(function_name, arguments)
            print(f"   結果: {tool_result.get('message', '執行完成')}\n")

            # 追蹤使用的工具
            self.last_used_tools.append(function_name)

            # 將工具調用和結果加入對話歷史
            messages.append({
                'role': 'assistant',
                'content': '',
                'tool_calls': [tool_call]
            })

            messages.append({
                'role': 'tool',
                'content': json.dumps(tool_result, ensure_ascii=False)
            })

        # 將工具結果發送給 AI，獲取最終回應
        tools = self.tool_manager.get_tools_schema()
        final_response = self.connection.send_message_with_tools(
            messages=messages,
            tools=tools,
            temperature=self.config.temperature,
            top_p=self.config.top_p,
            top_k=self.config.top_k,
            think=self.config.thinking_mode
        )

        return self._strip_thinking(final_response.get('message', {}).get('content', ''))
    
    def show_conversation(self):
        self.conversation.print_conversation()
    def reset(self,keep_system:bool=True):
        self.conversation.clear_conversation(keep_system)
        print("對話已重置")

if __name__=="__main__":
    test_config = AgentConfig(
        model_name="qwen3:4b",
        system_prompt="你是小智，一個友善的 AI 助手，回答簡潔明瞭，且全程使用繁體中文。"
    )
    agent = AIAgent(test_config)

    while True:
        user_input=input("你: ").strip()
        if not user_input:
            continue
        if user_input.lower() in ['exit','quit']:
            print("結束對話")
            break
        elif user_input.lower()=='reset':
            agent.reset()
            continue
        response=agent.chat(user_input)
        print(f"AI: {response}")