# main.py (放在專案根目錄)

from core.agent import AIAgent
from config.settings import DEFAULT_CONFIG
from datetime import datetime

def main():
    """主函數"""
    # 創建 Agent（使用預設配置）
    agent = AIAgent(DEFAULT_CONFIG)
    
    
    while True:
       
        user_input = input("你: ").strip()
        
        if not user_input:
            continue
        
        # 處理命令
        if user_input.lower() == 'quit':
            print("再見！")
            break
        
        if user_input.lower() == 'history':
            agent.show_conversation()
            continue

        
        if user_input.lower() == 'reset':
            agent.reset()
            continue

        # AI 回應
        print("AI: ", end="", flush=True)
        response = agent.chat(user_input)
        print(response + "\n")


if __name__ == "__main__":
    main()