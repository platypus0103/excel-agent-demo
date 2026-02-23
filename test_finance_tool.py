# test_finance_tool.py
"""測試財務工具的功能"""
from tool.tool_manager import ToolManager
import json


def test_finance_tools():
    """測試財務工具"""
    print("=" * 60)
    print("財務工具測試")
    print("=" * 60)

    # 初始化工具管理器
    tool_manager = ToolManager()

    # 列出所有可用工具
    print("\n可用工具:")
    for tool_name in tool_manager.list_tools():
        print(f"  - {tool_name}")

    # 測試 1: 計算 IRR
    print("\n" + "=" * 60)
    print("測試 1: 計算 IRR")
    print("=" * 60)
    result = tool_manager.execute_tool("calculate_irr", {})
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # 測試 2: 獲取專案摘要
    print("\n" + "=" * 60)
    print("測試 2: 獲取專案摘要")
    print("=" * 60)
    result = tool_manager.execute_tool("get_project_summary", {})
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # 測試 3: 獲取所有年份現金流
    print("\n" + "=" * 60)
    print("測試 3: 獲取所有年份現金流")
    print("=" * 60)
    result = tool_manager.execute_tool("get_cash_flow", {})
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # 測試 4: 獲取第 5 年現金流
    print("\n" + "=" * 60)
    print("測試 4: 獲取第 5 年現金流")
    print("=" * 60)
    result = tool_manager.execute_tool("get_cash_flow", {"year": 5})
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # 測試 5: 獲取第 5 年詳細資料
    print("\n" + "=" * 60)
    print("測試 5: 獲取第 5 年詳細資料")
    print("=" * 60)
    result = tool_manager.execute_tool("get_year_detail", {"year": 5})
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # 測試 6: 獲取所有年份稅後淨利
    print("\n" + "=" * 60)
    print("測試 6: 獲取所有年份稅後淨利")
    print("=" * 60)
    result = tool_manager.execute_tool("get_net_profit", {})
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    test_finance_tools()
