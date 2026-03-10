# tool/tool_manager.py
import json
from typing import Dict, List, Callable, Any, Optional
from tool.excel_tool import ExcelTool, EXCEL_TOOLS_SCHEMA
from tool.finance_tool import FinanceTool, FINANCE_TOOLS_SCHEMA
from tool.price_rolling_tool import PriceRollingTool, PRICE_ROLLING_TOOLS_SCHEMA
from tool.equipment_cost_tool import EquipmentCostTool, EQUIPMENT_COST_TOOLS_SCHEMA


class ToolManager:
    """工具管理器：負責註冊、管理和執行工具"""

    def __init__(self):
        """初始化工具管理器"""
        self.tools: Dict[str, Callable] = {}
        self.tools_schema: List[Dict] = []
        self._register_default_tools()

    def _register_default_tools(self):
        """註冊預設工具"""
        # 創建 Excel 工具實例
        self.excel_tool = ExcelTool()

        # 註冊 Excel 相關工具
        self.register_tool(
            name="write_excel_cell",
            function=self.excel_tool.write_cell,
            schema=EXCEL_TOOLS_SCHEMA[0]
        )
        self.register_tool(
            name="read_excel_cell",
            function=self.excel_tool.read_cell,
            schema=EXCEL_TOOLS_SCHEMA[1]
        )
        self.register_tool(
            name="delete_excel_cell",
            function=self.excel_tool.delete_cell,
            schema=EXCEL_TOOLS_SCHEMA[2]
        )
        self.register_tool(
            name="edit_sheet_by_field",
            function=self.excel_tool.edit_by_field_and_year,
            schema=EXCEL_TOOLS_SCHEMA[3]
        )
        self.register_tool(
            name="list_excel_sheets",
            function=self.excel_tool.list_sheets,
            schema=EXCEL_TOOLS_SCHEMA[4]
        )
        self.register_tool(
            name="read_sheet_by_field",
            function=self.excel_tool.read_sheet_by_field,
            schema=EXCEL_TOOLS_SCHEMA[5]
        )

        # 創建財務工具實例
        self.finance_tool = FinanceTool()

        # 註冊財務相關工具
        self.register_tool(
            name="calculate_irr",
            function=self.finance_tool.calculate_irr,
            schema=FINANCE_TOOLS_SCHEMA[0]
        )
        self.register_tool(
            name="get_cash_flow",
            function=self.finance_tool.get_cash_flow,
            schema=FINANCE_TOOLS_SCHEMA[1]
        )
        self.register_tool(
            name="get_net_profit",
            function=self.finance_tool.get_net_profit,
            schema=FINANCE_TOOLS_SCHEMA[2]
        )
        self.register_tool(
            name="get_year_detail",
            function=self.finance_tool.get_year_detail,
            schema=FINANCE_TOOLS_SCHEMA[3]
        )
        self.register_tool(
            name="get_project_summary",
            function=self.finance_tool.get_project_summary,
            schema=FINANCE_TOOLS_SCHEMA[4]
        )

        # 創建價金滾算工具實例
        self.price_rolling_tool = PriceRollingTool()

        # 註冊價金滾算工具（純計算，不寫入檔案）
        self.register_tool(
            name="calculate_price_rolling",
            function=self.price_rolling_tool.calculate_price_rolling,
            schema=PRICE_ROLLING_TOOLS_SCHEMA[0]
        )

        # 創建設備成本工具實例
        self.equipment_cost_tool = EquipmentCostTool()

        # 註冊設備成本滾算工具（完整流程：讀取→計算→寫入Excel）
        self.register_tool(
            name="execute_price_rolling",
            function=self.equipment_cost_tool.execute_price_rolling,
            schema=EQUIPMENT_COST_TOOLS_SCHEMA[0]
        )

        print(f"已註冊 {len(self.tools)} 個工具")

    def register_tool(self, name: str, function: Callable, schema: Dict):
        """
        註冊一個新工具

        Args:
            name: 工具名稱
            function: 工具函數
            schema: 工具的 JSON schema
        """
        self.tools[name] = function
        self.tools_schema.append(schema)

    def get_tools_schema(self) -> List[Dict]:
        """
        獲取所有工具的 schema（用於 function calling）

        Returns:
            工具 schema 列表
        """
        return self.tools_schema

    def execute_tool(self, tool_name: str, arguments: Dict[str, Any]) -> Dict:
        """
        執行指定的工具

        Args:
            tool_name: 工具名稱
            arguments: 工具參數

        Returns:
            執行結果
        """
        if tool_name not in self.tools:
            return {
                "success": False,
                "message": f"未知的工具: {tool_name}",
                "available_tools": list(self.tools.keys())
            }

        try:
            # 執行工具函數
            tool_function = self.tools[tool_name]
            result = tool_function(**arguments)
            return result
        except Exception as e:
            return {
                "success": False,
                "message": f"執行工具 {tool_name} 時發生錯誤: {str(e)}",
                "arguments": arguments
            }

    def list_tools(self) -> List[str]:
        """
        列出所有已註冊的工具

        Returns:
            工具名稱列表
        """
        return list(self.tools.keys())

    def get_tool_info(self, tool_name: str) -> Optional[Dict]:
        """
        獲取工具詳細資訊

        Args:
            tool_name: 工具名稱

        Returns:
            工具的 schema 資訊
        """
        for schema in self.tools_schema:
            if schema["function"]["name"] == tool_name:
                return schema
        return None

    def set_finance_excel_file(self, excel_file_path: str, sheet_name: str = None):
        """
        設定財務工具使用的 Excel 檔案路徑和工作表

        Args:
            excel_file_path: Excel 檔案的完整路徑
            sheet_name: 工作表名稱，如果為 None 則使用第一個工作表
        """
        if hasattr(self, 'finance_tool'):
            self.finance_tool.set_excel_file(excel_file_path, sheet_name)

        if hasattr(self, 'price_rolling_tool'):
            self.price_rolling_tool.set_excel_file(excel_file_path, sheet_name)

        # 同時設定 Excel 編輯工具的檔案路徑
        if hasattr(self, 'excel_tool'):
            self.excel_tool.file_path = excel_file_path

        if sheet_name:
            print(f"財務工具 Excel 檔案已設定: {excel_file_path}, 工作表: {sheet_name}")
        else:
            print(f"財務工具 Excel 檔案已設定: {excel_file_path}, 使用第一個工作表")
