# tool/tool_manager.py
import json
from typing import Dict, List, Callable, Any, Optional
from tool.excel_tool import ExcelTool, EXCEL_TOOLS_SCHEMA


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
        self.register_tool(
            name="query_financial_data",
            function=self.excel_tool.query_financial_data,
            schema=EXCEL_TOOLS_SCHEMA[6]
        )
        self.register_tool(
            name="delete_excel_sheet",
            function=self.excel_tool.delete_sheet,
            schema=EXCEL_TOOLS_SCHEMA[7]
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

    def set_excel_file(self, excel_file_path: str, sheet_name: str = None):
        """設定 Excel 工具的檔案路徑"""
        if hasattr(self, 'excel_tool'):
            self.excel_tool.file_path = excel_file_path

        if sheet_name:
            print(f"Excel 檔案已設定: {excel_file_path}, 工作表: {sheet_name}")
        else:
            print(f"Excel 檔案已設定: {excel_file_path}")
