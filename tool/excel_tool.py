# tool/excel_tool.py
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from typing import Optional, Union, List

class ExcelTool:
    """Excel 操作工具類"""

    def __init__(self, file_path: str = "Excel/data.xlsx"):
        """
        初始化 Excel 工具

        Args:
            file_path: Excel 文件路徑（相對於專案根目錄）
        """
        self.file_path = file_path
        
        # 如果使用預設路徑且檔案不存在，嘗試搜尋現有檔案
        if self.file_path == "Excel/data.xlsx" and not os.path.exists(self.file_path):
            existing_file = self._find_existing_file()
            if existing_file:
                self.file_path = existing_file
                print(f"自動選用現有 Excel 檔案: {self.file_path}")

        self._ensure_file_exists()

    def _find_existing_file(self) -> Optional[str]:
        """搜尋 Excel 資料夾中現有的 Excel 檔案"""
        directory = os.path.dirname(self.file_path)
        if os.path.exists(directory):
            for file in os.listdir(directory):
                if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                    return os.path.join(directory, file)
        return None

    def _ensure_file_exists(self):
        """確保 Excel 文件存在"""
        # 確保目錄存在
        directory = os.path.dirname(self.file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)

        # 如果文件不存在，創建新文件
        if not os.path.exists(self.file_path):
            wb = Workbook()
            wb.save(self.file_path)
            print(f"創建新的 Excel 文件: {self.file_path}")

    def _parse_cell_range(self, cell_range: str) -> tuple:
        """
        解析儲存格範圍

        Args:
            cell_range: 儲存格範圍，如 "A1" 或 "A1:B5"

        Returns:
            (start_cell, end_cell) 如果是範圍，否則 (cell, None)
        """
        if ':' in cell_range:
            start, end = cell_range.split(':')
            return start.strip(), end.strip()
        return cell_range.strip(), None

    def write_cell(self, cell: str, value: Union[str, int, float]) -> dict:
        """
        寫入單個儲存格或範圍

        Args:
            cell: 儲存格位置，如 "A1" 或範圍 "A1:B2"
            value: 要寫入的值

        Returns:
            操作結果字典
        """
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active

            start_cell, end_cell = self._parse_cell_range(cell)

            if end_cell:
                # 範圍操作：將相同的值寫入所有儲存格
                for row in ws[start_cell:end_cell]:
                    for cell_obj in row:
                        cell_obj.value = value
                message = f"已將 {value} 寫入範圍 {cell}"
            else:
                # 單個儲存格操作
                ws[start_cell] = value
                message = f"已將 {value} 寫入儲存格 {cell}"

            wb.save(self.file_path)
            return {
                "success": True,
                "message": message,
                "cell": cell,
                "value": value
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"寫入失敗: {str(e)}",
                "cell": cell,
                "value": value
            }

    def read_cell(self, cell: str) -> dict:
        """
        讀取儲存格或範圍

        Args:
            cell: 儲存格位置，如 "A1" 或範圍 "A1:B2"

        Returns:
            操作結果字典
        """
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active

            start_cell, end_cell = self._parse_cell_range(cell)

            if end_cell:
                # 範圍操作：讀取所有儲存格的值
                values = []
                for row in ws[start_cell:end_cell]:
                    row_values = [cell_obj.value for cell_obj in row]
                    values.append(row_values)
                return {
                    "success": True,
                    "message": f"已讀取範圍 {cell}",
                    "cell": cell,
                    "value": values
                }
            else:
                # 單個儲存格操作
                value = ws[start_cell].value
                return {
                    "success": True,
                    "message": f"已讀取儲存格 {cell}",
                    "cell": cell,
                    "value": value
                }
        except Exception as e:
            return {
                "success": False,
                "message": f"讀取失敗: {str(e)}",
                "cell": cell
            }

    def delete_cell(self, cell: str) -> dict:
        """
        清空儲存格或範圍的內容

        Args:
            cell: 儲存格位置，如 "A1" 或範圍 "A1:B2"

        Returns:
            操作結果字典
        """
        try:
            wb = load_workbook(self.file_path)
            ws = wb.active

            start_cell, end_cell = self._parse_cell_range(cell)

            if end_cell:
                # 範圍操作：清空所有儲存格
                for row in ws[start_cell:end_cell]:
                    for cell_obj in row:
                        cell_obj.value = None
                message = f"已清空範圍 {cell}"
            else:
                # 單個儲存格操作
                ws[start_cell].value = None
                message = f"已清空儲存格 {cell}"

            wb.save(self.file_path)
            return {
                "success": True,
                "message": message,
                "cell": cell
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"刪除失敗: {str(e)}",
                "cell": cell
            }


# 定義工具的 schema（用於 function calling）
EXCEL_TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "write_excel_cell",
            "description": "在 Excel 文件中寫入或修改儲存格的值。可以寫入單個儲存格（如 A1）或範圍（如 A1:B5）",
            "parameters": {
                "type": "object",
                "properties": {
                    "cell": {
                        "type": "string",
                        "description": "儲存格位置，如 'A1' 或範圍 'A1:B5'"
                    },
                    "value": {
                        "type": ["string", "number"],
                        "description": "要寫入的值，可以是文字或數字"
                    }
                },
                "required": ["cell", "value"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_excel_cell",
            "description": "從 Excel 文件中讀取儲存格的值。可以讀取單個儲存格（如 A1）或範圍（如 A1:B5）",
            "parameters": {
                "type": "object",
                "properties": {
                    "cell": {
                        "type": "string",
                        "description": "儲存格位置，如 'A1' 或範圍 'A1:B5'"
                    }
                },
                "required": ["cell"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "delete_excel_cell",
            "description": "清空 Excel 文件中儲存格的內容。可以清空單個儲存格（如 A1）或範圍（如 A1:B5）",
            "parameters": {
                "type": "object",
                "properties": {
                    "cell": {
                        "type": "string",
                        "description": "儲存格位置，如 'A1' 或範圍 'A1:B5'"
                    }
                },
                "required": ["cell"]
            }
        }
    }
]
