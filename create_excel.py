from openpyxl import Workbook
import os

# 創建 Excel 資料夾
excel_dir = "Excel"
if not os.path.exists(excel_dir):
    os.makedirs(excel_dir)
    print(f"Created {excel_dir} folder")

# 創建預設 Excel 文件
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# 儲存文件
file_path = os.path.join(excel_dir, "data.xlsx")
wb.save(file_path)
print(f"Created {file_path}")
