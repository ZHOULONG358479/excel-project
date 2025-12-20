from openpyxl import load_workbook
from datetime import datetime
import os
import glob
import re
desktop_path = "/Users/zhouqingquan/Desktop"
file_prefix = "第二作业项目部在用承包商员工数统计表"
files = glob.glob(os.path.join(desktop_path, f"{file_prefix}*.xlsx"))
if not files:
    print("未找到符合条件的 Excel 文件")
    exit()
# 4. 取第一个匹配文件（一般桌面只会有一个）
old_path = files[0]
dir_name = os.path.dirname(old_path)
old_name = os.path.basename(old_path)
# 5. 获取今天日期
today = datetime.now().strftime("%Y年%m月%d日")
# 6. 正则：匹配文件名末尾的8位日期
new_name = re.sub(r"\d{8}(?=\.xlsx$)", today, old_name)
# 7. 如果文件名中没有日期，就追加
if new_name == old_name:
    new_name = f"{file_prefix}_{today}.xlsx"
new_path = os.path.join(dir_name, new_name)
# 8. 重命名
os.rename(old_path, new_path)
print(f"文件已重命名为：{new_name}")
excel_path = files[0]
# 9. 打开 Excel
wb = load_workbook(excel_path)
# 10. 找到最后一个工作表
last_sheet = wb.worksheets[-1]
# 11. 复制工作表（自动放在最后）
new_sheet = wb.copy_worksheet(last_sheet)
# 12. 用当天日期命名
today = datetime.now().strftime("%Y%m%d")
if today in wb.sheetnames:
    del wb[today]
new_sheet.title = today
# 13. 保存 Excel
print(f"已复制最后一个工作表，并命名为：{today}")
# 1. 打开 Excel（你前面已经定位到 excel_path 了）
sheet_name = datetime.now().strftime("%Y%m%d")
ws = wb[sheet_name]
today = datetime.now().strftime("%Y.%m.%d")
# 3. 读取标题单元格（通常是 A1）
title_cell = ws["A1"]
old_title = title_cell.value
# 4. 用正则替换括号里的日期
new_title = re.sub(r"[（(].*?[）)]", f"（{today}）", old_title)
# 5. 写回单元格
title_cell.value = new_title
# 6. 保存
wb.save(excel_path)
print("表头更改已完成")
