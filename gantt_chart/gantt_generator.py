from matplotlib import font_manager, rcParams
# 尝试按顺序寻找可用的中文字体
chinese_fonts = [
    "PingFang SC",
    "Songti SC",
    "STSong",
    "Heiti SC",
    "SimHei",
    "Arial Unicode MS"
]

available_fonts = {f.name for f in font_manager.fontManager.ttflist}

for font in chinese_fonts:
    if font in available_fonts:
        rcParams['font.family'] = font
        break
else:
    raise RuntimeError("系统中未找到可用的中文字体")

rcParams['axes.unicode_minus'] = False

import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from pathlib import Path
import matplotlib.dates as mdates
from datetime import datetime
import textwrap

# -----------------------------
# 1️⃣ 找到桌面 Excel 文件
# -----------------------------
desktop_path = Path("/Users/zhouqingquan/Desktop")
files = list(desktop_path.glob("*施工计划表*.xlsx"))
if not files:
    raise FileNotFoundError("桌面上没有找到包含“施工计划表”的Excel文件")
elif len(files) > 1:
    raise FileExistsError("桌面上找到多个包含“施工计划表”的Excel文件，请确保唯一")
else:
    excel_file = files[0]
    print(f"找到文件：{excel_file}")

wb = load_workbook(excel_file, read_only=True)
ws = wb.active
chart_title = ws['A1'].value  # 第1行第1列，第1行是五列合并
print(f"甘特图标题：{chart_title}")

df = pd.read_excel(excel_file, header=1)
df = df[['井号', '施工工序', '开始日期', '结束日期', '时长（天）']]
df['开始日期'] = pd.to_datetime(df['开始日期']).dt.date
df['结束日期'] = pd.to_datetime(df['结束日期']).dt.date
df['时长（天）'] = (pd.to_datetime(df['结束日期']) - pd.to_datetime(df['开始日期'])).dt.days + 1
df = df.reset_index(drop=True)
print(df)

# -----------------------------
# 2️⃣ 工序颜色函数
# -----------------------------
def get_process_color(process_name: str) -> str:
    if any(k in process_name for k in ["立井架", "搬迁", "上井", "施工准备","设备安装", "就位","压前准备","开工验收"]):
        return "#4CAF50"  # 绿色（准备类）
    elif any(k in process_name for k in ["打包","放井架","归拢", "撤场", "交井","拆解", "收尾"]):
        return "#FF9800"  # 橙色（收尾类）
    else:
        return "#2196F3"  # 蓝色（施工类，默认）

# -----------------------------
# 3️⃣ 绘制甘特图
# -----------------------------
fig, ax = plt.subplots(figsize=(12, max(6, len(df)*0.5)))

start_dates = [datetime.combine(d, datetime.min.time()) for d in df['开始日期']]
durations = df['时长（天）']

def format_text_by_days(text, days):
    if days >= 6:
        return text, 10
    elif days > 1:
        return "\n".join(textwrap.wrap(text, width=7)), 9
    else:
        return text, 9
# 绘制条形及文字（短工期智能避让）
for i, (start, dur, proc) in enumerate(zip(start_dates, durations, df['施工工序'])):
    color = get_process_color(proc)

    # 画条形
    ax.barh(
        y=i,
        width=dur,
        left=start,
        height=0.5,
        color=color,
        edgecolor='black'
    )

    text_str, font_size = format_text_by_days(proc, dur)
    start_num = mdates.date2num(start)
    center_x = start_num + dur / 2
    right_x = start_num + dur + 0.2  # 条形右侧，留一点空隙
    if dur == 1:
        # 🔹 1 天：文字写在条形右侧
        ax.text(
            right_x,
            i,
            proc,
            ha='left',
            va='center',
            fontsize=9,
            color='black'
        )

    else:
        ax.text(
            center_x,
            i,
            text_str,
            ha='center',
            va='center',
            fontsize=font_size,
            color='black',
            wrap=True
        )
# 设置纵轴
ax.set_yticks(range(len(df)))
ax.set_yticklabels(df['施工工序'])
ax.invert_yaxis()

# -----------------------------
# 纵轴标签“工序”动态适配图高度
# -----------------------------
ax.text(
    x=-0.05,             # 横向可微调
    y=1-0.02,
    s="工序",
    fontsize=12,
    fontweight='bold',
    rotation=0,
    va='bottom',
    ha='center',
    transform=ax.transAxes
)

# -----------------------------
# 横轴自动日期刻度
# -----------------------------
min_date = min(start_dates)
max_date = max([start + pd.Timedelta(days=d-1) for start, d in zip(start_dates, durations)])
date_span = (max_date - min_date).days + 1

# 根据跨度自动选择刻度
if date_span <= 30:
    locator = mdates.DayLocator(interval=1)   # 每天
elif date_span <= 50:
    locator = mdates.DayLocator(interval=2)   # 每两天
elif date_span <= 200:
    locator = mdates.WeekdayLocator(interval=1)  # 每周
elif date_span <= 365:
    locator = mdates.MonthLocator(interval=1)   # 每月
else:
    locator = mdates.MonthLocator(interval=3)   # 每季度

ax.xaxis.set_major_locator(locator)
ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
fig.autofmt_xdate(rotation=45)

# 横轴标签
ax.set_xlabel("日期", fontsize=12)
# ✅ 在这里加网格线
ax.grid(axis='x', color='#cccccc', linestyle='--', linewidth=0.7)  # 横向网格，淡灰色虚线
# ✅ 在这里加纵坐标网格线（新增）
ax.grid(axis='y', color='#e6e6e6', linestyle=':', linewidth=0.6)
ax.set_axisbelow(True)  # 网格线在条形下方显示，更美观
# -----------------------------
# 设置标题
# -----------------------------
ax.set_title(chart_title, fontsize=16, fontweight='bold',color='red')
plt.subplots_adjust(left=0.25)
plt.tight_layout()

# -----------------------------
# 保存 PNG / PDF
# -----------------------------
base_name = excel_file.stem
png_path = excel_file.with_name(f"{base_name}.png")
pdf_path = excel_file.with_name(f"{base_name}.pdf")
plt.savefig(png_path, dpi=300, bbox_inches="tight")
plt.savefig(pdf_path, bbox_inches="tight")
plt.savefig(pdf_path)
plt.savefig(png_path, dpi=300, bbox_inches="tight")
print(f"甘特图已保存为：{pdf_path} 和 {png_path}")

plt.show()
