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
from matplotlib.backends.backend_pdf import PdfPages

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
chart_title = ws['A1'].value  # 第1行第1列，第1行是七列合并
print(f"甘特图标题：{chart_title}")

df = pd.read_excel(excel_file, header=1)
df = df[['业务类型', '井号', '施工队伍', '施工工序', '开始日期', '结束日期', '时长（天）']]  # 新增施工队伍列
df["业务类型"] = df["业务类型"].ffill()
df["井号"] = df["井号"].ffill()
df["施工队伍"] = df["施工队伍"].ffill()
df['开始日期'] = pd.to_datetime(df['开始日期']).dt.date
df['结束日期'] = pd.to_datetime(df['结束日期']).dt.date
df['时长（天）'] = (pd.to_datetime(df['结束日期']) - pd.to_datetime(df['开始日期'])).dt.days + 1
df = df.reset_index(drop=True)
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
# 3️⃣ 多井循环绘图，保存到同一 PDF + 单口井 PNG
# -----------------------------
#井列表 = df['井号'].dropna().unique()
业务列表 = df['业务类型'].dropna().unique()

pdf_path = excel_file.with_name(f"{excel_file.stem}.pdf")
with PdfPages(pdf_path) as pdf:
    for 业务 in 业务列表:
        df_业务 = df[df['业务类型'] == 业务].copy()
        grouped_by_well = df_业务.groupby("井号", sort=False)
        start_dates = [datetime.combine(d, datetime.min.time()) for d in df_业务['开始日期']]
        durations = df_业务['时长（天）']

        fig, ax = plt.subplots(figsize=(12, max(6, len(df_业务)*0.5)))

        def format_text_by_days(text, days):
            if days >= 6:
                return text, 10
            elif days > 1:
                return "\n".join(textwrap.wrap(text, width=7)), 9
            else:
                return text, 9


        y_base = 0  # 当前 y 轴画到哪一行
        y_ticks = []  # y 轴刻度位置
        y_labels = []  # y 轴刻度文字

        # 按井号分组后的循环
        for 井号, df_井 in df_业务.groupby("井号", sort=False):

            井开始_y = y_base  # 记录这一口井的起始 y

            for _, row in df_井.iterrows():
                start = datetime.combine(row["开始日期"], datetime.min.time())
                dur = row["时长（天）"]
                proc = row["施工工序"]

                color = get_process_color(proc)

                # 画甘特条
                ax.barh(
                    y=y_base,
                    width=dur,
                    left=start,
                    height=0.5,
                    color=color,
                    edgecolor="black"
                )

                # 写工序文字（沿用你原来的逻辑）
                text_str, font_size = format_text_by_days(proc, dur)
                start_num = mdates.date2num(start)
                center_x = start_num + dur / 2
                right_x = start_num + dur + 0.2

                if dur == 1:
                    ax.text(right_x, y_base, proc,
                            ha='left', va='center', fontsize=9)
                else:
                    ax.text(center_x, y_base, text_str,
                            ha='center', va='center', fontsize=font_size)

                y_ticks.append(y_base)
                y_labels.append(proc)

                y_base += 1  # 下一道工序占用下一行

            # —— 井与井之间留一行空白 ——
            y_base += 1

            # —— 在该井所有工序的最上方标注井号 ——
            ax.text(
                mdates.date2num(min(start_dates)) - 0.5,
                井开始_y - 0.5,
                井号,
                ha="right",
                va="bottom",
                fontsize=14,  # 字号调大（原来是 11）
                fontweight="bold",  # 加粗
                color="red"  # 红色字体
            )

        # 设置纵轴
        ax.set_yticks(y_ticks)
        ax.set_yticklabels(y_labels)
        ax.invert_yaxis()



        # 横轴自动日期刻度
        min_date = min(start_dates)
        max_date = max([start + pd.Timedelta(days=d-1) for start, d in zip(start_dates, durations)])
        date_span = (max_date - min_date).days + 1

        if date_span <= 30:
            locator = mdates.DayLocator(interval=1)
        elif date_span <= 60:
            locator = mdates.DayLocator(interval=2)
        elif date_span <= 200:
            locator = mdates.DayLocator(interval=5)
        elif date_span <= 365:
            locator = mdates.WeekdayLocator(interval=1)
        else:
            locator = mdates.WeekdayLocator(interval=2)

        ax.xaxis.set_major_locator(locator)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
        fig.autofmt_xdate(rotation=45)

        ax.set_xlabel("日期", fontsize=12)
        ax.grid(axis='x', color='#cccccc', linestyle='--', linewidth=0.7)
        ax.grid(axis='y', color='#e6e6e6', linestyle=':', linewidth=0.6)
        ax.set_axisbelow(True)

        # -----------------------------
        # 标题 + 副标题（施工队伍）
        # -----------------------------
        业务类型 = df_业务['业务类型'].iloc[0]
        ax.set_title(f"{chart_title}\n业务类型：{业务类型} ", fontsize=16, fontweight='bold', color='red')
        plt.subplots_adjust(left=0.25)
        plt.tight_layout()

        # 保存 PDF
        pdf.savefig(fig)

print(f"所有井甘特图已保存到 PDF：{pdf_path}")
