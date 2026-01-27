from matplotlib import font_manager, rcParams
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from pathlib import Path
import matplotlib.dates as mdates
from datetime import datetime
import textwrap
from matplotlib.backends.backend_pdf import PdfPages

# -----------------------------
# 0️⃣ 中文字体适配（macOS优先）
# -----------------------------
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

# 读取标题（A1 合并单元格）
wb = load_workbook(excel_file, read_only=True)
ws = wb.active
chart_title = ws['A1'].value
print(f"甘特图标题：{chart_title}")

# -----------------------------
# 2️⃣ 读取数据（工程级时间：保留 datetime，不要 .dt.date）
# -----------------------------
df = pd.read_excel(excel_file, header=1)
df = df[['业务类型', '井号', '施工队伍', '施工工序', '开始日期', '结束日期', '时长（天）']]

df["业务类型"] = df["业务类型"].ffill()
df["井号"] = df["井号"].ffill()
df["施工队伍"] = df["施工队伍"].ffill()

df['开始日期'] = pd.to_datetime(df['开始日期'])
df['结束日期'] = pd.to_datetime(df['结束日期'])

# 清理非法行（工程加固）
df = df.dropna(subset=["开始日期", "结束日期", "施工工序", "井号", "业务类型"])

bad = df[df["结束日期"] < df["开始日期"]]
if not bad.empty:
    raise ValueError(
        "发现结束日期早于开始日期，请检查Excel：\n"
        + bad[["业务类型", "井号", "施工工序", "开始日期", "结束日期"]].to_string(index=False)
    )

df = df.reset_index(drop=True)

# -----------------------------
# 3️⃣ 工序颜色函数
# -----------------------------
def get_process_color(process_name: str) -> str:
    if any(k in process_name for k in ["立井架", "搬迁", "上井", "施工准备", "设备安装", "就位", "压前准备", "开工验收"]):
        return "#4CAF50"  # 绿色（准备类）
    elif any(k in process_name for k in ["打包", "放井架", "归拢", "撤场", "交井", "拆解", "收尾"]):
        return "#FF9800"  # 橙色（收尾类）
    else:
        return "#2196F3"  # 蓝色（施工类，默认）

# -----------------------------
# 4️⃣ 工程级时间引擎：闭区间（含结束日）→ 右开区间（结束+1天00:00不含）
# -----------------------------
def to_engineering_interval(start_dt: pd.Timestamp, end_dt: pd.Timestamp):
    """
    工程口径：开始/结束 都包含当天（闭区间）
    画图口径：Matplotlib bar 采用左闭右开 [start, end_exclusive)
    所以：end_exclusive = 结束日00:00 + 1天
    """
    start = pd.to_datetime(start_dt).normalize()  # 当天 00:00
    end_exclusive = pd.to_datetime(end_dt).normalize() + pd.Timedelta(days=1)
    if end_exclusive <= start:
        raise ValueError(f"日期区间异常：start={start} end={end_dt}")
    return start, end_exclusive

def inclusive_days(start_dt: pd.Timestamp, end_dt: pd.Timestamp) -> int:
    """工程闭区间天数：同日=1天"""
    s = pd.to_datetime(start_dt).normalize()
    e = pd.to_datetime(end_dt).normalize()
    return (e - s).days + 1

# -----------------------------
# 5️⃣ 多业务循环绘图，输出同一 PDF
# -----------------------------
业务列表 = df['业务类型'].dropna().unique()

pdf_path = excel_file.with_name(f"{excel_file.stem}.pdf")
with PdfPages(pdf_path) as pdf:
    for 业务 in 业务列表:
        df_业务 = df[df['业务类型'] == 业务].copy()

        # 业务级最小开始（含）/最大结束（不含）
        min_date = pd.to_datetime(df_业务["开始日期"]).min().normalize()
        max_date = pd.to_datetime(df_业务["结束日期"]).max().normalize() + pd.Timedelta(days=1)  # 结束不含
        date_span = (max_date - min_date).days  # 不再 +1

        fig, ax = plt.subplots(figsize=(12, max(6, len(df_业务) * 0.5)))

        def format_text_by_days(text, days):
            if days >= 6:
                return text, 10
            elif days > 1:
                return "\n".join(textwrap.wrap(text, width=7)), 9
            else:
                return text, 9

        y_base = 0
        y_ticks = []
        y_labels = []

        # 按井号分组（保持 Excel 原顺序）
        for 井号, df_井 in df_业务.groupby("井号", sort=False):
            井开始_y = y_base

            for _, row in df_井.iterrows():
                proc = str(row["施工工序"])
                color = get_process_color(proc)

                # 工程级时间区间
                start_ts, end_excl_ts = to_engineering_interval(row["开始日期"], row["结束日期"])

                # barh 使用数值坐标：date2num
                left_num = mdates.date2num(start_ts.to_pydatetime())
                right_num = mdates.date2num(end_excl_ts.to_pydatetime())
                width_days = right_num - left_num

                ax.barh(
                    y=y_base,
                    width=width_days,
                    left=left_num,
                    height=0.5,
                    color=color,
                    edgecolor="black"
                )

                # 文字布局：用工程闭区间天数决定换行/字号
                days_inc = inclusive_days(row["开始日期"], row["结束日期"])
                text_str, font_size = format_text_by_days(proc, days_inc)

                center_x = left_num + width_days / 2
                right_x = left_num + width_days + 0.2

                if days_inc == 1:
                    ax.text(right_x, y_base, proc, ha='left', va='center', fontsize=9)
                else:
                    ax.text(center_x, y_base, text_str, ha='center', va='center', fontsize=font_size)

                y_ticks.append(y_base)
                y_labels.append(proc)

                y_base += 1

            # 井与井之间留白一行
            y_base += 1

            # 井号标注（左侧）
            ax.text(
                mdates.date2num(min_date.to_pydatetime()) - 0.5,
                井开始_y - 0.5,
                str(井号),
                ha="right",
                va="bottom",
                fontsize=14,
                fontweight="bold",
                color="red"
            )

        # 纵轴设置
        ax.set_yticks(y_ticks)
        ax.set_yticklabels(y_labels)
        ax.invert_yaxis()

        # -----------------------------
        # 横轴刻度：按业务跨度自适应
        # -----------------------------
        if date_span <= 40:
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

        # 标题
        ax.set_title(f"{chart_title}\n业务类型：{业务} ", fontsize=16, fontweight='bold', color='red')

        plt.subplots_adjust(left=0.25)
        plt.tight_layout()

        pdf.savefig(fig)
        plt.close(fig)

print(f"所有业务甘特图已保存到 PDF：{pdf_path}")
