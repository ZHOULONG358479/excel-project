from matplotlib import font_manager, rcParams
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from pathlib import Path
import matplotlib.dates as mdates
import textwrap
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import FixedLocator
import numpy as np

# -----------------------------
# 0️⃣ 中文字体适配（macOS优先）
# -----------------------------
chinese_fonts = ["PingFang SC", "Songti SC", "STSong", "Heiti SC", "SimHei", "Arial Unicode MS"]
available_fonts = {f.name for f in font_manager.fontManager.ttflist}
for font in chinese_fonts:
    if font in available_fonts:
        rcParams["font.family"] = font
        break
else:
    raise RuntimeError("系统中未找到可用的中文字体")
rcParams["axes.unicode_minus"] = False

# -----------------------------
# 1️⃣ 找到桌面 Excel 文件
# -----------------------------
desktop_path = Path("/Users/zhouqingquan/Desktop")
files = list(desktop_path.glob("*施工计划表*.xlsx"))
if not files:
    raise FileNotFoundError("桌面上没有找到包含“施工计划表”的Excel文件")
elif len(files) > 1:
    raise FileExistsError("桌面上找到多个包含“施工计划表”的Excel文件，请确保唯一")
excel_file = files[0]
print(f"找到文件：{excel_file}")

# 读取标题（A1 合并单元格）
wb = load_workbook(excel_file, read_only=True)
ws = wb.active
chart_title = ws["A1"].value
print(f"甘特图标题：{chart_title}")

# -----------------------------
# 2️⃣ 读取数据（工程级时间：保留 datetime）
# -----------------------------
df = pd.read_excel(excel_file, header=1)
df = df[["业务类型", "井号", "施工队伍", "施工工序", "开始日期", "结束日期", "时长（天）"]]

df["业务类型"] = df["业务类型"].ffill()
df["井号"] = df["井号"].ffill()
df["施工队伍"] = df["施工队伍"].ffill()

df["开始日期"] = pd.to_datetime(df["开始日期"])
df["结束日期"] = pd.to_datetime(df["结束日期"])

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
        return "#4CAF50"  # 准备类
    elif any(k in process_name for k in ["打包", "放井架", "归拢", "撤场", "交井", "拆解", "收尾"]):
        return "#FF9800"  # 收尾类
    else:
        return "#2196F3"  # 施工类

# -----------------------------
# 4️⃣ 工程级时间引擎
# -----------------------------
def to_engineering_interval(start_dt: pd.Timestamp, end_dt: pd.Timestamp):
    """闭区间（含结束日）→ 右开区间（结束+1天00:00不含）"""
    start = pd.to_datetime(start_dt).normalize()
    end_exclusive = pd.to_datetime(end_dt).normalize() + pd.Timedelta(days=1)
    if end_exclusive <= start:
        raise ValueError(f"日期区间异常：start={start} end={end_dt}")
    return start, end_exclusive

def inclusive_days(start_dt: pd.Timestamp, end_dt: pd.Timestamp) -> int:
    s = pd.to_datetime(start_dt).normalize()
    e = pd.to_datetime(end_dt).normalize()
    return (e - s).days + 1

def fmt_mmdd_dot(ts: pd.Timestamp) -> str:
    return pd.to_datetime(ts).strftime("%m.%d")

# -----------------------------
# ✅ 5️⃣ Excel 单元格二维网格（底纹 + 竖线 + 横线 + 粗分组线）+ 段中心日期标签
# -----------------------------
def set_excel_2d_grid(ax, min_date: pd.Timestamp, max_date: pd.Timestamp, date_span: int, y_total: int, group_separators=None):
    """
    - 交替底纹（像Excel单元格底色）
    - 竖向边框（实线）
    - 横向行边框（实线，形成二维表格）
    - 井与井之间分组边框（更粗横线）
    - 日期标签放在每段中心（底部x轴）
    返回：用于顶部双轴的 locator/formatter 信息
    """
    if group_separators is None:
        group_separators = []

    # 1) 决定“每段长度”（沿用你原规则）
    if date_span <= 40:
        step_days = 1
        fmt = "%m-%d"
    elif date_span <= 60:
        step_days = 2
        fmt = "%m-%d"
    elif date_span <= 200:
        step_days = 5
        fmt = "%m-%d"
    elif date_span <= 365:
        step_days = 7
        fmt = "%m-%d"
    else:
        step_days = 14
        fmt = "%m-%d"

    # 2) 生成段边界
    edges = pd.date_range(start=min_date, end=max_date, freq=f"{step_days}D")
    if len(edges) == 0 or edges[-1] < max_date:
        edges = edges.append(pd.DatetimeIndex([max_date]))
    centers = edges[:-1] + pd.Timedelta(days=step_days / 2)

    edges_num = mdates.date2num(edges.to_pydatetime())
    centers_num = mdates.date2num(centers.to_pydatetime())

    # 3) X轴刻度：major=中心标签，minor=边界
    ax.xaxis.set_major_locator(FixedLocator(centers_num))
    ax.xaxis.set_major_formatter(mdates.DateFormatter(fmt))
    ax.xaxis.set_minor_locator(FixedLocator(edges_num))

    # 4) 画底纹（交替）
    for i in range(len(edges) - 1):
        if i % 2 == 0:
            left = mdates.date2num(edges[i].to_pydatetime())
            right = mdates.date2num(edges[i + 1].to_pydatetime())
            ax.axvspan(left, right, facecolor="#f7f7f7", alpha=1.0, zorder=0)

    # 5) 画竖线边框（实线）
    y_min = -1
    y_max = max(2, y_total)
    for x in edges_num:
        ax.vlines(x, ymin=y_min, ymax=y_max, colors="#bfbfbf", linewidth=0.8, zorder=1)

    # 6) 画横线边框（实线）——二维表格
    x_left = mdates.date2num(min_date.to_pydatetime())
    x_right = mdates.date2num(max_date.to_pydatetime())
    y_lines = np.arange(-0.5, y_total + 0.5, 1.0)
    ax.hlines(y_lines, xmin=x_left, xmax=x_right, colors="#d0d0d0", linewidth=0.6, zorder=1)

    # 7) 分组粗线
    for y_sep in group_separators:
        ax.hlines(y_sep, xmin=x_left, xmax=x_right, colors="#9e9e9e", linewidth=1.8, zorder=2)

    # 8) 显示范围
    ax.set_xlim(x_left, x_right)

    return {
        "centers_num": centers_num,
        "fmt": fmt,
        "x_left": x_left,
        "x_right": x_right
    }

def apply_top_axis(ax, tick_info):
    """✅ 顶部双轴：显示同样的日期标签"""
    ax_top = ax.twiny()
    ax_top.set_xlim(tick_info["x_left"], tick_info["x_right"])
    ax_top.xaxis.set_major_locator(FixedLocator(tick_info["centers_num"]))
    ax_top.xaxis.set_major_formatter(mdates.DateFormatter(tick_info["fmt"]))
    ax_top.tick_params(axis="x", labelrotation=45)
    ax_top.grid(False)
    ax_top.spines["top"].set_color("#bfbfbf")
    ax_top.spines["top"].set_linewidth(0.8)
    return ax_top

# -----------------------------
# ✅ 6️⃣ 井号分区淡红底块
# -----------------------------
def draw_well_zones(ax, well_spans):
    """
    well_spans: [(y_start, y_end), ...]  y_end 为“最后一道工序的下一行”(exclusive)
    画法：覆盖 y_start-0.5 ~ y_end-0.5（正好包住这些行）
    """
    for (ys, ye) in well_spans:
        ax.axhspan(
            ys - 0.5,
            ye - 0.5,
            xmin=0, xmax=1,
            facecolor="#ffe5e5",  # 淡红
            alpha=0.45,
            zorder=0.4            # 在灰底纹之上，在网格线/条形之下
        )

# -----------------------------
# 7️⃣ 多业务循环绘图，输出同一 PDF
# -----------------------------
业务列表 = df["业务类型"].dropna().unique()
pdf_path = excel_file.with_name(f"{excel_file.stem}.pdf")

with PdfPages(pdf_path) as pdf:
    for 业务 in 业务列表:
        df_业务 = df[df["业务类型"] == 业务].copy()

        min_date = pd.to_datetime(df_业务["开始日期"]).min().normalize()
        max_date = pd.to_datetime(df_业务["结束日期"]).max().normalize() + pd.Timedelta(days=1)
        date_span = (max_date - min_date).days

        fig, ax = plt.subplots(figsize=(14, max(6, len(df_业务) * 0.55)))

        def format_text_by_days(text, days):
            # 工序名+日期更长，换行适当放宽
            if days >= 6:
                return text, 9
            elif days > 1:
                return "\n".join(textwrap.wrap(text, width=14)), 8
            else:
                return "\n".join(textwrap.wrap(text, width=14)), 8

        y_base = 0
        y_ticks = []
        y_labels = []

        group_separators = []   # 分组粗线位置
        well_spans = []         # ✅ 井号分区淡红底块范围：[(y_start, y_end), ...]

        # ---- 先跑数据，记录每口井的行范围 + 画条形 ----
        for 井号, df_井 in df_业务.groupby("井号", sort=False):
            y_start = y_base        # 该井第一行工序
            井开始_y = y_base

            for _, row in df_井.iterrows():
                proc = str(row["施工工序"])
                color = get_process_color(proc)

                # 工程区间（绘图用）
                start_ts, end_excl_ts = to_engineering_interval(row["开始日期"], row["结束日期"])
                left_num = mdates.date2num(start_ts.to_pydatetime())
                right_num = mdates.date2num(end_excl_ts.to_pydatetime())
                width_days = right_num - left_num

                # 标签：工序 + 日期字段（例：修井车施工准备06.03-06.05）
                label = f"{proc}{fmt_mmdd_dot(row['开始日期'])}-{fmt_mmdd_dot(row['结束日期'])}"

                # 条形
                ax.barh(
                    y=y_base,
                    width=width_days,
                    left=left_num,
                    height=0.5,
                    color=color,
                    edgecolor="black",
                    zorder=3
                )

                # 文字
                days_inc = inclusive_days(row["开始日期"], row["结束日期"])
                text_str, font_size = format_text_by_days(label, days_inc)

                center_x = left_num + width_days / 2
                right_x = left_num + width_days + 0.2

                if width_days < 0.6:
                    ax.text(right_x, y_base, label, ha="left", va="center", fontsize=8, zorder=4)
                else:
                    ax.text(
                        center_x, y_base, text_str,
                        ha="center", va="center",
                        fontsize=font_size,
                        bbox=dict(facecolor="white", alpha=0.6, edgecolor="none"),
                        zorder=4
                    )

                y_ticks.append(y_base)
                y_labels.append(proc)
                y_base += 1

            y_end = y_base           # ✅ 该井工序结束的下一行（exclusive）
            well_spans.append((y_start, y_end))

            # ✅ 分组粗线：最后一道工序下边界
            group_separators.append(y_end - 0.5)

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
                color="red",
                zorder=6
            )

        # ---- 画底层：先单元格竖向底纹，再井号淡红分区，再网格线/边框 ----
        tick_info = set_excel_2d_grid(
            ax, min_date, max_date, date_span,
            y_total=y_base,
            group_separators=group_separators
        )

        # ✅ 井号淡红底块（放在灰底纹之上、网格线/条形之下）
        draw_well_zones(ax, well_spans)

        # ✅ 顶部同样日期标签（双轴）
        apply_top_axis(ax, tick_info)

        # 纵轴设置
        ax.set_yticks(y_ticks)
        ax.set_yticklabels(y_labels)
        ax.invert_yaxis()

        # 其他外观
        fig.autofmt_xdate(rotation=45)
        ax.set_xlabel("日期", fontsize=12)

        # 标题
        ax.set_title(f"{chart_title}\n业务类型：{业务}", fontsize=16, fontweight="bold", color="red")

        plt.subplots_adjust(left=0.25)
        plt.tight_layout()

        pdf.savefig(fig)
        plt.close(fig)

print(f"✅ 已输出：{pdf_path}")
