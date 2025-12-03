import datetime
import os
import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename

import pandas as pd
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button, Combobox, Frame, Label, Separator, Treeview

# ReportLab 相关导入（生成PDF用）
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Spacer, SimpleDocTemplate, Table, TableStyle,
                                Paragraph)
from xml.sax.saxutils import escape


# 注册中文字体 SimSun，请确保 simsun.ttf 存在于当前目录
try:
    pdfmetrics.registerFont(TTFont("SimSun", "simsun.ttf"))
    pdfmetrics.registerFontFamily(
        "SimSun",
        normal="SimSun",
        bold="SimSun",
        italic="SimSun",
        boldItalic="SimSun",
    )
except Exception as e:  # pragma: no cover - 环境不一定存在 simsun.ttf
    print("注册中文字体失败，请检查 simsun.ttf 是否存在：", e)


class DashboardApp:
    """保养合同看板应用（UI 与导出报告均做了强化）。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("保养合同跟踪情况")
        # 默认窗口改为中等尺寸，避免初次启动窗口过大需要手动调整
        self.root.geometry("1200x760")
        self.style = Style(theme="superhero")
        self.file_path: str | None = None
        self.year_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.loaded_at: datetime.datetime | None = None
        self.loaded_rows: int = 0
        self.overall_stats: dict[str, int] = {}
        self.block_stats: dict[str, dict[str, int]] = {}
        self._build_ui()

    # ------------------------------- UI 构建 ------------------------------- #
    def _build_ui(self) -> None:
        main_frame = Frame(self.root, padding=16)
        main_frame.pack(padx=12, pady=12, fill="both", expand=True)
        for i in range(6):
            main_frame.columnconfigure(i, weight=1)
        main_frame.rowconfigure(6, weight=1)

        # 顶部标题与说明
        header = Frame(main_frame)
        header.grid(row=0, column=0, columnspan=6, sticky="we")
        header.columnconfigure(0, weight=1)
        title_label = Label(header, text="保养合同跟踪情况", font=("SimSun", 18, "bold"))
        title_label.grid(row=0, column=0, sticky="w")
        self.subtitle_label = Label(
            header,
            text="选择Excel文件并设置筛选条件，点击“加载看板”即可查看",
            font=("SimSun", 11),
            anchor="e",
        )
        self.subtitle_label.grid(row=1, column=0, sticky="we", pady=(2, 8))

        Separator(main_frame, orient="horizontal").grid(
            row=1, column=0, columnspan=6, sticky="we", pady=8
        )

        # 文件、年月选择
        file_frame = Frame(main_frame)
        file_frame.grid(row=2, column=0, columnspan=6, sticky="we", pady=(0, 8))
        file_frame.columnconfigure(1, weight=1)
        Label(file_frame, text="Excel文件:", font=("SimSun", 12)).grid(
            row=0, column=0, sticky="e", padx=(0, 6)
        )
        self.file_label = Label(
            file_frame, text="未选择", width=60, anchor="w", font=("SimSun", 12)
        )
        self.file_label.grid(row=0, column=1, sticky="we")
        Button(file_frame, text="浏览", command=self._browse_file, bootstyle="primary").grid(
            row=0, column=2, padx=6
        )

        now = datetime.datetime.now()
        years = [str(y) for y in range(now.year - 2, now.year + 5)]
        months = [f"{m:02d}" for m in range(1, 13)]
        self.year_var.set(str(now.year))
        self.month_var.set(f"{now.month:02d}")

        Label(file_frame, text="年份:", font=("SimSun", 12)).grid(
            row=1, column=0, sticky="e", padx=(0, 6), pady=(6, 0)
        )
        self.year_combo = Combobox(
            file_frame,
            textvariable=self.year_var,
            values=years,
            state="readonly",
            width=8,
        )
        self.year_combo.grid(row=1, column=1, sticky="w", pady=(6, 0))

        Label(file_frame, text="月份:", font=("SimSun", 12)).grid(
            row=1, column=2, sticky="e", padx=(8, 6), pady=(6, 0)
        )
        self.month_combo = Combobox(
            file_frame,
            textvariable=self.month_var,
            values=months,
            state="readonly",
            width=6,
        )
        self.month_combo.grid(row=1, column=3, sticky="w", pady=(6, 0))

        actions = Frame(file_frame)
        actions.grid(row=2, column=0, columnspan=4, sticky="we", pady=(10, 0))
        actions.columnconfigure(0, weight=1)
        Button(actions, text="加载看板", command=self._load_dashboard, bootstyle="info").grid(
            row=0, column=0, padx=(0, 8), sticky="w"
        )
        Button(actions, text="导出PDF报告", command=self._export_pdf, bootstyle="success").grid(
            row=0, column=1, sticky="w"
        )

        # 文件信息与整体统计
        info_frame = Frame(main_frame)
        info_frame.grid(row=3, column=0, columnspan=6, sticky="we", pady=(4, 6))
        info_frame.columnconfigure(3, weight=1)
        self.file_info_label = Label(info_frame, text="文件：--", font=("SimSun", 11))
        self.file_info_label.grid(row=0, column=0, sticky="w")
        self.loaded_info_label = Label(info_frame, text="记录：0  |  已加载：--", font=("SimSun", 11))
        self.loaded_info_label.grid(row=0, column=1, sticky="w", padx=(10, 0))
        self.stats_label = Label(info_frame, text="", anchor="e", font=("SimSun", 11))
        self.stats_label.grid(row=0, column=3, sticky="e")

        # 快速概览卡片
        self.cards_frame = Frame(main_frame)
        self.cards_frame.grid(row=4, column=0, columnspan=6, sticky="we", pady=(0, 8))
        for i in range(4):
            self.cards_frame.columnconfigure(i, weight=1)
        self.card_labels = {}
        self._create_stat_card("超期未做", "#dc3545", 0, col=0)
        self._create_stat_card("超期已做未回", "#fd7e14", 0, col=1)
        self._create_stat_card("已做未回", "#ffc107", 0, col=2)
        self._create_stat_card("已做已回", "#198754", 0, col=3)

        legend_text = "图例： 红色=超期 | 黄色=当月 | 绿色=未来或已做已回"
        self.legend_label = Label(
            main_frame, text=legend_text, font=("SimSun", 11), anchor="center"
        )
        self.legend_label.grid(row=5, column=0, columnspan=6, pady=4, sticky="we")

        # 看板内容区域
        self.dashboard_frame = Frame(main_frame)
        self.dashboard_frame.grid(row=6, column=0, columnspan=6, sticky="nsew")
        self.dashboard_frame.columnconfigure(0, weight=1)
        self.dashboard_frame.columnconfigure(1, weight=1)
        self.dashboard_frame.columnconfigure(2, weight=1)

        self.red_tree = self._build_block(
            parent=self.dashboard_frame,
            color="red",
            title="超期",
            col=0,
        )
        self.yellow_tree = self._build_block(
            parent=self.dashboard_frame,
            color="yellow",
            title="当月",
            col=1,
        )
        self.green_tree = self._build_block(
            parent=self.dashboard_frame,
            color="green",
            title="未来",
            col=2,
        )

    def _create_stat_card(self, label: str, color: str, value: int, col: int) -> None:
        card = Frame(self.cards_frame, padding=10)
        card.grid(row=0, column=col, padx=6, sticky="we")
        card.configure(bootstyle="light")
        title = Label(card, text=label, font=("SimSun", 11, "bold"))
        title.pack(anchor="w")
        val_label = Label(card, text=str(value), font=("SimSun", 18, "bold"), foreground=color)
        val_label.pack(anchor="w")
        self.card_labels[label] = val_label

    def _build_block(self, parent: Frame, color: str, title: str, col: int) -> Treeview:
        frame = Frame(parent, padding=8)
        frame.grid(row=0, column=col, sticky="nsew")
        parent.columnconfigure(col, weight=1)
        header_frame = Frame(frame)
        header_frame.pack(fill="x")
        circle = Label(header_frame, text="●", font=("SimSun", 12), foreground=color)
        circle.pack(side="left")
        stats_label = Label(
            header_frame,
            text="未做: 0  已做: 0  已做未回: 0  已做已回: 0",
            font=("SimSun", 11),
        )
        stats_label.pack(side="left", padx=6)

        columns = (
            "保养合同号",
            "合同客户名称",
            "保养合同结束日期",
            "状态",
            "合同原件",
            "台数",
        )
        tree = Treeview(frame, columns=columns, show="headings", height=10)
        for col_name in columns:
            tree.heading(col_name, text=col_name)
            width = 130 if col_name != "合同客户名称" else 150
            tree.column(col_name, width=width, anchor="center")
        tree.pack(fill="both", expand=True, pady=(6, 0))
        tree.tag_configure("red", background="#FF8080", foreground="black")
        tree.tag_configure("yellow", background="#FFFF80", foreground="black")
        tree.tag_configure("green", background="#80FF80", foreground="black")

        setattr(self, f"{color}_stats_label", stats_label)
        return tree

    # ------------------------------ 核心逻辑 ------------------------------- #
    def _browse_file(self) -> None:
        path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_path = path
            self.file_label.config(text=os.path.basename(path))
            self.file_info_label.config(text=f"文件：{os.path.basename(path)}")

    def _get_block(self, end_date, start_of_month: datetime.datetime, next_month_start: datetime.datetime) -> str:
        """根据日期判断所属块（过期/当月/未来）。"""
        if pd.notnull(end_date):
            if end_date < start_of_month:
                return "red"
            if start_of_month <= end_date < next_month_start:
                return "yellow"
            return "green"
        return "green"

    def _reset_trees(self) -> None:
        for tree in (self.red_tree, self.yellow_tree, self.green_tree):
            for row_id in tree.get_children():
                tree.delete(row_id)

    def _load_dashboard(self) -> None:
        if not self.file_path:
            messagebox.showwarning("警告", "请先选择文件")
            return
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
        except ValueError:
            messagebox.showerror("错误", "年份或月份无效")
            return
        try:
            df = pd.read_excel(self.file_path)
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel出错:\n{e}")
            return

        required_cols = ["保养合同号", "合同客户名称", "保养合同结束日期", "状态", "合同原件", "台数"]
        for rc in required_cols:
            if rc not in df.columns:
                messagebox.showerror("错误", f"缺少必要列: {rc}")
                return

        df["保养合同结束日期"] = pd.to_datetime(df["保养合同结束日期"], errors="coerce")

        self._reset_trees()
        font = tkFont.Font(family="SimSun", size=9)
        max_cust_width = font.measure("合同客户名称")

        start_of_month = datetime.datetime(year, month, 1)
        next_month_start = (
            datetime.datetime(year + 1, 1, 1) if month == 12 else datetime.datetime(year, month + 1, 1)
        )

        block_stats = {
            "red": {"未做": 0, "已做": 0, "已做未回": 0, "已做已回": 0},
            "yellow": {"未做": 0, "已做": 0, "已做未回": 0, "已做已回": 0},
            "green": {"未做": 0, "已做": 0, "已做未回": 0, "已做已回": 0},
        }

        overdue_not_done = 0
        overdue_done_not_returned = 0
        total_not_done = 0
        total_done = 0
        done_not_returned = 0
        done_returned = 0

        for _, row in df.iterrows():
            contract_no = row["保养合同号"]
            cust_name = row["合同客户名称"]
            end_date = row["保养合同结束日期"]
            status = str(row["状态"]).strip() if pd.notnull(row["状态"]) else ""
            original = str(row["合同原件"]).strip() if pd.notnull(row["合同原件"]) else ""
            taishu = row["台数"] if pd.notnull(row["台数"]) else ""

            if status.lower() == "流失":
                continue

            block = self._get_block(end_date, start_of_month, next_month_start)

            if status.lower() == "未做":
                display_color = block
                block_stats[block]["未做"] += 1
                total_not_done += 1
                if pd.notnull(end_date) and end_date < start_of_month:
                    overdue_not_done += 1
            elif status.lower() == "已做":
                block_stats[block]["已做"] += 1
                total_done += 1
                if original == "未回":
                    block_stats[block]["已做未回"] += 1
                    done_not_returned += 1
                    if pd.notnull(end_date) and end_date < start_of_month:
                        overdue_done_not_returned += 1
                    display_color = block
                elif original == "已回":
                    block_stats[block]["已做已回"] += 1
                    done_returned += 1
                    display_color = "green"
                else:
                    display_color = block
            else:
                display_color = block

            values = (
                contract_no,
                cust_name,
                end_date.strftime("%Y-%m-%d") if pd.notnull(end_date) else "",
                status,
                original,
                taishu,
            )
            target_tree = {
                "red": self.red_tree,
                "yellow": self.yellow_tree,
                "green": self.green_tree,
            }[block]
            target_tree.insert("", "end", values=values, tags=(display_color,))

            cust_width = font.measure(cust_name)
            if cust_width > max_cust_width:
                max_cust_width = cust_width

        stats_text = (
            f"超期未做: {overdue_not_done}  |  "
            f"超期已做未回: {overdue_done_not_returned}  |  "
            f"未做: {total_not_done}  |  "
            f"已做: {total_done}  |  "
            f"已做未回: {done_not_returned}  |  "
            f"已做已回: {done_returned}"
        )
        self.stats_label.config(text=stats_text)

        for tree in (self.red_tree, self.yellow_tree, self.green_tree):
            tree.column("合同客户名称", width=max_cust_width + 14)

        self.block_stats = block_stats
        self.overall_stats = {
            "overdue_not_done": overdue_not_done,
            "overdue_done_not_returned": overdue_done_not_returned,
            "total_not_done": total_not_done,
            "total_done": total_done,
            "done_not_returned": done_not_returned,
            "done_returned": done_returned,
        }
        self.red_stats_label.config(
            text=f"未做: {block_stats['red']['未做']}  已做: {block_stats['red']['已做']}  已做未回: {block_stats['red']['已做未回']}  已做已回: {block_stats['red']['已做已回']}"
        )
        self.yellow_stats_label.config(
            text=f"未做: {block_stats['yellow']['未做']}  已做: {block_stats['yellow']['已做']}  已做未回: {block_stats['yellow']['已做未回']}  已做已回: {block_stats['yellow']['已做已回']}"
        )
        self.green_stats_label.config(
            text=f"未做: {block_stats['green']['未做']}  已做: {block_stats['green']['已做']}  已做未回: {block_stats['green']['已做未回']}  已做已回: {block_stats['green']['已做已回']}"
        )

        self.loaded_rows = len(df)
        self.loaded_at = datetime.datetime.now()
        self.file_info_label.config(text=f"文件：{os.path.basename(self.file_path)}")
        self.loaded_info_label.config(
            text=f"记录：{self.loaded_rows}  |  已加载：{self.loaded_at:%Y-%m-%d %H:%M}"
        )
        self.card_labels["超期未做"].config(text=str(overdue_not_done))
        self.card_labels["超期已做未回"].config(text=str(overdue_done_not_returned))
        self.card_labels["已做未回"].config(text=str(done_not_returned))
        self.card_labels["已做已回"].config(text=str(done_returned))

    # ------------------------------ 导出报告 ------------------------------- #
    def _export_pdf(self) -> None:
        if not self.file_path:
            messagebox.showwarning("警告", "请先选择并加载文件")
            return

        if not self.block_stats:
            messagebox.showwarning("警告", "请先加载看板数据")
            return

        excel_basename = os.path.basename(self.file_path)
        station = ""
        if "江宁" in excel_basename:
            station = "江宁"
        elif "禄口" in excel_basename:
            station = "禄口"
        default_filename = (
            f"{station}站{self.month_var.get()}保养合同跟踪情况.pdf" if station else "dashboard_report.pdf"
        )

        file_path = asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=default_filename,
            title="另存为PDF",
        )
        if not file_path:
            return

        try:
            self._build_pdf(file_path)
            messagebox.showinfo("成功", f"PDF已生成：{file_path}")
        except Exception as e:  # pragma: no cover - GUI环境
            messagebox.showerror("错误", f"生成PDF失败:\n{e}")

    def _build_pdf(self, file_path: str) -> None:
        page_size = A4
        LM = RM = 24
        TM = BM = 24

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "myTitle",
            parent=styles["Normal"],
            fontName="SimSun",
            fontSize=15,
            leading=19,
            alignment=1,
        )
        normal_style = ParagraphStyle(
            "normal",
            parent=styles["Normal"],
            fontName="SimSun",
            fontSize=11,
            leading=14,
        )
        small_style = ParagraphStyle(
            "small",
            parent=styles["Normal"],
            fontName="SimSun",
            fontSize=10,
            leading=12,
        )

        def P(x, style=normal_style):
            s = "" if x is None else str(x)
            return Paragraph(escape(s), style)

        year = self.year_var.get()
        month = self.month_var.get()
        station = ""
        if "江宁" in os.path.basename(self.file_path):
            station = "江宁"
        elif "禄口" in os.path.basename(self.file_path):
            station = "禄口"
        title_text = f"{station}站 {year}年{month}月 保养合同跟踪情况报告"

        headers = ["保养合同号", "合同客户名称", "保养合同结束日期", "状态", "合同原件", "台数"]

        def measure(text, font="SimSun", size=11):
            return pdfmetrics.stringWidth(text, font, size)

        col_widths = [measure(h) + 16 for h in headers]
        for tree in (self.red_tree, self.yellow_tree, self.green_tree):
            for item in tree.get_children():
                vals = tree.item(item, "values")
                for ci, v in enumerate(vals):
                    w = measure("" if v is None else str(v)) + 18
                    col_widths[ci] = max(col_widths[ci], w)
        col_widths[2] = max(col_widths[2], 110)
        min_widths = [80, 150, 110, 70, 70, 50]

        def available_width(psize):
            return psize[0] - LM - RM

        def fit_widths(widths, mins, max_total):
            widths = list(widths)
            if sum(widths) <= max_total:
                return widths
            for _ in range(12):
                total = sum(widths)
                over = total - max_total
                if over <= 0:
                    break
                slack = [w - m for w, m in zip(widths, mins)]
                room = sum(s for s in slack if s > 0)
                if room <= 1e-6:
                    break
                for i, s in enumerate(slack):
                    if s > 0:
                        reduce_i = over * (s / room)
                        widths[i] = max(mins[i], widths[i] - reduce_i)
            return widths

        aw = available_width(page_size)
        fitted = fit_widths(col_widths, min_widths, aw)
        if sum(fitted) > aw + 1:
            page_size = landscape(A4)
            aw = available_width(page_size)
            fitted = fit_widths(col_widths, min_widths, aw)
            total = sum(fitted)
            if total > aw:
                scale = aw / total
                fitted = [w * scale for w in fitted]

        doc = SimpleDocTemplate(
            file_path,
            pagesize=page_size,
            leftMargin=LM,
            rightMargin=RM,
            topMargin=TM,
            bottomMargin=BM,
        )

        story: list = []
        story.append(Paragraph(f"<font size=15>{title_text}</font>", title_style))
        story.append(Spacer(1, 12))

        # 汇总信息
        meta_text = (
            f"数据文件：{os.path.basename(self.file_path)}  |  记录：{self.loaded_rows}  |  "
            f"导出时间：{datetime.datetime.now():%Y-%m-%d %H:%M}"
        )
        story.append(P(meta_text, style=small_style))
        story.append(Spacer(1, 6))

        # 直接进入表格区域，不再插入额外的封面与元信息表格
        story.append(Spacer(1, 12))

        def export_tree(tree):
            data = [tuple(Paragraph(escape(h), normal_style) for h in headers)]
            row_bg = []
            for i, item in enumerate(tree.get_children(), start=1):
                vals = tree.item(item, "values")
                tags = tree.item(item, "tags")
                bg = None
                if tags:
                    t = tags[0]
                    if t == "red":
                        bg = colors.HexColor("#FF8080")
                    elif t == "yellow":
                        bg = colors.HexColor("#FFFF80")
                    elif t == "green":
                        bg = colors.HexColor("#80FF80")
                if bg:
                    row_bg.append((i, bg))
                data.append(tuple(P(v) for v in vals))

            tbl = Table(data, colWidths=fitted, repeatRows=1, hAlign="LEFT")
            ts = TableStyle(
                [
                    ("FONT", (0, 0), (-1, 0), "SimSun"),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONT", (0, 1), (-1, -1), "SimSun"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#ffffff"), colors.HexColor("#f8fafc")]),
                ]
            )
            ts.add("ALIGN", (-1, 1), (-1, -1), "RIGHT")
            for row_idx, bg in row_bg:
                ts.add("BACKGROUND", (0, row_idx), (-1, row_idx), bg)
            tbl.setStyle(ts)
            return tbl

        def section(title_html: str, tree: Treeview):
            story.append(Paragraph(title_html, normal_style))
            story.append(Spacer(1, 6))
            story.append(export_tree(tree))
            story.append(Spacer(1, 12))

        red_title_html = """
            <font size=12><font color="red">●</font> 超期区
            : 未做 {0}  已做 {1}  已做未回 {2}  已做已回 {3}</font>
        """.format(
            self.block_stats.get("red", {}).get("未做", 0),
            self.block_stats.get("red", {}).get("已做", 0),
            self.block_stats.get("red", {}).get("已做未回", 0),
            self.block_stats.get("red", {}).get("已做已回", 0),
        )
        yellow_title_html = """
            <font size=12><font color="yellow">●</font> 当月区
            : 未做 {0}  已做 {1}  已做未回 {2}  已做已回 {3}</font>
        """.format(
            self.block_stats.get("yellow", {}).get("未做", 0),
            self.block_stats.get("yellow", {}).get("已做", 0),
            self.block_stats.get("yellow", {}).get("已做未回", 0),
            self.block_stats.get("yellow", {}).get("已做已回", 0),
        )
        green_title_html = """
            <font size=12><font color="green">●</font> 未来区
            : 未做 {0}  已做 {1}  已做未回 {2}  已做已回 {3}</font>
        """.format(
            self.block_stats.get("green", {}).get("未做", 0),
            self.block_stats.get("green", {}).get("已做", 0),
            self.block_stats.get("green", {}).get("已做未回", 0),
            self.block_stats.get("green", {}).get("已做已回", 0),
        )

        section(red_title_html, self.red_tree)
        section(yellow_title_html, self.yellow_tree)
        section(green_title_html, self.green_tree)

        legend_tbl = Table(
            [
                [
                    Paragraph('<font color="red">●</font> 超期合同', normal_style),
                    Paragraph('<font color="yellow">●</font> 当月合同', normal_style),
                    Paragraph('<font color="green">●</font> 当月之后合同', normal_style),
                    Paragraph('<font color="green">●</font> 已做已回', normal_style),
                ]
            ],
            hAlign="LEFT",
        )
        legend_tbl.setStyle(
            TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")])
        )
        story.append(Spacer(1, 10))
        story.append(legend_tbl)

        doc.build(story)


def main() -> None:
    root = tk.Tk()
    app = DashboardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
