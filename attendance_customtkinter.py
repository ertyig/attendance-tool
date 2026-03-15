#!/usr/bin/env python3
"""CustomTkinter 版本的考勤统计桌面界面。"""

from __future__ import annotations

import os
import platform
import queue
import shutil
import subprocess
import sys
import threading
import traceback
from datetime import date, datetime
from pathlib import Path

if sys.platform == "darwin":
    os.environ.pop("SYSTEM_VERSION_COMPAT", None)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import customtkinter as ctk
except Exception as exc:  # pragma: no cover
    print("Missing dependency: customtkinter", file=sys.stderr)
    print("Install with: python -m pip install customtkinter", file=sys.stderr)
    raise SystemExit(1) from exc

from attendance_report import (
    DATA_DIR,
    OUTPUT_FILE,
    MonthlySourceBundle,
    ReportSummary,
    discover_monthly_source_bundles,
    generate_report,
    get_current_annual_leave_summary,
)

APP_NAME = "考勤统计助手"
APP_VERSION = "v1.0.0"
APP_BG = "#ECF2F6"
CARD_BG = "#FFFFFF"
CARD_ALT = "#F5F8FB"
HEADER_BG = "#123847"
TEXT = "#1E2A34"
TEXT_MUTED = "#617383"
PRIMARY = "#245B7A"
PRIMARY_DARK = "#1B4760"
SUCCESS = "#2D7A58"
SUCCESS_SOFT = "#EAF7F0"
WARNING = "#C56A3D"
WARNING_SOFT = "#FFF2EA"
ACCENT = "#5D87C6"
DANGER = "#C95A4D"
DANGER_SOFT = "#FCEDEA"
BORDER = "#D7E1E8"
INFO_SOFT = "#EEF5FA"

FILE_KIND_CONFIG = {
    "attendance": {
        "label": "考勤打卡记录表",
        "file_names": ["考勤打卡记录表.xls", "考勤打卡记录表.xlsx"],
    },
    "leave": {
        "label": "请假记录表",
        "file_names": ["请假记录表.xls", "请假记录表.xlsx"],
    },
    "annual": {
        "label": "员工年假总数表",
        "file_names": ["员工年假总数表.xls", "员工年假总数表.xlsx"],
    },
}
CURRENT_ANNUAL_TARGET_STEM = "当前员工年假总数表"


def _app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _default_data_dir() -> Path:
    return (_app_root() / DATA_DIR).resolve()


def _default_output_file() -> Path:
    return (_app_root() / OUTPUT_FILE).resolve()


def _detailed_usage_file() -> Path:
    return (_app_root() / "使用说明-给同事看.txt").resolve()


def _fallback_usage_file() -> Path:
    return (_app_root() / "data" / "月度文件" / "使用说明.txt").resolve()


def _startup_log_file() -> Path:
    return (_app_root() / "attendance_customtkinter_startup.log").resolve()


def _open_path(path: Path) -> None:
    target = path if path.exists() else path.parent
    if sys.platform.startswith("win"):
        os.startfile(str(target))  # type: ignore[attr-defined]
        return
    if sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=False)
        return
    subprocess.run(["xdg-open", str(target)], check=False)


def _write_startup_diagnostic(exc: BaseException) -> Path:
    log_file = _startup_log_file()
    lines = [
        f"time: {datetime.now().isoformat()}",
        f"platform: {platform.platform()}",
        f"python: {sys.version}",
        f"cwd: {Path.cwd()}",
        f"script_root: {_app_root()}",
        f"tk_version: {getattr(tk, 'TkVersion', 'unknown')}",
        "",
        f"exception_type: {type(exc).__name__}",
        f"exception: {exc}",
        "",
        traceback.format_exc(),
    ]
    log_file.write_text("\n".join(lines), encoding="utf-8")
    return log_file


def _friendly_scan_error(exc: Exception) -> str:
    message = str(exc)
    if "多个年份" in message:
        return "当前一次只能统计一个年份。请把不同年份的月份文件夹分开放。"
    if "发现重复月份目录" in message:
        return "发现重复月份。请检查是不是两个文件夹其实都是同一个月的数据。"
    if "不一致" in message and "目录" in message:
        return "月份文件夹名字和考勤表里的年月不一致，请检查后再试。"
    if "中找到多个" in message:
        return "某个月份文件夹里放了多个同类文件，请只保留一个。"
    if "未找到考勤打卡文件" in message:
        return "还没有找到考勤文件。请先把每个月的 2 个文件放进月份文件夹。"
    if "未找到当前员工年假总数表" in message:
        return "还没有上传当前年假总数表。请先点击“上传当前年假表”。"
    if "数据目录不存在" in message:
        return "放文件的文件夹不存在。请重新选择正确的文件夹。"
    return message


class AttendanceCustomTkinter(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(f"{APP_NAME} {APP_VERSION} | CustomTkinter")
        self.geometry("1220x960")
        self.minsize(1100, 860)
        self.configure(fg_color=APP_BG)

        self.data_dir_var = tk.StringVar(value=str(_default_data_dir()))
        self.output_file_var = tk.StringVar(value=str(_default_output_file()))
        self.year_var = tk.StringVar(value=str(date.today().year))
        self.month_var = tk.StringVar(value=f"{date.today().month:02d}")
        self.annual_info_var = tk.StringVar(value="当前年假表：未上传")
        self.summary_var = tk.StringVar(value="先上传当前年假表，再选择年月上传该月 2 个表")
        self.status_var = tk.StringVar(value="准备就绪。")
        self.current_bundles: list[MonthlySourceBundle] = []
        self.runtime_logs: list[str] = []
        self.month_issue_details: dict[str, str] = {}

        self.log_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.run_thread: threading.Thread | None = None

        self._configure_tree_style()
        self._build_ui()
        self._refresh_annual_info()
        self._refresh_year_values()
        self._on_selected_month_changed()
        self.after(120, self._poll_log_queue)
        self.after(220, self.scan_bundles)

    def _configure_tree_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview",
            rowheight=38,
            font=("Microsoft YaHei UI", 11),
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
        )
        style.map("Custom.Treeview", background=[("selected", "#DDEBF7")], foreground=[("selected", TEXT)])
        style.configure(
            "Custom.Treeview.Heading",
            font=("Microsoft YaHei UI", 11, "bold"),
            background="#F3F7FA",
            foreground=TEXT,
            relief="flat",
            bordercolor=BORDER,
        )

    def _available_years(self) -> list[int]:
        current_year = date.today().year
        years = set(range(current_year, current_year + 11))
        for bundle in self.current_bundles:
            years.add(bundle.year)
        base_dir = Path(self.data_dir_var.get())
        if base_dir.exists():
            for child in base_dir.iterdir():
                if child.is_dir() and len(child.name) >= 7 and child.name[:4].isdigit():
                    year = child.name[:4]
                    month = child.name[5:7] if len(child.name) >= 7 else ""
                    if year.isdigit() and month.isdigit():
                        years.add(int(year))
        selected_year = self.year_var.get().strip()
        if selected_year.isdigit():
            years.add(int(selected_year))
        return sorted(years)

    def _refresh_year_values(self) -> None:
        values = [str(year) for year in self._available_years()]
        self.year_menu.configure(values=values)
        if self.year_var.get() not in values and values:
            self.year_var.set(values[-1])
        if not self.month_var.get():
            self.month_var.set(f"{date.today().month:02d}")

    def _get_selected_year_month(self) -> tuple[int, int]:
        year_text = self.year_var.get().strip()
        month_text = self.month_var.get().strip()
        if not year_text.isdigit():
            raise ValueError("请先选择年份。")
        if not month_text.isdigit():
            raise ValueError("请先选择月份。")
        year = int(year_text)
        month = int(month_text)
        if month < 1 or month > 12:
            raise ValueError("月份不正确。")
        return year, month

    def _selected_month_folder_name(self) -> str:
        year, month = self._get_selected_year_month()
        return f"{year}-{month:02d}"

    def _set_selected_month(self, year: int, month: int) -> None:
        self.year_var.set(str(year))
        self.month_var.set(f"{month:02d}")
        self._refresh_year_values()
        self._on_selected_month_changed()

    def _on_selected_month_changed(self, _value: object = None) -> None:
        try:
            folder_name = self._selected_month_folder_name()
        except ValueError:
            return
        self.summary_var.set(f"当前准备上传：{folder_name}")
        self.status_var.set("选择好该月后，点击中间的“上传所选月份2个表”。")

    def _build_ui(self) -> None:
        header = ctk.CTkFrame(self, fg_color=HEADER_BG, corner_radius=0, height=118)
        header.pack(fill="x")
        header.pack_propagate(False)

        inner = ctk.CTkFrame(header, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=28, pady=18)
        inner.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(inner, text=APP_NAME, text_color="#F7FBFD", font=ctk.CTkFont(family="Microsoft YaHei UI", size=30, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(inner, text=f"CustomTkinter 版本  {APP_VERSION}", text_color="#D4E4EA", font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)).grid(row=1, column=0, sticky="w", pady=(6, 0))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=22)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(3, weight=0)
        body.grid_rowconfigure(4, weight=1)

        self.notice_card = ctk.CTkFrame(body, fg_color=INFO_SOFT, corner_radius=16, border_width=1, border_color="#D5E3EE")
        self.notice_card.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        self.notice_title_label = ctk.CTkLabel(self.notice_card, text="请按顺序操作：先上传当前年假表，再选年月上传本月两个表，最后生成结果文件。", text_color=PRIMARY, font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"))
        self.notice_title_label.pack(anchor="w", padx=16, pady=(10, 2))
        self.notice_desc_label = ctk.CTkLabel(self.notice_card, text="年假表通常不用每月都传，只有员工年假发生变化时再更新。", text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=10))
        self.notice_desc_label.pack(anchor="w", padx=16, pady=(0, 10))

        ctk.CTkLabel(body, textvariable=self.summary_var, text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=17, weight="bold")).grid(row=1, column=0, sticky="ew")
        ctk.CTkLabel(body, textvariable=self.status_var, text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)).grid(row=2, column=0, sticky="ew", pady=(6, 16))

        self._build_action_card(body).grid(row=3, column=0, sticky="ew", pady=(0, 18))
        self._build_month_card(body).grid(row=4, column=0, sticky="nsew")

    def _build_action_card(self, parent: tk.Widget) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=18, border_width=1, border_color=BORDER)
        card.grid_columnconfigure(0, weight=1)

        top_line = ctk.CTkFrame(card, fg_color=PRIMARY, corner_radius=16, height=4)
        top_line.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))
        ctk.CTkLabel(card, text="本月操作区", text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold")).grid(row=1, column=0, sticky="w", padx=24, pady=(16, 4))
        ctk.CTkLabel(card, text="先维护当前年假表，再选年月上传本月 2 个表，最后生成结果。", text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)).grid(row=2, column=0, sticky="w", padx=24, pady=(0, 14))

        annual_box = ctk.CTkFrame(card, fg_color="#F3F8FC", corner_radius=16, border_width=1, border_color=BORDER)
        annual_box.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 12))
        annual_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(annual_box, text="当前年假表", text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=18, pady=(12, 4))
        annual_btn_row = ctk.CTkFrame(annual_box, fg_color="transparent")
        annual_btn_row.grid(row=0, column=2, sticky="e", padx=18, pady=(10, 8))
        ctk.CTkButton(annual_btn_row, text="上传当前年假表", command=self.upload_current_annual_file, width=130, height=38, corner_radius=10, fg_color="#E6EEF4", text_color="#41586F", hover_color="#D8E2EA").grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(annual_btn_row, text="打开当前年假表", command=self.open_current_annual_file, width=130, height=38, corner_radius=10, fg_color="#EFF4F8", text_color="#4F657B", hover_color="#E3EBF1").grid(row=0, column=1)
        ctk.CTkLabel(annual_box, text="年假总数表只有在员工年假信息变化时才需要更新。", text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=11)).grid(row=1, column=0, columnspan=3, sticky="w", padx=18, pady=(0, 6))
        ctk.CTkLabel(annual_box, textvariable=self.annual_info_var, text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=11)).grid(row=2, column=0, columnspan=3, sticky="w", padx=18, pady=(0, 12))

        box = ctk.CTkFrame(card, fg_color="#F8FBFD", corner_radius=16, border_width=1, border_color=BORDER)
        box.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 16))
        box.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        ctk.CTkLabel(box, text="统计月份", text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=18, weight="bold")).grid(row=0, column=0, columnspan=6, sticky="w", padx=24, pady=(16, 4))
        ctk.CTkLabel(box, text="选择的年月就是本次上传和统计的目标月份。", text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=11)).grid(row=1, column=0, columnspan=6, sticky="w", padx=24, pady=(0, 12))

        ctk.CTkLabel(box, text="年份", text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold")).grid(row=2, column=1, sticky="e", padx=(0, 12), pady=(0, 14))
        self.year_menu = ctk.CTkOptionMenu(box, variable=self.year_var, values=[self.year_var.get()], command=self._on_selected_month_changed, width=150, height=42, fg_color=PRIMARY, button_color=PRIMARY_DARK, button_hover_color=PRIMARY_DARK, dropdown_fg_color="#FFFFFF", dropdown_hover_color="#EAF2F8", text_color="#FFFFFF")
        self.year_menu.grid(row=2, column=2, sticky="ew", padx=(0, 18), pady=(0, 14))

        ctk.CTkLabel(box, text="月份", text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold")).grid(row=2, column=3, sticky="e", padx=(0, 12), pady=(0, 14))
        self.month_menu = ctk.CTkOptionMenu(box, variable=self.month_var, values=[f"{m:02d}" for m in range(1, 13)], command=self._on_selected_month_changed, width=120, height=42, fg_color=PRIMARY, button_color=PRIMARY_DARK, button_hover_color=PRIMARY_DARK, dropdown_fg_color="#FFFFFF", dropdown_hover_color="#EAF2F8", text_color="#FFFFFF")
        self.month_menu.grid(row=2, column=4, sticky="ew", pady=(0, 14))

        button_row = ctk.CTkFrame(box, fg_color="transparent")
        button_row.grid(row=3, column=0, columnspan=6, pady=(6, 12))
        self.upload_button = ctk.CTkButton(button_row, text="上传所选月份2个表", command=self.upload_all_monthly_files, width=210, height=50, corner_radius=12, fg_color=SUCCESS, hover_color="#256549", font=ctk.CTkFont(family="Microsoft YaHei UI", size=15, weight="bold"))
        self.upload_button.grid(row=0, column=0, padx=8)
        self.run_button = ctk.CTkButton(button_row, text="生成结果文件", command=self.run_report, width=190, height=50, corner_radius=12, fg_color=PRIMARY, hover_color=PRIMARY_DARK, font=ctk.CTkFont(family="Microsoft YaHei UI", size=15, weight="bold"))
        self.run_button.grid(row=0, column=1, padx=8)
        self.open_button = ctk.CTkButton(button_row, text="打开生成好的 Excel", command=self.open_output_file, width=210, height=50, corner_radius=12, fg_color="#E7EEF4", text_color="#41586F", hover_color="#D8E2EA", font=ctk.CTkFont(family="Microsoft YaHei UI", size=15, weight="bold"))
        self.open_button.grid(row=0, column=2, padx=8)

        link_row = ctk.CTkFrame(box, fg_color="transparent")
        link_row.grid(row=4, column=0, columnspan=6, pady=(0, 12))
        ctk.CTkButton(link_row, text="打开所选月份文件夹", command=self.open_selected_month_folder, width=160, height=38, corner_radius=10, fg_color="#EFF4F8", text_color="#4F657B", hover_color="#E3EBF1").grid(row=0, column=0, padx=6)
        ctk.CTkButton(link_row, text="打开放文件夹", command=self.open_data_dir, width=130, height=38, corner_radius=10, fg_color="#EFF4F8", text_color="#4F657B", hover_color="#E3EBF1").grid(row=0, column=1, padx=6)

        info = ctk.CTkFrame(card, fg_color=CARD_ALT, corner_radius=14, border_width=1, border_color=BORDER)
        info.grid(row=5, column=0, sticky="ew", padx=18, pady=(0, 18))
        info.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(info, text="放文件位置", text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=11, weight="bold")).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 8))
        ctk.CTkEntry(info, textvariable=self.data_dir_var, height=38, corner_radius=10, border_color=BORDER, fg_color="#FFFFFF", text_color=TEXT).grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(14, 8))
        ctk.CTkButton(info, text="选择文件夹", command=self.choose_data_dir, width=120, height=38, corner_radius=10, fg_color="#E6EEF4", text_color="#41586F", hover_color="#D8E2EA").grid(row=0, column=2, sticky="ew", padx=(0, 16), pady=(14, 8))
        ctk.CTkLabel(info, text="结果文件", text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=11, weight="bold")).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 14))
        ctk.CTkLabel(info, text="固定生成在程序目录，文件名为：考勤统计结果.xlsx", text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=11)).grid(row=1, column=1, columnspan=2, sticky="w", padx=(0, 16), pady=(0, 14))

        return card

    def _build_month_card(self, parent: tk.Widget) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=18, border_width=1, border_color=BORDER)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(3, weight=1)

        top_line = ctk.CTkFrame(card, fg_color=ACCENT, corner_radius=16, height=4)
        top_line.grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=(16, 0))
        ctk.CTkLabel(card, text="已识别月份", text_color=TEXT, font=ctk.CTkFont(family="Microsoft YaHei UI", size=18, weight="bold")).grid(row=1, column=0, sticky="w", padx=24, pady=(16, 4))
        legend = ctk.CTkFrame(card, fg_color="transparent")
        legend.grid(row=1, column=1, sticky="e", padx=24)
        ctk.CTkLabel(legend, text="[已就绪]", text_color=SUCCESS, font=ctk.CTkFont(family="Microsoft YaHei UI", size=10, weight="bold"), fg_color=SUCCESS_SOFT, corner_radius=10, padx=10, pady=5).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkLabel(legend, text="[需处理]", text_color=DANGER, font=ctk.CTkFont(family="Microsoft YaHei UI", size=10, weight="bold"), fg_color=DANGER_SOFT, corner_radius=10, padx=10, pady=5).grid(row=0, column=1)
        ctk.CTkLabel(card, text="双击某个月份可以直接打开对应文件夹。", text_color=TEXT_MUTED, font=ctk.CTkFont(family="Microsoft YaHei UI", size=11)).grid(row=2, column=0, columnspan=2, sticky="w", padx=24, pady=(0, 14))

        table_wrap = ctk.CTkFrame(card, fg_color="#FFFFFF", corner_radius=14, border_width=1, border_color=BORDER)
        table_wrap.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=18, pady=(0, 14))
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        columns = ("month", "attendance", "leave", "status")
        tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=10, style="Custom.Treeview")
        tree.heading("month", text="月份")
        tree.heading("attendance", text="考勤文件")
        tree.heading("leave", text="请假文件")
        tree.heading("status", text="状态")
        tree.column("month", width=110, anchor="center")
        tree.column("attendance", width=260, anchor="w")
        tree.column("leave", width=260, anchor="w")
        tree.column("status", width=140, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        tree.tag_configure("ok", foreground=SUCCESS, background="#EDF8F2")
        tree.tag_configure("warn", foreground=DANGER, background="#FCEDEA")
        tree.bind("<Double-1>", lambda _event: self.open_tree_selected_month_folder())
        scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=tree.yview)
        scroll.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)
        tree.configure(yscrollcommand=scroll.set)
        self.bundle_tree = tree

        ctk.CTkButton(card, text="打开选中月份文件夹", command=self.open_tree_selected_month_folder, width=180, height=40, corner_radius=10, fg_color="#E7EDF3", text_color="#4F657B", hover_color="#DCE5EC").grid(row=4, column=0, columnspan=2, pady=(0, 18))
        return card

    def log(self, message: str) -> None:
        self.runtime_logs.append(message)
        if len(self.runtime_logs) > 300:
            self.runtime_logs = self.runtime_logs[-300:]

    def set_notice(self, title: str, desc: str, level: str = "info") -> None:
        if level == "success":
            fg = SUCCESS_SOFT
            border = "#C8E5D4"
            title_color = SUCCESS
        elif level == "error":
            fg = DANGER_SOFT
            border = "#F0C9C4"
            title_color = DANGER
        else:
            fg = INFO_SOFT
            border = "#D5E3EE"
            title_color = PRIMARY
        self.notice_card.configure(fg_color=fg, border_color=border)
        self.notice_title_label.configure(text=title, text_color=title_color)
        self.notice_desc_label.configure(text=desc, text_color=TEXT_MUTED)

    def clear_log(self) -> None:
        self.runtime_logs.clear()

    def choose_data_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.data_dir_var.get() or str(_default_data_dir()))
        if selected:
            self.data_dir_var.set(selected)
            self.scan_bundles()

    def open_data_dir(self) -> None:
        data_dir = Path(self.data_dir_var.get())
        data_dir.mkdir(parents=True, exist_ok=True)
        _open_path(data_dir)

    def open_output_file(self) -> None:
        output_file = Path(self.output_file_var.get())
        if not output_file.exists():
            messagebox.showinfo("提示", "结果文件还没有生成。")
            return
        _open_path(output_file)

    def _current_annual_target_path(self, suffix: str = ".xlsx") -> Path:
        data_dir = Path(self.data_dir_var.get())
        data_dir.mkdir(parents=True, exist_ok=True)
        return data_dir / f"{CURRENT_ANNUAL_TARGET_STEM}{suffix}"

    def _refresh_annual_info(self) -> None:
        try:
            summary = get_current_annual_leave_summary(self.data_dir_var.get())
        except Exception:
            self.annual_info_var.set("当前年假表：未上传")
            return
        updated_at = summary["updated_at"].strftime("%Y-%m-%d %H:%M")
        self.annual_info_var.set(f"{summary['file_name']} | 员工数 {summary['employee_count']} | 更新于 {updated_at}")

    def upload_current_annual_file(self) -> None:
        data_dir = Path(self.data_dir_var.get())
        data_dir.mkdir(parents=True, exist_ok=True)
        selected = filedialog.askopenfilename(
            title="请选择当前员工年假总数表",
            initialdir=str(data_dir),
            filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")],
        )
        if not selected:
            return
        source_path = Path(selected)
        if not self._confirm_selected_source("annual", source_path):
            return
        suffix = source_path.suffix.lower()
        if suffix not in {".xls", ".xlsx"}:
            messagebox.showerror("上传失败", "只支持 Excel 文件：.xls 或 .xlsx")
            return
        target_path = self._current_annual_target_path(suffix)
        for old_suffix in (".xls", ".xlsx"):
            old_path = self._current_annual_target_path(old_suffix)
            if old_path.exists() and old_path != target_path:
                old_path.unlink()
        if target_path.exists() and not messagebox.askyesno("确认覆盖", "当前年假总数表已存在，是否用新文件覆盖？"):
            return
        if target_path.exists():
            target_path.unlink()
        shutil.copy2(source_path, target_path)
        self._refresh_annual_info()
        self.log(f"已上传当前年假表 -> {target_path}")
        self.set_notice("当前年假表已更新。", "只有员工年假信息发生变化时，才需要再次上传。", "success")
        self.scan_bundles(preserve_log=True)

    def open_current_annual_file(self) -> None:
        try:
            summary = get_current_annual_leave_summary(self.data_dir_var.get())
        except Exception:
            messagebox.showinfo("提示", "当前还没有上传年假总数表。")
            return
        _open_path(Path(summary["path"]))

    def open_usage_file(self) -> None:
        usage_file = _detailed_usage_file()
        if usage_file.exists():
            _open_path(usage_file)
            return
        fallback_file = _fallback_usage_file()
        if fallback_file.exists():
            _open_path(fallback_file)
            return
        messagebox.showinfo("提示", "未找到使用说明文件。")

    def _populate_bundles(self, bundles: list[MonthlySourceBundle]) -> None:
        folder_rows, issues = self._inspect_month_folders()
        self.current_bundles = list(bundles)
        self.month_issue_details = {
            str(row["month"]): str(row.get("detail", ""))
            for row in folder_rows
            if row.get("detail")
        }
        self._refresh_year_values()
        for item in self.bundle_tree.get_children():
            self.bundle_tree.delete(item)

        if folder_rows:
            for row in folder_rows:
                self.bundle_tree.insert(
                    "",
                    "end",
                    values=(row["month"], row["attendance"], row["leave"], row["status"]),
                    tags=(str(row["tag"]),),
                )
        else:
            for bundle in bundles:
                self.bundle_tree.insert(
                    "",
                    "end",
                    values=(
                        f"{bundle.year}-{bundle.month:02d}",
                        bundle.attendance_file.name,
                        bundle.leave_file.name,
                        "[已就绪]",
                    ),
                    tags=("ok",),
                )

        children = self.bundle_tree.get_children()
        if children:
            self.bundle_tree.selection_set(children[0])
            self.bundle_tree.focus(children[0])

        selected_folder = ""
        try:
            selected_folder = self._selected_month_folder_name()
        except ValueError:
            pass

        if issues:
            self.status_var.set("有月份文件夹需要处理，请先补齐或去重。")
            self.summary_var.set("有文件没放全或放重了，暂时不能生成")
            self.set_notice("有月份文件夹还需要处理。", "请看月份列表的状态标签，显示 [需处理] 的月份先处理。", "error")
        elif bundles:
            if not selected_folder or not (Path(self.data_dir_var.get()) / selected_folder).exists():
                latest = max(self.current_bundles, key=lambda item: (item.year, item.month))
                year, month = latest.year, latest.month + 1
                if month > 12:
                    year += 1
                    month = 1
                self._set_selected_month(year, month)
            self.status_var.set(f"已找到 {len(bundles)} 个月份，年度合计会自动一起计算。")
            self.summary_var.set(f"已识别 {len(bundles)} 个月份，可以直接生成")
            self.set_notice("检查通过，可以直接生成结果文件。", f"当前共识别到 {len(bundles)} 个月份。点击中间的“生成结果文件”即可。", "success")
        else:
            self._on_selected_month_changed()
            self.status_var.set("没有找到可统计的月份。")
            self.summary_var.set("请先上传当前年假表，并上传某个月份的 2 个表")
            self.set_notice("还没有找到可以统计的月份。", "请先上传当前年假表，再选好年份和月份，点击中间的“上传所选月份2个表”。", "info")

    def _has_folder_issues(self) -> bool:
        _, issues = self._inspect_month_folders()
        return bool(issues)

    def _get_next_month_folder_name(self) -> str:
        if self.current_bundles:
            latest = max(self.current_bundles, key=lambda item: (item.year, item.month))
            year, month = latest.year, latest.month + 1
            if month > 12:
                year += 1
                month = 1
            return f"{year}-{month:02d}"
        today = date.today()
        return f"{today.year}-{today.month:02d}"

    def _find_existing_month_files(self, month_dir: Path, kind: str) -> list[Path]:
        return [month_dir / name for name in FILE_KIND_CONFIG[kind]["file_names"] if (month_dir / name).exists()]

    def _inspect_month_folders(self) -> tuple[list[dict[str, object]], list[str]]:
        base_dir = Path(self.data_dir_var.get())
        if not base_dir.exists():
            return [], ["放文件的文件夹不存在。"]
        rows: list[dict[str, object]] = []
        issues: list[str] = []
        for month_dir in sorted([path for path in base_dir.iterdir() if path.is_dir()]):
            attendance_files = self._find_existing_month_files(month_dir, "attendance")
            leave_files = self._find_existing_month_files(month_dir, "leave")
            problems = []
            if not attendance_files:
                problems.append("缺少考勤文件")
            elif len(attendance_files) > 1:
                problems.append("重复考勤文件")
            if not leave_files:
                problems.append("缺少请假文件")
            elif len(leave_files) > 1:
                problems.append("重复请假文件")
            if problems:
                status_text = "[需处理]"
                issues.append(f"{month_dir.name}: {', '.join(problems)}")
                detail_text = "；".join(problems)
                tag = "warn"
            else:
                status_text = "[已就绪]"
                detail_text = ""
                tag = "ok"
            rows.append({
                "month": month_dir.name,
                "attendance": attendance_files[0].name if len(attendance_files) == 1 else ("未放" if not attendance_files else f"{len(attendance_files)}个文件"),
                "leave": leave_files[0].name if len(leave_files) == 1 else ("未放" if not leave_files else f"{len(leave_files)}个文件"),
                "status": status_text,
                "detail": detail_text,
                "tag": tag,
            })
        return rows, issues

    def _ensure_selected_month_dir(self) -> Path:
        base_dir = Path(self.data_dir_var.get())
        base_dir.mkdir(parents=True, exist_ok=True)
        month_dir = base_dir / self._selected_month_folder_name()
        month_dir.mkdir(parents=True, exist_ok=True)
        return month_dir

    def _ensure_selected_month_dir_or_warn(self) -> Path | None:
        try:
            return self._ensure_selected_month_dir()
        except ValueError as exc:
            self.summary_var.set("请先选好年份和月份")
            self.status_var.set("年份和月份还没有选择完整。")
            self.set_notice("请选择年份和月份。", str(exc), "error")
            messagebox.showwarning("请选择年月", str(exc))
            return None

    def _selected_tree_month_dir(self) -> Path | None:
        selected = self.bundle_tree.selection()
        if not selected:
            return None
        values = self.bundle_tree.item(selected[0], "values")
        if not values:
            return None
        return Path(self.data_dir_var.get()) / str(values[0])

    def _selected_tree_issue_detail(self) -> str:
        selected = self.bundle_tree.selection()
        if not selected:
            return ""
        values = self.bundle_tree.item(selected[0], "values")
        if not values:
            return ""
        return self.month_issue_details.get(str(values[0]), "")

    def open_selected_month_folder(self) -> None:
        month_dir = self._ensure_selected_month_dir_or_warn()
        if month_dir is None:
            return
        guide_file = month_dir / "请把这2个文件放到这里.txt"
        if not guide_file.exists():
            guide_file.write_text(
                "\n".join([
                    "请把下面 2 个文件放到这个文件夹里：",
                    "",
                    "1. 考勤打卡记录表.xls",
                    "2. 请假记录表.xls",
                    "",
                    "当前年假总数表不放在这里，请在程序里单独上传。",
                    "",
                    "也可以直接在程序界面里点击上传按钮。",
                ]),
                encoding="utf-8",
            )
        self.log(f"已打开月份文件夹：{month_dir}")
        self.set_notice(f"已准备好 {month_dir.name} 文件夹。", "你可以直接把 Excel 拖进去，也可以回到程序点击上传按钮。", "info")
        _open_path(month_dir)

    def open_tree_selected_month_folder(self) -> None:
        month_dir = self._selected_tree_month_dir()
        if month_dir is None:
            messagebox.showinfo("提示", "请先在月份列表里点中一个月份。")
            return
        month_dir.mkdir(parents=True, exist_ok=True)
        try:
            year = int(month_dir.name[:4])
            month = int(month_dir.name[5:7])
            self._set_selected_month(year, month)
        except Exception:
            pass
        issue_detail = self._selected_tree_issue_detail()
        if issue_detail:
            messagebox.showinfo("该月份需要处理", f"{month_dir.name} 当前问题：\n{issue_detail}")
        self.set_notice(f"已打开 {month_dir.name} 文件夹。", "如果这个月份缺文件，可以直接拖进去，或者点击上传按钮补齐。", "info")
        _open_path(month_dir)

    def _confirm_selected_source(self, kind: str, source_path: Path) -> bool:
        expected_label = FILE_KIND_CONFIG[kind]["label"]
        if expected_label in source_path.stem:
            return True
        return messagebox.askyesno(
            "确认文件",
            f"你选择的文件名是：{source_path.name}\n\n这个文件名里没有明显看到“{expected_label}”。\n如果选错文件，后面统计会出错。\n\n要继续上传吗？",
        )

    def _copy_monthly_file(self, month_dir: Path, kind: str, source_path: Path) -> Path:
        suffix = source_path.suffix.lower()
        if suffix not in {".xls", ".xlsx"}:
            raise ValueError("只支持 Excel 文件：.xls 或 .xlsx")
        existing_files = [month_dir / name for name in FILE_KIND_CONFIG[kind]["file_names"] if (month_dir / name).exists()]
        if existing_files:
            overwrite = messagebox.askyesno("确认覆盖", f"{month_dir.name} 文件夹里已经有“{FILE_KIND_CONFIG[kind]['label']}”。\n是否用新文件覆盖？")
            if not overwrite:
                raise RuntimeError("用户取消覆盖")
            for existing in existing_files:
                existing.unlink()
        target_path = month_dir / f"{FILE_KIND_CONFIG[kind]['label']}{suffix}"
        shutil.copy2(source_path, target_path)
        return target_path

    def upload_all_monthly_files(self) -> None:
        month_dir = self._ensure_selected_month_dir_or_warn()
        if month_dir is None:
            return
        upload_plan = [
            ("attendance", FILE_KIND_CONFIG["attendance"]["label"]),
            ("leave", FILE_KIND_CONFIG["leave"]["label"]),
        ]
        uploaded = 0
        for kind, label in upload_plan:
            selected = filedialog.askopenfilename(
                title=f"请选择{label}",
                initialdir=str(month_dir if month_dir.exists() else Path.home()),
                filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")],
            )
            if not selected:
                if uploaded == 0:
                    self.set_notice("还没有上传文件。", "你可以重新点击“上传所选月份2个表”。", "info")
                else:
                    self.set_notice("已上传部分文件。", "请继续上传剩余文件，2 个文件齐全后直接点“生成结果文件”。", "info")
                return
            source_path = Path(selected)
            if not self._confirm_selected_source(kind, source_path):
                self.set_notice("已暂停上传。", "你取消了当前文件的确认。可以重新点击“上传所选月份2个表”。", "info")
                return
            try:
                target_path = self._copy_monthly_file(month_dir, kind, source_path)
            except RuntimeError:
                continue
            except Exception as exc:
                messagebox.showerror("上传失败", str(exc))
                return
            uploaded += 1
            self.log(f"已上传 {label} -> {target_path}")
        self.summary_var.set(f"{month_dir.name} 的 2 个文件已上传完成")
        self.status_var.set(f"{month_dir.name} 文件夹文件已齐，请点击“生成结果文件”。")
        self.set_notice(f"{month_dir.name} 的 2 个文件已上传完成。", "现在可以直接点击“生成结果文件”。程序会自动检查后再生成。", "success")
        self.scan_bundles(preserve_log=True)

    def scan_bundles(self, preserve_log: bool = False) -> None:
        data_dir = Path(self.data_dir_var.get())
        try:
            bundles = discover_monthly_source_bundles(str(data_dir))
        except Exception as exc:
            friendly_message = _friendly_scan_error(exc)
            self._populate_bundles([])
            self._refresh_annual_info()
            self.status_var.set("检查失败，请先处理文件夹中的问题。")
            self.set_notice("文件检查失败。", friendly_message, "error")
            if not preserve_log:
                self.clear_log()
            self.log(f"检查失败：{friendly_message}")
            self.log(str(exc))
            return
        if not preserve_log:
            self.clear_log()
            self.log(f"放文件的文件夹：{data_dir}")
            try:
                annual_summary = get_current_annual_leave_summary(str(data_dir))
                self.log(f"当前年假表：{annual_summary['file_name']} | 员工数={annual_summary['employee_count']}")
            except Exception:
                self.log("当前年假表：未上传")
            for bundle in bundles:
                self.log(f"找到月份 {bundle.year}-{bundle.month:02d} | {bundle.attendance_file.name} | {bundle.leave_file.name}")
        self._refresh_annual_info()
        self._populate_bundles(bundles)

    def run_report(self) -> None:
        if self.run_thread and self.run_thread.is_alive():
            return
        self.scan_bundles()
        if not self.current_bundles:
            self.summary_var.set("还没有可统计的月份")
            self.status_var.set("请先上传当前年假表，并上传该月 2 个表。")
            self.set_notice("还没有找到可统计的月份。", "请先上传当前年假表，再选好年份和月份，点击“上传所选月份2个表”。", "error")
            messagebox.showinfo("无法生成", "还没有找到可统计的月份。请先上传当前年假表，并上传至少一个月份的 2 个表。")
            return
        if self._has_folder_issues():
            self.summary_var.set("有文件没放全，不能生成")
            self.status_var.set("请先补齐月份列表里缺少的文件。")
            self.set_notice("暂时不能生成。", "月份列表里还有“缺少”或“重复”的项目，请先处理后再生成。", "error")
            messagebox.showwarning("无法生成", "还有月份文件夹缺少文件，请先补齐后再生成。")
            return
        data_dir = Path(self.data_dir_var.get())
        output_file = Path(self.output_file_var.get())
        output_file.parent.mkdir(parents=True, exist_ok=True)
        self.clear_log()
        self.log(f"开始生成统计。放文件的文件夹：{data_dir}")
        self.log(f"结果会保存到：{output_file}")
        self.upload_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.summary_var.set("正在生成，请稍等 10-30 秒")
        self.set_notice("正在生成结果文件，请稍等。", "生成过程中不要重复点击按钮。完成后会自动提示你打开结果文件。", "info")

        def worker() -> None:
            try:
                summary = generate_report(str(data_dir), str(output_file), logger=lambda msg: self.log_queue.put(("log", msg)))
                self.log_queue.put(("done", summary))
            except Exception:
                self.log_queue.put(("error", traceback.format_exc()))

        self.run_thread = threading.Thread(target=worker, daemon=True)
        self.run_thread.start()

    def _poll_log_queue(self) -> None:
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "log":
                    self.log(str(payload))
                elif kind == "done":
                    summary = payload
                    assert isinstance(summary, ReportSummary)
                    self.log("统计完成。")
                    self.summary_var.set(f"已经生成完成，共统计 {len(summary.monthly_results)} 个月份")
                    self.status_var.set(f"结果文件已生成：{summary.output_file}")
                    self.set_notice("结果文件已生成完成。", f"保存位置：{summary.output_file}", "success")
                    messagebox.showinfo("生成完成", f"结果文件已生成：\n{summary.output_file}")
                    self.scan_bundles(preserve_log=True)
                elif kind == "error":
                    self.log("生成失败：")
                    self.log(str(payload))
                    self.summary_var.set("生成失败，请检查放进去的文件")
                    self.status_var.set("生成失败，请检查文件是否完整、文件名是否正确。")
                    self.set_notice("生成失败。", "请检查当前年假表和各月份的 2 个文件是否完整、文件名是否正确。", "error")
                    messagebox.showerror("生成失败", "统计过程中出现错误。\n请检查文件是否完整、文件名是否正确。")
                if kind in {"done", "error"}:
                    self.upload_button.configure(state="normal")
                    self.run_button.configure(state="normal")
        except queue.Empty:
            pass
        self.after(120, self._poll_log_queue)


def main() -> None:
    try:
        app = AttendanceCustomTkinter()
        app.mainloop()
    except Exception as exc:
        log_file = _write_startup_diagnostic(exc)
        print("CustomTkinter GUI startup failed.", file=sys.stderr)
        print(f"Diagnostic log: {log_file}", file=sys.stderr)
        print(str(exc), file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
