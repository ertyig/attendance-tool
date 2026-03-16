#!/usr/bin/env python3
"""Windows 友好的考勤统计桌面界面。"""

from __future__ import annotations

import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
from datetime import datetime
from datetime import date
from fnmatch import fnmatch
from pathlib import Path

# Some macOS setups export SYSTEM_VERSION_COMPAT=1, which makes older
# frameworks see the OS version as 16.x instead of the real 26.x value.
# That can break Tk initialization with messages like:
# "macOS 26 (...) or later required, have instead 16 (...) !"
if sys.platform == "darwin":
    os.environ.pop("SYSTEM_VERSION_COMPAT", None)

import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import INFO, OUTLINE, PRIMARY, SECONDARY, SUCCESS as BS_SUCCESS, WARNING

    BOOTSTRAP_ENABLED = True
except Exception:  # pragma: no cover
    from tkinter import ttk  # type: ignore

    INFO = "info"
    OUTLINE = "outline"
    PRIMARY = "primary"
    SECONDARY = "secondary"
    BS_SUCCESS = "success"
    WARNING = "warning"
    BOOTSTRAP_ENABLED = False

from attendance_report import (
    DATA_DIR,
    OUTPUT_FILE,
    MonthFolderInspection,
    MonthlySourceBundle,
    ReportSummary,
    discover_monthly_source_bundles,
    generate_report,
    get_current_annual_leave_summary,
    inspect_month_source_folders,
)


APP_BG = "#EEF3F6"
CARD_BG = "#FFFFFF"
SURFACE_ALT = "#F6F8FA"
INK = "#1D2A33"
MUTED = "#5E7180"
ACCENT = "#D56A3D"
ACCENT_DARK = "#B9552B"
HEADER_BG = "#153847"
HEADER_FG = "#F4F8FB"
LINE = "#D7E0E7"
SUCCESS = "#2E7D5B"
WARN = "#A86C1E"
ERROR = "#BC4E3C"
NOTICE_BG_INFO = "#EFF5FA"
NOTICE_BG_SUCCESS = "#EAF6EF"
NOTICE_BG_ERROR = "#FAECE8"
CHECK_BG = "#EAF5EF"
CHECK_FG = "#245A40"
RUN_BG = "#2F6FA3"
RUN_FG = "#FFFFFF"
OPEN_BG = "#EEF3F8"
OPEN_FG = "#4C6278"
RUNTIME_LOG_MAX_BYTES = 2 * 1024 * 1024
RUNTIME_LOG_BACKUP_COUNT = 1
FILE_KIND_CONFIG = {
    "attendance": {
        "label": "考勤打卡记录表",
        "file_names": ["考勤打卡记录表.xls", "考勤打卡记录表.xlsx"],
        "patterns": ["*考勤*.xls", "*考勤*.xlsx", "a.xls", "a.xlsx"],
    },
    "leave": {
        "label": "请假记录表",
        "file_names": ["请假记录表.xls", "请假记录表.xlsx"],
        "patterns": ["*请假*.xls", "*请假*.xlsx", "b.xls", "b.xlsx"],
    },
    "annual": {
        "label": "员工年假总数表",
        "file_names": ["员工年假总数表.xls", "员工年假总数表.xlsx", "当前员工年假总数表.xls", "当前员工年假总数表.xlsx"],
        "patterns": ["*年假*.xls", "*年假*.xlsx", "c.xls", "c.xlsx"],
    },
}
CURRENT_ANNUAL_TARGET_STEM = "当前员工年假总数表"
APP_NAME = "考勤统计助手"
APP_VERSION = "v1.0.0"


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
    return (_app_root() / "attendance_gui_startup.log").resolve()


def _runtime_log_file() -> Path:
    return (_app_root() / "attendance_runtime.log").resolve()


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


def _parse_month_input(raw_text: str) -> tuple[int, int] | None:
    text = raw_text.strip()
    if not text:
        return None

    normalized = (
        text.replace("_", "-")
        .replace("/", "-")
        .replace(".", "-")
        .replace("年", "-")
        .replace("月", "")
        .strip()
    )
    match = re.fullmatch(r"(\d{4})-(\d{1,2})", normalized)
    if not match:
        return None

    year = int(match.group(1))
    month = int(match.group(2))
    if month < 1 or month > 12:
        return None
    return year, month


def _find_matching_files(directory: Path, patterns: list[str]) -> list[Path]:
    if not directory.exists():
        return []
    matches: list[Path] = []
    seen: set[Path] = set()
    for child in sorted(directory.iterdir()):
        if not child.is_file():
            continue
        lower_name = child.name.lower()
        if any(fnmatch(lower_name, pattern.lower()) for pattern in patterns):
            resolved = child.resolve()
            if resolved not in seen:
                matches.append(child)
                seen.add(resolved)
    return matches


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


class AttendanceGui(ttk.Window if BOOTSTRAP_ENABLED else tk.Tk):
    def __init__(self) -> None:
        if BOOTSTRAP_ENABLED:
            super().__init__(themename="litera")
        else:
            super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1180x920")
        self.minsize(1080, 840)
        self.configure(bg=APP_BG)

        self.data_dir_var = tk.StringVar(value=str(_default_data_dir()))
        self.output_file_var = tk.StringVar(value=str(_default_output_file()))
        self.year_var = tk.StringVar(value=str(date.today().year))
        self.month_var = tk.StringVar(value=f"{date.today().month:02d}")
        self.annual_info_var = tk.StringVar(value="当前年假表：未上传")
        self.summary_var = tk.StringVar(value="先上传当前年假表，再选择年月上传该月 2 个表")
        self.status_var = tk.StringVar(value="准备就绪。")
        self.selected_month_state_var = tk.StringVar(value="当前月份状态：未检查")
        self.selected_month_detail_var = tk.StringVar(value="选择年份和月份后，会显示该月是否已经有完整文件。")
        self.selected_month_attendance_var = tk.StringVar(value="考勤文件：未检查")
        self.selected_month_leave_var = tk.StringVar(value="请假文件：未检查")
        self.issue_only_var = tk.BooleanVar(value=False)
        self.current_bundles: list[MonthlySourceBundle] = []
        self.runtime_logs: list[str] = []
        self._runtime_log_buffer: list[str] = []
        self.month_issue_details: dict[str, str] = {}
        self.last_scan_error_message: str = ""
        self.runtime_log_path = _runtime_log_file()
        self._scan_cache_key: object | None = None
        self._scan_cache_bundles: list[MonthlySourceBundle] = []
        self._scan_cache_error: str = ""
        self._last_full_scan_key: object | None = None
        self._last_scan_used_cache = False
        self._last_scan_feedback = ""
        self._folder_cache_key: object | None = None
        self._folder_cache_inspections: list[MonthFolderInspection] = []
        self._last_scan_duration_seconds = 0.0

        self.log_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.run_thread: threading.Thread | None = None
        self._closing = False
        self._poll_after_id: str | None = None
        self._startup_scan_after_id: str | None = None
        self._runtime_log_flush_after_id: str | None = None

        self._configure_style()
        self._build_ui()
        self._start_runtime_log_session()
        self._refresh_annual_info()
        self._refresh_year_values()
        self._on_selected_month_changed()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_after_id = self.after(120, self._poll_log_queue)
        self._startup_scan_after_id = self.after(200, self.refresh_folder_overview)

    def _on_close(self) -> None:
        self._closing = True
        for after_id in (self._poll_after_id, self._startup_scan_after_id, self._runtime_log_flush_after_id):
            if after_id:
                try:
                    self.after_cancel(after_id)
                except Exception:
                    pass
        self._poll_after_id = None
        self._startup_scan_after_id = None
        self._runtime_log_flush_after_id = None
        self._flush_runtime_log_buffer()
        try:
            self.quit()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    def _build_scan_cache_key(self, data_dir: Path, selected_year: int | None) -> tuple[object, ...]:
        files: list[Path] = []
        files.extend(_find_matching_files(data_dir, FILE_KIND_CONFIG["annual"]["patterns"]))
        if data_dir.exists():
            for child in sorted(data_dir.iterdir()):
                if not child.is_dir():
                    continue
                parsed = _parse_month_input(child.name)
                if parsed is None:
                    continue
                year, _ = parsed
                if selected_year is not None and year != selected_year:
                    continue
                files.extend(_find_matching_files(child, FILE_KIND_CONFIG["attendance"]["patterns"]))
                files.extend(_find_matching_files(child, FILE_KIND_CONFIG["leave"]["patterns"]))
                files.extend(_find_matching_files(child, FILE_KIND_CONFIG["annual"]["patterns"]))
        signature = []
        for path in sorted({p.resolve(): p for p in files}.values(), key=lambda item: str(item)):
            try:
                stat = path.stat()
                signature.append((str(path.resolve()), stat.st_mtime_ns, stat.st_size))
            except FileNotFoundError:
                continue
        return (str(data_dir.resolve()), selected_year, tuple(signature))

    def refresh_folder_overview(self) -> None:
        if self._closing:
            return
        self.last_scan_error_message = ""
        self._refresh_annual_info()
        self._populate_bundles([], full_scan=False)

    def _get_folder_inspections(self, selected_year: int | None) -> list[MonthFolderInspection]:
        data_dir = Path(self.data_dir_var.get())
        cache_key = ("folder", self._build_scan_cache_key(data_dir, selected_year))
        if self._folder_cache_key == cache_key:
            return list(self._folder_cache_inspections)
        inspections = inspect_month_source_folders(str(data_dir), target_year=selected_year)
        self._folder_cache_key = cache_key
        self._folder_cache_inspections = list(inspections)
        return list(inspections)

    def _bs(self, bootstyle: str | None = None) -> dict[str, str]:
        if BOOTSTRAP_ENABLED and bootstyle:
            return {"bootstyle": bootstyle}
        return {}

    def _normalize_bootstyle(self, bootstyle: str | None, widget_kind: str) -> str | None:
        if not bootstyle or not BOOTSTRAP_ENABLED:
            return bootstyle
        if widget_kind in {"label", "frame"} and "outline" in bootstyle:
            return bootstyle.replace("-outline", "")
        return bootstyle

    def _theme_label(
        self,
        parent: tk.Widget,
        text: str | None = None,
        textvariable: tk.StringVar | None = None,
        *,
        bootstyle: str | None = None,
        font: tuple[str, int, str] | tuple[str, int] | None = None,
        foreground: str | None = None,
        background: str | None = None,
        anchor: str = "w",
        justify: str = "left",
        padding: tuple[int, int] | None = None,
    ) -> tk.Widget:
        if BOOTSTRAP_ENABLED:
            kwargs: dict[str, object] = {"anchor": anchor, "justify": justify}
            if text is not None:
                kwargs["text"] = text
            if textvariable is not None:
                kwargs["textvariable"] = textvariable
            if font is not None:
                kwargs["font"] = font
            if padding is not None:
                kwargs["padding"] = padding
            kwargs.update(self._bs(self._normalize_bootstyle(bootstyle, "label")))
            return ttk.Label(parent, **kwargs)

        kwargs2: dict[str, object] = {"anchor": anchor, "justify": justify}
        if text is not None:
            kwargs2["text"] = text
        if textvariable is not None:
            kwargs2["textvariable"] = textvariable
        if font is not None:
            kwargs2["font"] = font
        if foreground is not None:
            kwargs2["fg"] = foreground
        if background is not None:
            kwargs2["bg"] = background
        return tk.Label(parent, **kwargs2)

    def _available_years(self) -> list[int]:
        current_year = date.today().year
        years = set(range(current_year, current_year + 11))
        for bundle in self.current_bundles:
            years.add(bundle.year)

        base_dir = Path(self.data_dir_var.get())
        if base_dir.exists():
            for child in base_dir.iterdir():
                if child.is_dir():
                    parsed = _parse_month_input(child.name)
                    if parsed:
                        years.add(parsed[0])

        selected_year = self.year_var.get().strip()
        if selected_year.isdigit():
            years.add(int(selected_year))
        return sorted(years)

    def _refresh_year_values(self) -> None:
        values = [str(year) for year in self._available_years()]
        self.year_combo.configure(values=values)
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

    def _get_selected_year(self) -> int:
        year, _ = self._get_selected_year_month()
        return year

    def _selected_month_folder_name(self) -> str:
        year, month = self._get_selected_year_month()
        return f"{year}-{month:02d}"

    def _set_selected_month(self, year: int, month: int) -> None:
        self.year_var.set(str(year))
        self.month_var.set(f"{month:02d}")
        self._refresh_year_values()
        self._on_selected_month_changed()

    def _on_selected_month_changed(self) -> None:
        try:
            year, month = self._get_selected_year_month()
            folder_name = f"{year}-{month:02d}"
        except ValueError:
            return
        base_dir = Path(self.data_dir_var.get())
        month_dir = base_dir / folder_name
        attendance_files = self._find_existing_month_files(month_dir, "attendance")
        leave_files = self._find_existing_month_files(month_dir, "leave")
        has_any_data = month_dir.exists() and any(month_dir.iterdir()) if month_dir.exists() else False

        if not attendance_files and not leave_files and not has_any_data:
            self.selected_month_state_var.set(f"{folder_name}：未放文件")
            self.selected_month_detail_var.set("这个月份文件夹里还没有识别到考勤或请假文件。")
            self.selected_month_attendance_var.set("考勤文件：未找到")
            self.selected_month_leave_var.set("请假文件：未找到")
            self.summary_var.set(f"当前准备上传：{folder_name}")
            self.status_var.set(f"{folder_name} 还没有放入文件。选择好该月后，点击“上传所选月份2个表”。")
        elif len(attendance_files) == 1 and len(leave_files) == 1:
            self.selected_month_state_var.set(f"{folder_name}：已就绪")
            self.selected_month_detail_var.set(f"考勤：{attendance_files[0].name}；请假：{leave_files[0].name}")
            self.selected_month_attendance_var.set(f"考勤文件：已找到（{attendance_files[0].name}）")
            self.selected_month_leave_var.set(f"请假文件：已找到（{leave_files[0].name}）")
            self.summary_var.set(f"{folder_name} 已有完整文件")
            self.status_var.set(f"{folder_name} 已识别到考勤和请假文件，可以直接生成结果。")
        else:
            self.selected_month_state_var.set(f"{folder_name}：待处理")
            detail_parts = []
            if not attendance_files:
                detail_parts.append("缺少考勤文件")
            elif len(attendance_files) > 1:
                detail_parts.append("重复考勤文件")
            if not leave_files:
                detail_parts.append("缺少请假文件")
            elif len(leave_files) > 1:
                detail_parts.append("重复请假文件")
            self.selected_month_detail_var.set("；".join(detail_parts) or "该月文件还不完整，请补齐后再生成。")
            if len(attendance_files) == 1:
                self.selected_month_attendance_var.set(f"考勤文件：已找到（{attendance_files[0].name}）")
            elif len(attendance_files) > 1:
                self.selected_month_attendance_var.set("考勤文件：重复")
            else:
                self.selected_month_attendance_var.set("考勤文件：未找到")
            if len(leave_files) == 1:
                self.selected_month_leave_var.set(f"请假文件：已找到（{leave_files[0].name}）")
            elif len(leave_files) > 1:
                self.selected_month_leave_var.set("请假文件：重复")
            else:
                self.selected_month_leave_var.set("请假文件：未找到")
            self.summary_var.set(f"{folder_name} 已有部分文件")
            self.status_var.set(f"{folder_name} 还需处理：{self.selected_month_detail_var.get()}")
        self._apply_selected_file_status_colors()

    def _selected_file_status_color(self, text: str) -> str:
        if "已找到" in text:
            return SUCCESS
        if "重复" in text or "未找到" in text:
            return ERROR
        return INK

    def _apply_selected_file_status_colors(self) -> None:
        if hasattr(self, "selected_month_attendance_label"):
            self.selected_month_attendance_label.configure(fg=self._selected_file_status_color(self.selected_month_attendance_var.get()))
        if hasattr(self, "selected_month_leave_label"):
            self.selected_month_leave_label.configure(fg=self._selected_file_status_color(self.selected_month_leave_var.get()))

    def check_selected_year_files(self) -> None:
        self.scan_bundles()
        self._show_year_check_result()

    def _on_issue_only_changed(self) -> None:
        self._populate_bundles(self.current_bundles)
        self._apply_detail_column_visibility()

    def _apply_detail_column_visibility(self) -> None:
        if not hasattr(self, "bundle_tree"):
            return
        if self.issue_only_var.get():
            self.bundle_tree.column("detail", width=260, minwidth=220, stretch=True)
            self.bundle_tree.heading("detail", text="问题说明")
        else:
            self.bundle_tree.column("detail", width=0, minwidth=0, stretch=False)
            self.bundle_tree.heading("detail", text="")

    def _show_year_check_result(self) -> None:
        try:
            year = self._get_selected_year()
        except ValueError:
            return
        rows, issues = self._inspect_month_folders()
        ready_months = [str(row["month"])[5:7] for row in rows if row.get("tag") == "ok"]
        issue_months = [str(row["month"]) for row in rows if row.get("detail")]

        lines = [f"{year} 年检查结果：", ""]
        if self._last_scan_feedback:
            lines.append(self._last_scan_feedback)
            lines.append("")
        if ready_months:
            lines.append(f"已识别月份：{'、'.join(ready_months)}")
        else:
            lines.append("已识别月份：无")

        if issue_months:
            lines.append(f"待处理月份：{'、'.join(issue_months)}")
            lines.append("")
            lines.extend(issues[:6])
        else:
            lines.append("待处理月份：无")

        messagebox.showinfo("检查结果", "\n".join(lines))

    def _show_generation_success_dialog(self, summary: ReportSummary) -> None:
        processed_months = [f"{item.bundle.month:02d}" for item in summary.monthly_results]
        rows, issues = self._inspect_month_folders()
        skipped_months = [str(row["month"]) for row in rows if row.get("detail")]

        dialog = tk.Toplevel(self)
        dialog.title("生成完成")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg="#F7FAFC")
        dialog.resizable(False, False)

        box = tk.Frame(dialog, bg="#FFFFFF", highlightbackground=LINE, highlightthickness=1, bd=0)
        box.pack(fill="both", expand=True, padx=18, pady=18)

        tk.Label(box, text="结果文件已生成", bg="#FFFFFF", fg=INK, font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w", padx=18, pady=(18, 8))
        tk.Label(
            box,
            text="\n".join(
                [
                    f"保存位置：{summary.output_file}",
                    f"已统计月份：{'、'.join(processed_months) if processed_months else '无'}",
                    f"待处理月份：{'、'.join(skipped_months) if skipped_months else '无'}",
                ]
            ),
            bg="#FFFFFF",
            fg=INK,
            justify="left",
            anchor="w",
            font=("Microsoft YaHei UI", 11),
        ).pack(anchor="w", padx=18)
        if issues:
            tk.Label(
                box,
                text="\n".join(issues[:5]),
                bg="#FFFFFF",
                fg=MUTED,
                justify="left",
                anchor="w",
                font=("Microsoft YaHei UI", 10),
            ).pack(anchor="w", padx=18, pady=(10, 0))

        btn_row = tk.Frame(box, bg="#FFFFFF")
        btn_row.pack(anchor="e", padx=18, pady=(16, 18))
        ttk.Button(btn_row, text="打开 Excel", style="App.TButton", command=lambda: (_open_path(summary.output_file), dialog.destroy()), **self._bs(f"{PRIMARY}")).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(btn_row, text="打开结果文件夹", style="App.TButton", command=lambda: (_open_path(summary.output_file.parent), dialog.destroy()), **self._bs(f"{SECONDARY}-{OUTLINE}")).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(btn_row, text="关闭", style="App.TButton", command=dialog.destroy, **self._bs(f"{INFO}-{OUTLINE}")).grid(row=0, column=2)

    def _configure_style(self) -> None:
        style = getattr(self, "style", None) or ttk.Style(self)
        if not BOOTSTRAP_ENABLED:
            style.theme_use("clam")
        style.configure("Panel.TFrame", background=CARD_BG)
        style.configure("PanelTitle.TLabel", background=CARD_BG, foreground=INK, font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Hint.TLabel", background=CARD_BG, foreground=MUTED, font=("Microsoft YaHei UI", 12))
        style.configure("Data.TLabel", background=CARD_BG, foreground=INK, font=("Microsoft YaHei UI", 12))
        style.configure("App.TButton", font=("Microsoft YaHei UI", 12), padding=(18, 12))
        style.map("App.TButton", background=[("active", "#E9EEF2")], foreground=[("!disabled", INK)])
        style.configure("Clean.TCombobox", padding=10, arrowsize=18)
        style.configure(
            "Treeview",
            rowheight=40,
            font=("Microsoft YaHei UI", 12),
            background="white",
            fieldbackground="white",
            foreground=INK,
            bordercolor=LINE,
            lightcolor=LINE,
            darkcolor=LINE,
        )
        style.map(
            "Treeview",
            background=[("selected", "#DCEAF6")],
            foreground=[("selected", INK)],
        )
        style.configure(
            "Treeview.Heading",
            font=("Microsoft YaHei UI", 12, "bold"),
            background="#F3F7FA",
            foreground=INK,
            relief="flat",
            bordercolor=LINE,
        )
        style.map("Treeview.Heading", background=[("active", "#EAF0F4")])

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=HEADER_BG, height=132)
        header.pack(fill="x")
        header.pack_propagate(False)

        header_inner = tk.Frame(header, bg=HEADER_BG)
        header_inner.pack(fill="both", expand=True, padx=30, pady=20)
        header_inner.grid_columnconfigure(0, weight=1)
        header_inner.grid_columnconfigure(1, weight=0)

        title_label = tk.Label(
            header_inner,
            text=APP_NAME,
            bg=HEADER_BG,
            fg=HEADER_FG,
            font=("Microsoft YaHei UI", 31, "bold"),
            anchor="w",
        )
        title_label.grid(row=0, column=0, sticky="w")
        self._theme_label(
            header_inner,
            text=APP_VERSION,
            bootstyle=SECONDARY,
            font=("Microsoft YaHei UI", 10, "bold"),
            foreground="#EAF4F8",
            background="#315464",
            anchor="center",
            justify="center",
            padding=(12, 6),
        ).grid(row=0, column=1, sticky="e", padx=(16, 0))
        body = tk.Frame(self, bg=APP_BG)
        body.pack(fill="both", expand=True, padx=26, pady=22)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._build_workspace(body).grid(row=0, column=0, sticky="nsew")

    def _make_action_button(
        self,
        parent: tk.Widget,
        text: str,
        command: object,
        bg: str,
        fg: str,
        border: str,
        bootstyle: str | None = None,
    ) -> tk.Widget:
        if BOOTSTRAP_ENABLED:
            return ttk.Button(
                parent,
                text=text,
                command=command,
                width=18,
                padding=(22, 18),
                **self._bs(bootstyle or PRIMARY),
            )
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=border,
            activeforeground=fg,
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Microsoft YaHei UI", 17, "bold"),
            padx=30,
            pady=22,
            highlightthickness=2,
            highlightbackground=border,
            wraplength=300,
        )

    def _build_workspace(self, parent: tk.Widget) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Panel.TFrame", padding=0)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=8)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_rowconfigure(2, weight=0)
        frame.grid_rowconfigure(3, weight=1)

        self._build_notice_banner(frame).grid(row=0, column=1, sticky="ew", pady=(0, 16))

        top_hint = tk.Frame(frame, bg=APP_BG)
        top_hint.grid(row=1, column=1, sticky="ew", pady=(0, 14))
        tk.Label(top_hint, textvariable=self.summary_var, bg=APP_BG, fg=INK, font=("Microsoft YaHei UI", 16, "bold"), anchor="center", justify="center").pack(fill="x")
        tk.Label(top_hint, textvariable=self.status_var, bg=APP_BG, fg=MUTED, font=("Microsoft YaHei UI", 12), anchor="center", justify="center").pack(fill="x", pady=(6, 0))

        self._build_action_panel(frame).grid(row=2, column=1, sticky="ew", pady=(0, 18))
        self._build_month_card(frame).grid(row=3, column=1, sticky="nsew")

        return frame

    def _build_notice_banner(self, parent: tk.Widget) -> ttk.Frame:
        card = ttk.Frame(parent, style="Panel.TFrame", padding=0)
        if BOOTSTRAP_ENABLED:
            banner = ttk.Frame(card, padding=0, **self._bs(self._normalize_bootstyle(f"{INFO}-{OUTLINE}", "frame")))
            banner.pack(fill="x")
            self.notice_frame = banner
            self.notice_title_label = self._theme_label(
                banner,
                text="请按顺序操作：先上传当前年假表，再选年月上传本月两个表，最后生成结果文件。",
                bootstyle=INFO,
                font=("Microsoft YaHei UI", 13, "bold"),
                anchor="w",
                padding=(16, 8),
            )
            self.notice_title_label.pack(fill="x")
            self.notice_desc_label = self._theme_label(
                banner,
                text="年假表通常不用每月都传，只有员工年假发生变化时再更新。",
                bootstyle=SECONDARY,
                font=("Microsoft YaHei UI", 10),
                anchor="w",
                padding=(16, 0),
            )
            self.notice_desc_label.pack(fill="x", pady=(0, 10))
            return card

        banner = tk.Frame(card, bg=NOTICE_BG_INFO, highlightbackground="#D4E0EA", highlightthickness=1, bd=0)
        banner.pack(fill="x")
        self.notice_frame = banner
        self.notice_title_label = tk.Label(
            banner,
            text="请按顺序操作：先上传当前年假表，再选年月上传本月两个表，最后生成结果文件。",
            bg=NOTICE_BG_INFO,
            fg="#28485C",
            font=("Microsoft YaHei UI", 13, "bold"),
            anchor="w",
        )
        self.notice_title_label.pack(fill="x", padx=18, pady=(12, 2))
        self.notice_desc_label = tk.Label(
            banner,
            text="年假表通常不用每月都传，只有员工年假发生变化时再更新。",
            bg=NOTICE_BG_INFO,
            fg=MUTED,
            font=("Microsoft YaHei UI", 10),
            anchor="w",
        )
        self.notice_desc_label.pack(fill="x", padx=18, pady=(0, 10))
        return card

    def _build_action_panel(self, parent: tk.Widget) -> ttk.Frame:
        card = ttk.Frame(parent, style="Panel.TFrame", padding=26)
        card.grid_columnconfigure(0, weight=1)

        accent_bar = tk.Frame(card, bg=ACCENT, height=4)
        accent_bar.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        self._theme_label(
            card,
            text="本月操作区",
            bootstyle=PRIMARY,
            font=("Microsoft YaHei UI", 20, "bold"),
            foreground=INK,
            background=CARD_BG,
            anchor="w",
        ).grid(row=1, column=0, sticky="w")
        ttk.Label(card, text="先维护当前年假表，再选年月上传本月 2 个表，最后生成结果。", style="Hint.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 14))

        annual_box = tk.Frame(card, bg="#F3F8FC", highlightbackground=LINE, highlightthickness=1, bd=0)
        annual_box.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        annual_box.grid_columnconfigure(1, weight=1)
        tk.Label(annual_box, text="当前年假表", bg="#F3F8FC", fg=INK, font=("Microsoft YaHei UI", 14, "bold")).grid(row=0, column=0, sticky="w", padx=18, pady=(12, 4))
        annual_btn_row = tk.Frame(annual_box, bg="#F3F8FC")
        annual_btn_row.grid(row=0, column=2, sticky="e", padx=18, pady=(10, 8))
        ttk.Button(annual_btn_row, text="上传当前年假表", style="App.TButton", command=self.upload_current_annual_file, **self._bs(f"{INFO}-{OUTLINE}")).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(annual_btn_row, text="打开当前年假表", style="App.TButton", command=self.open_current_annual_file, **self._bs(f"{SECONDARY}-{OUTLINE}")).grid(row=0, column=1)
        tk.Label(annual_box, text="年假总数表只有在员工年假信息变化时才需要更新。", bg="#F3F8FC", fg=MUTED, font=("Microsoft YaHei UI", 10)).grid(row=1, column=0, columnspan=3, sticky="w", padx=18, pady=(0, 6))
        tk.Label(annual_box, textvariable=self.annual_info_var, bg="#F3F8FC", fg=INK, font=("Microsoft YaHei UI", 11), anchor="w", justify="left").grid(row=2, column=0, columnspan=3, sticky="w", padx=18, pady=(0, 12))

        select_box = tk.Frame(card, bg="#FCF8F3", highlightbackground=LINE, highlightthickness=1, bd=0)
        select_box.grid(row=4, column=0, sticky="ew", pady=(0, 14))
        select_box.grid_columnconfigure(0, weight=1)
        select_box.grid_columnconfigure(1, weight=1)
        select_box.grid_columnconfigure(2, weight=1)
        select_box.grid_columnconfigure(3, weight=1)
        select_box.grid_columnconfigure(4, weight=1)
        select_box.grid_columnconfigure(5, weight=1)

        self._theme_label(
            select_box,
            text="统计月份",
            bootstyle=PRIMARY,
            font=("Microsoft YaHei UI", 18, "bold"),
            foreground=INK,
            background="#FCF8F3",
            anchor="w",
        ).grid(row=0, column=0, columnspan=6, sticky="w", padx=28, pady=(18, 6))
        self._theme_label(
            select_box,
            text="选择的年月就是本次上传和统计的目标月份。",
            bootstyle=SECONDARY,
            font=("Microsoft YaHei UI", 11),
            foreground=MUTED,
            background="#FCF8F3",
            anchor="w",
        ).grid(row=1, column=0, columnspan=6, sticky="w", padx=28, pady=(0, 12))

        tk.Label(select_box, text="年份", bg="#FCF8F3", fg=INK, font=("Microsoft YaHei UI", 13, "bold")).grid(row=2, column=1, sticky="e", padx=(0, 12), pady=(0, 16))
        self.year_combo = ttk.Combobox(
            select_box,
            textvariable=self.year_var,
            state="readonly",
            values=[],
            style="Clean.TCombobox",
            font=("Microsoft YaHei UI", 13),
            **self._bs(f"{INFO}-{OUTLINE}"),
        )
        self.year_combo.grid(row=2, column=2, sticky="ew", padx=(0, 18), pady=(0, 14), ipady=3)

        tk.Label(select_box, text="月份", bg="#FCF8F3", fg=INK, font=("Microsoft YaHei UI", 13, "bold")).grid(row=2, column=3, sticky="e", padx=(0, 12), pady=(0, 16))
        self.month_combo = ttk.Combobox(
            select_box,
            textvariable=self.month_var,
            state="readonly",
            values=[f"{month:02d}" for month in range(1, 13)],
            style="Clean.TCombobox",
            font=("Microsoft YaHei UI", 13),
            **self._bs(f"{INFO}-{OUTLINE}"),
        )
        self.month_combo.grid(row=2, column=4, sticky="ew", padx=(0, 0), pady=(0, 14), ipady=3)

        month_state_box = tk.Frame(select_box, bg="#F6FAFD", highlightbackground="#D8E4EE", highlightthickness=1, bd=0)
        month_state_box.grid(row=3, column=0, columnspan=6, sticky="ew", padx=18, pady=(0, 12))
        tk.Label(month_state_box, text="当前月份状态", bg="#F6FAFD", fg=MUTED, font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 2))
        tk.Label(month_state_box, textvariable=self.selected_month_state_var, bg="#F6FAFD", fg=INK, font=("Microsoft YaHei UI", 12, "bold"), anchor="w", justify="left").grid(row=1, column=0, sticky="w", padx=14)
        self.selected_month_attendance_label = tk.Label(month_state_box, textvariable=self.selected_month_attendance_var, bg="#F6FAFD", fg=INK, font=("Microsoft YaHei UI", 10), anchor="w", justify="left", wraplength=760)
        self.selected_month_attendance_label.grid(row=2, column=0, sticky="w", padx=14, pady=(4, 0))
        self.selected_month_leave_label = tk.Label(month_state_box, textvariable=self.selected_month_leave_var, bg="#F6FAFD", fg=INK, font=("Microsoft YaHei UI", 10), anchor="w", justify="left", wraplength=760)
        self.selected_month_leave_label.grid(row=3, column=0, sticky="w", padx=14, pady=(2, 0))
        tk.Label(month_state_box, textvariable=self.selected_month_detail_var, bg="#F6FAFD", fg=MUTED, font=("Microsoft YaHei UI", 10), anchor="w", justify="left", wraplength=760).grid(row=4, column=0, sticky="w", padx=14, pady=(2, 10))

        button_row = tk.Frame(select_box, bg="#FCF8F3")
        button_row.grid(row=4, column=0, columnspan=6, pady=(6, 14))

        self.upload_button = self._make_action_button(
            button_row,
            "上传所选月份2个表",
            self.upload_all_monthly_files,
            CHECK_BG,
            CHECK_FG,
            border="#C8D9CC",
            bootstyle=f"{BS_SUCCESS}-{OUTLINE}",
        )
        self.upload_button.grid(row=0, column=0, padx=10)

        self.check_button = self._make_action_button(
            button_row,
            "检查当前年份文件",
            self.check_selected_year_files,
            OPEN_BG,
            OPEN_FG,
            border="#BFD2F2",
            bootstyle=f"{INFO}-{OUTLINE}",
        )
        self.check_button.grid(row=0, column=1, padx=10)

        self.run_button = self._make_action_button(
            button_row,
            "生成结果文件",
            self.run_report,
            RUN_BG,
            RUN_FG,
            border=ACCENT_DARK,
            bootstyle=f"{PRIMARY}",
        )
        self.run_button.grid(row=0, column=2, padx=10)

        self.open_button = self._make_action_button(
            button_row,
            "打开生成好的 Excel",
            self.open_output_file,
            OPEN_BG,
            OPEN_FG,
            border="#BFD2F2",
            bootstyle=f"{SECONDARY}-{OUTLINE}",
        )
        self.open_button.grid(row=0, column=3, padx=10)

        link_row = tk.Frame(select_box, bg="#FCF8F3")
        link_row.grid(row=5, column=0, columnspan=6, pady=(0, 14))
        ttk.Button(link_row, text="打开所选月份文件夹", style="App.TButton", command=self.open_selected_month_folder, **self._bs(f"{SECONDARY}-{OUTLINE}")).grid(row=0, column=0, padx=6)
        ttk.Button(link_row, text="打开放文件夹", style="App.TButton", command=self.open_data_dir, **self._bs(f"{SECONDARY}-{OUTLINE}")).grid(row=0, column=1, padx=6)
        ttk.Button(link_row, text="打开运行日志", style="App.TButton", command=self.open_runtime_log, **self._bs(f"{SECONDARY}-{OUTLINE}")).grid(row=0, column=2, padx=6)

        info_box = tk.Frame(card, bg=SURFACE_ALT, highlightbackground=LINE, highlightthickness=1, bd=0)
        info_box.grid(row=6, column=0, sticky="ew")
        info_box.grid_columnconfigure(1, weight=1)
        info_top_line = tk.Frame(info_box, bg="#CBD8E2", height=2)
        info_top_line.grid(row=0, column=0, columnspan=3, sticky="ew")
        tk.Label(info_box, text="放文件位置", bg=SURFACE_ALT, fg=MUTED, font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0, sticky="w", padx=16, pady=(14, 8))
        ttk.Entry(info_box, textvariable=self.data_dir_var).grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=(14, 8), ipady=4)
        ttk.Button(info_box, text="选择文件夹", style="App.TButton", command=self.choose_data_dir, **self._bs(f"{INFO}-{OUTLINE}")).grid(row=1, column=2, sticky="ew", padx=(0, 16), pady=(14, 8))

        tk.Label(info_box, text="结果文件", bg=SURFACE_ALT, fg=MUTED, font=("Microsoft YaHei UI", 10, "bold")).grid(row=2, column=0, sticky="w", padx=16, pady=(0, 14))
        tk.Label(info_box, text="固定生成在程序目录，文件名为：考勤统计结果.xlsx", bg=SURFACE_ALT, fg=INK, font=("Microsoft YaHei UI", 11)).grid(row=2, column=1, columnspan=2, sticky="w", padx=(0, 16), pady=(0, 14))

        self.year_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_selected_month_changed())
        self.month_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_selected_month_changed())
        return card

    def _build_month_card(self, parent: tk.Widget) -> ttk.Frame:
        card = ttk.Frame(parent, style="Panel.TFrame", padding=24)
        card.grid_rowconfigure(2, weight=1)
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=0)

        accent_bar = tk.Frame(card, bg="#5D86C5", height=4)
        accent_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 18))
        self._theme_label(
            card,
            text="已识别月份",
            bootstyle=PRIMARY,
            font=("Microsoft YaHei UI", 18, "bold"),
            foreground=INK,
            background=CARD_BG,
            anchor="w",
        ).grid(row=1, column=0, sticky="w")

        legend_row = tk.Frame(card, bg=CARD_BG)
        legend_row.grid(row=1, column=1, sticky="e")
        for idx, (text, fg) in enumerate((
            ("[已就绪]", SUCCESS),
            ("[已放1个文件]", WARN),
            ("[文件重复]", ERROR),
            ("[未放文件]", MUTED),
        )):
            self._theme_label(
                legend_row,
                text=text,
                bootstyle=f"{SECONDARY}-{OUTLINE}" if text == "[未放文件]" else (f"{BS_SUCCESS}-{OUTLINE}" if text == "[已就绪]" else ("danger-outline" if text == "[文件重复]" else f"{INFO}-{OUTLINE}")),
                font=("Microsoft YaHei UI", 10, "bold"),
                foreground=fg,
                background=CARD_BG,
                padding=(10, 5),
            ).grid(row=0, column=idx, padx=(0, 8))

        ttk.Label(card, text="双击某个月份可以直接打开对应文件夹。状态会区分已就绪、缺1个文件、重复文件和未放文件。", style="Hint.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 16))
        ttk.Checkbutton(card, text="仅显示待处理月份", variable=self.issue_only_var, command=self._on_issue_only_changed).grid(row=2, column=1, sticky="e", pady=(6, 16))

        columns = ("month", "attendance", "leave", "status", "detail")
        tree = ttk.Treeview(card, columns=columns, show="headings", height=10)
        tree.heading("month", text="月份")
        tree.heading("attendance", text="考勤文件")
        tree.heading("leave", text="请假文件")
        tree.heading("status", text="状态")
        tree.heading("detail", text="")
        tree.column("month", width=110, anchor="center")
        tree.column("attendance", width=220, anchor="w")
        tree.column("leave", width=220, anchor="w")
        tree.column("status", width=140, anchor="center")
        tree.column("detail", width=0, minwidth=0, stretch=False)
        tree.grid(row=3, column=0, sticky="nsew")
        tree.tag_configure("ok", foreground=SUCCESS, background="#EDF8F2")
        tree.tag_configure("warn", foreground=ERROR, background="#FCEDEA")
        tree.tag_configure("partial", foreground=WARN, background="#FFF6E8")
        tree.tag_configure("empty", foreground=MUTED, background="#F4F6F8")
        tree.bind("<Double-1>", lambda _event: self.open_tree_selected_month_folder())

        scroll = ttk.Scrollbar(card, orient="vertical", command=tree.yview)
        scroll.grid(row=3, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll.set)
        self.bundle_tree = tree
        self._apply_detail_column_visibility()

        button_row = tk.Frame(card, bg=CARD_BG)
        button_row.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        button_row.grid_columnconfigure(0, weight=1)
        ttk.Button(button_row, text="打开选中月份文件夹", style="App.TButton", command=self.open_tree_selected_month_folder, **self._bs(f"{SECONDARY}-{OUTLINE}")).grid(row=0, column=0, sticky="ew")
        return card

    def log(self, message: str, color: str = INK) -> None:
        self.runtime_logs.append(message)
        if len(self.runtime_logs) > 300:
            self.runtime_logs = self.runtime_logs[-300:]
        self._append_runtime_log(message)

    def _start_runtime_log_session(self) -> None:
        self._rotate_runtime_log_if_needed()
        header = [
            "",
            "=" * 72,
            f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {APP_NAME} {APP_VERSION} session start",
            f"cwd: {Path.cwd()}",
            f"script_root: {_app_root()}",
            "=" * 72,
        ]
        with self.runtime_log_path.open("a", encoding="utf-8") as handle:
            handle.write("\n".join(header) + "\n")

    def _append_runtime_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = message.splitlines() or [""]
        for line in lines:
            self._runtime_log_buffer.append(f"{timestamp} | {line}\n")
        if self._runtime_log_flush_after_id is None and not self._closing:
            self._runtime_log_flush_after_id = self.after(250, self._flush_runtime_log_buffer)

    def _flush_runtime_log_buffer(self) -> None:
        self._runtime_log_flush_after_id = None
        if not self._runtime_log_buffer:
            return
        self.runtime_log_path.parent.mkdir(parents=True, exist_ok=True)
        self._rotate_runtime_log_if_needed()
        with self.runtime_log_path.open("a", encoding="utf-8") as handle:
            handle.writelines(self._runtime_log_buffer)
        self._runtime_log_buffer.clear()

    def _format_scan_duration(self, seconds: float) -> str:
        return f"耗时 {seconds:.2f} 秒"

    def _rotate_runtime_log_if_needed(self) -> None:
        try:
            if not self.runtime_log_path.exists():
                return
            if self.runtime_log_path.stat().st_size < RUNTIME_LOG_MAX_BYTES:
                return
        except OSError:
            return
        for index in range(RUNTIME_LOG_BACKUP_COUNT, 0, -1):
            source = self.runtime_log_path.with_suffix(self.runtime_log_path.suffix + ("" if index == 1 else f".{index - 1}"))
            target = self.runtime_log_path.with_suffix(self.runtime_log_path.suffix + f".{index}")
            if source.exists():
                try:
                    if target.exists():
                        target.unlink()
                    source.replace(target)
                except OSError:
                    continue

    def set_notice(self, title: str, desc: str, level: str = "info") -> None:
        if BOOTSTRAP_ENABLED:
            if level == "success":
                frame_style = self._normalize_bootstyle(f"{BS_SUCCESS}-{OUTLINE}", "frame")
                title_style = BS_SUCCESS
                desc_style = SECONDARY
            elif level == "error":
                frame_style = self._normalize_bootstyle("danger-outline", "frame")
                title_style = "danger"
                desc_style = SECONDARY
            else:
                frame_style = self._normalize_bootstyle(f"{INFO}-{OUTLINE}", "frame")
                title_style = INFO
                desc_style = SECONDARY

            self.notice_frame.configure(**self._bs(frame_style))
            self.notice_title_label.configure(text=title, **self._bs(title_style))
            self.notice_desc_label.configure(text=desc, **self._bs(desc_style))
            return

        if level == "success":
            bg = NOTICE_BG_SUCCESS
            border = "#A9D0AE"
            fg = SUCCESS
        elif level == "error":
            bg = NOTICE_BG_ERROR
            border = "#E0B0A8"
            fg = ERROR
        else:
            bg = NOTICE_BG_INFO
            border = "#E3C6A0"
            fg = WARN

        self.notice_frame.configure(bg=bg, highlightbackground=border)
        self.notice_title_label.configure(text=title, bg=bg, fg=fg)
        self.notice_desc_label.configure(text=desc, bg=bg, fg=INK if level != "info" else MUTED)

    def clear_log(self) -> None:
        self.runtime_logs.clear()

    def choose_data_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.data_dir_var.get() or str(_default_data_dir()))
        if selected:
            self.data_dir_var.set(selected)
            self._scan_cache_key = None
            self._scan_cache_bundles = []
            self._scan_cache_error = ""
            self._last_full_scan_key = None
            self._last_scan_used_cache = False
            self._last_scan_feedback = ""
            self._folder_cache_key = None
            self._folder_cache_inspections = []
            self.refresh_folder_overview()

    def open_data_dir(self) -> None:
        data_dir = Path(self.data_dir_var.get())
        data_dir.mkdir(parents=True, exist_ok=True)
        _open_path(data_dir)

    def open_runtime_log(self) -> None:
        self.runtime_log_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.runtime_log_path.exists():
            self._start_runtime_log_session()
        self._flush_runtime_log_buffer()
        _open_path(self.runtime_log_path)

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
        existing_annual_files = _find_matching_files(data_dir, FILE_KIND_CONFIG["annual"]["patterns"])
        if existing_annual_files and not messagebox.askyesno("确认覆盖", "当前年假总数表已存在，是否用新文件覆盖？"):
            return
        for existing in existing_annual_files:
            if existing.exists():
                existing.unlink()
        shutil.copy2(source_path, target_path)
        self._refresh_annual_info()
        self.log(f"已上传当前年假表 -> {target_path}", SUCCESS)
        self.set_notice("当前年假表已更新。", "只有员工年假信息发生变化时，才需要再次上传。", "success")
        self._scan_cache_key = None
        self._folder_cache_key = None
        self.refresh_folder_overview()

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

    def _populate_bundles(self, bundles: list[MonthlySourceBundle], full_scan: bool = True) -> None:
        folder_rows, issues = self._inspect_month_folders()
        self.current_bundles = list(bundles)
        self.month_issue_details = {
            str(row["month"]): str(row.get("detail", ""))
            for row in folder_rows
            if row.get("detail")
        }
        display_rows = [row for row in folder_rows if (not self.issue_only_var.get() or row.get("detail"))]
        self._refresh_year_values()
        for item in self.bundle_tree.get_children():
            self.bundle_tree.delete(item)

        if display_rows:
            for row in display_rows:
                self.bundle_tree.insert(
                    "",
                    "end",
                    values=(
                        row["month"],
                        row["attendance"],
                        row["leave"],
                        row["status"],
                        row["detail"],
                    ),
                    tags=(str(row["tag"]),),
                )
        elif folder_rows and self.issue_only_var.get():
            pass
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
                        "",
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

        ready_folder_count = sum(1 for row in folder_rows if row.get("tag") == "ok")

        if issues and bundles:
            self.status_var.set(f"已识别 {len(bundles)} 个月份，另有部分月份待处理。")
            self.summary_var.set(f"当前可统计 {len(bundles)} 个月份，其余月份待处理")
            self.set_notice(
                "部分月份还需要处理。",
                "已就绪的月份仍可先生成；显示 [需处理] 的月份可以稍后补齐。",
                "info",
            )
        elif issues:
            self.status_var.set("有月份文件夹需要处理，请先补齐或去重。")
            self.summary_var.set("当前还没有完整月份可统计")
            self.set_notice(
                "有月份文件夹还需要处理。",
                "请看月份列表最右边“状态”这一列。哪里显示“缺少”或“重复”，就先处理哪里。",
                "error",
            )
        elif not full_scan and ready_folder_count:
            self.status_var.set(f"当前年份发现 {ready_folder_count} 个月份文件已就绪。")
            self.summary_var.set("已完成轻量检查，点击“检查当前年份文件”可做完整识别")
            self.set_notice(
                "已完成轻量检查。",
                f"当前年份发现 {ready_folder_count} 个已就绪月份。需要完整识别时，请点击“检查当前年份文件”。",
                "info",
            )
        elif bundles:
            if not selected_folder or not (Path(self.data_dir_var.get()) / selected_folder).exists():
                parsed = _parse_month_input(self._get_next_month_folder_name())
                if parsed:
                    self._set_selected_month(*parsed)
            self.status_var.set(f"已找到 {len(bundles)} 个月份，年度合计会自动一起计算。")
            self.summary_var.set(f"已识别 {len(bundles)} 个月份，可以直接生成")
            self.set_notice(
                "检查通过，可以直接生成结果文件。",
                f"当前共识别到 {len(bundles)} 个月份。点击中间的“生成结果文件”即可。",
                "success",
            )
        else:
            self._on_selected_month_changed()
            self.status_var.set("没有找到可统计的月份。")
            self.summary_var.set("请先上传当前年假表，并上传某个月份的 2 个表")
            self.set_notice(
                "还没有找到可以统计的月份。",
                "请先上传当前年假表，再选好年份和月份，点击中间的“上传所选月份2个表”。",
                "info",
            )

    def _has_folder_issues(self) -> bool:
        _, issues = self._inspect_month_folders()
        return bool(issues)

    def _selected_year_folder_issues(self) -> list[str]:
        try:
            year = self._get_selected_year()
        except ValueError:
            return []
        rows, _ = self._inspect_month_folders()
        return [
            f"{row['month']}: {row['detail']}"
            for row in rows
            if str(row["month"]).startswith(f"{year}-") and row.get("detail")
        ]

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
        return _find_matching_files(month_dir, FILE_KIND_CONFIG[kind]["patterns"])

    def _find_existing_month_file(self, month_dir: Path, kind: str) -> Path | None:
        files = self._find_existing_month_files(month_dir, kind)
        return files[0] if files else None

    def _inspect_month_folders(self) -> tuple[list[dict[str, object]], list[str]]:
        base_dir = Path(self.data_dir_var.get())
        if not base_dir.exists():
            return [], ["放文件的文件夹不存在。"]

        rows: list[dict[str, object]] = []
        issues: list[str] = []
        target_year = None
        try:
            target_year = self._get_selected_year()
        except ValueError:
            pass

        inspections = self._get_folder_inspections(target_year)
        for item in inspections:
            attendance_count = len(item.attendance_files)
            leave_count = len(item.leave_files)
            if item.ready:
                status_text = "[已就绪]"
                detail_text = ""
                tag = "ok"
            elif attendance_count == 0 and leave_count == 0:
                status_text = "[未放文件]"
                detail_text = item.detail or "还没有识别到考勤或请假文件"
                tag = "empty"
            elif "重复" in item.detail:
                status_text = "[文件重复]"
                detail_text = item.detail
                tag = "warn"
            else:
                status_text = "[已放1个文件]"
                detail_text = item.detail or "该月文件还不完整"
                tag = "partial"
            if detail_text:
                issues.append(f"{item.folder_name}: {detail_text}")

            rows.append(
                {
                    "month": item.folder_name,
                    "attendance": item.attendance_files[0].name if len(item.attendance_files) == 1 else ("未放" if not item.attendance_files else f"{len(item.attendance_files)}个文件"),
                    "leave": item.leave_files[0].name if len(item.leave_files) == 1 else ("未放" if not item.leave_files else f"{len(item.leave_files)}个文件"),
                    "status": status_text,
                    "detail": detail_text,
                    "tag": tag,
                }
            )

        return rows, issues

    def _ensure_selected_month_dir(self) -> Path:
        base_dir = Path(self.data_dir_var.get())
        base_dir.mkdir(parents=True, exist_ok=True)
        folder_name = self._selected_month_folder_name()
        month_dir = base_dir / folder_name
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
        month_name = str(values[0])
        return self.month_issue_details.get(month_name, "")

    def open_selected_month_folder(self) -> None:
        month_dir = self._ensure_selected_month_dir_or_warn()
        if month_dir is None:
            return
        guide_file = month_dir / "请把这2个文件放到这里.txt"
        if not guide_file.exists():
            guide_file.write_text(
                "\n".join(
                    [
                        "请把下面 2 个文件放到这个文件夹里：",
                        "",
                        "1. 考勤打卡记录表.xls",
                        "2. 请假记录表.xls",
                        "",
                        "当前年假总数表不放在这里，请在程序里单独上传。",
                        "",
                        "也可以直接在程序界面里点击上传按钮。",
                    ]
                ),
                encoding="utf-8",
            )
        self.log(f"已打开月份文件夹：{month_dir}", SUCCESS)
        self.set_notice(
            f"已准备好 {month_dir.name} 文件夹。",
            "你可以直接把 Excel 拖进去，也可以回到程序点击中间的上传按钮。",
            "info",
        )
        _open_path(month_dir)

    def open_tree_selected_month_folder(self) -> None:
        month_dir = self._selected_tree_month_dir()
        if month_dir is None:
            messagebox.showinfo("提示", "请先在月份列表里点中一个月份。")
            return
        month_dir.mkdir(parents=True, exist_ok=True)
        parsed = _parse_month_input(month_dir.name)
        if parsed:
            self._set_selected_month(*parsed)
        self.log(f"已打开月份文件夹：{month_dir}", SUCCESS)
        issue_detail = self._selected_tree_issue_detail()
        if issue_detail:
            messagebox.showinfo("该月份需要处理", f"{month_dir.name} 当前问题：\n{issue_detail}")
        self.set_notice(
            f"已打开 {month_dir.name} 文件夹。",
            "如果这个月份缺文件，可以直接拖进去，或者点击中间的上传按钮补齐。",
            "info",
        )
        _open_path(month_dir)

    def _confirm_selected_source(self, kind: str, source_path: Path) -> bool:
        expected_label = FILE_KIND_CONFIG[kind]["label"]
        if expected_label in source_path.stem:
            return True
        return messagebox.askyesno(
            "确认文件",
            f"你选择的文件名是：{source_path.name}\n\n"
            f"这个文件名里没有明显看到“{expected_label}”。\n"
            "如果选错文件，后面统计会出错。\n\n"
            "要继续上传吗？",
        )

    def _copy_monthly_file(self, month_dir: Path, kind: str, source_path: Path) -> Path:
        suffix = source_path.suffix.lower()
        if suffix not in {".xls", ".xlsx"}:
            raise ValueError("只支持 Excel 文件：.xls 或 .xlsx")

        existing_files = self._find_existing_month_files(month_dir, kind)
        if existing_files:
            overwrite = messagebox.askyesno(
                "确认覆盖",
                f"{month_dir.name} 文件夹里已经有“{FILE_KIND_CONFIG[kind]['label']}”。\n是否用新文件覆盖？",
            )
            if not overwrite:
                raise RuntimeError("用户取消覆盖")
            for existing in existing_files:
                existing.unlink()

        target_path = month_dir / f"{FILE_KIND_CONFIG[kind]['label']}{suffix}"
        shutil.copy2(source_path, target_path)
        return target_path

    def upload_monthly_file(self, kind: str) -> None:
        month_dir = self._ensure_selected_month_dir_or_warn()
        if month_dir is None:
            return
        initial_dir = str(month_dir if month_dir.exists() else Path.home())
        label = FILE_KIND_CONFIG[kind]["label"]
        selected = filedialog.askopenfilename(
            title=f"请选择{label}",
            initialdir=initial_dir,
            filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")],
        )
        if not selected:
            return

        source_path = Path(selected)
        if not self._confirm_selected_source(kind, source_path):
            return
        try:
            target_path = self._copy_monthly_file(month_dir, kind, source_path)
        except RuntimeError:
            return
        except Exception as exc:
            messagebox.showerror("上传失败", str(exc))
            return

        self.log(f"已上传 {label} -> {target_path}", SUCCESS)
        self.summary_var.set(f"已上传：{label}")
        self.status_var.set(f"{month_dir.name} 文件夹已更新，请继续上传剩余文件，或直接生成结果。")
        self.set_notice(
            f"{label} 已上传成功。",
            f"文件已放入 {month_dir.name}。如果本月 2 个文件都放好了，下一步直接点击“生成结果文件”。",
            "success",
        )
        self._scan_cache_key = None
        self._folder_cache_key = None
        self.refresh_folder_overview()

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
            initial_dir = str(month_dir if month_dir.exists() else Path.home())
            selected = filedialog.askopenfilename(
                title=f"请选择{label}",
                initialdir=initial_dir,
                filetypes=[("Excel 文件", "*.xls *.xlsx"), ("所有文件", "*.*")],
            )
            if not selected:
                if uploaded == 0:
                    self.set_notice(
                        "还没有上传文件。",
                        "你可以重新点击中间的“上传所选月份2个表”。",
                        "info",
                    )
                else:
                    self.set_notice(
                        "已上传部分文件。",
                        "请继续上传剩余文件，2 个文件齐全后直接点“生成结果文件”。",
                        "info",
                    )
                return

            source_path = Path(selected)
            if not self._confirm_selected_source(kind, source_path):
                self.set_notice(
                    "已暂停上传。",
                    "你取消了当前文件的确认。可以重新点击中间的“上传所选月份2个表”。",
                    "info",
                )
                return
            try:
                target_path = self._copy_monthly_file(month_dir, kind, source_path)
            except RuntimeError:
                continue
            except Exception as exc:
                messagebox.showerror("上传失败", str(exc))
                return
            uploaded += 1
            self.log(f"已上传 {label} -> {target_path}", SUCCESS)

        self.summary_var.set(f"{month_dir.name} 的 2 个文件已上传完成")
        self.status_var.set(f"{month_dir.name} 文件夹文件已齐，请点击“生成结果文件”。")
        self.set_notice(
            f"{month_dir.name} 的 2 个文件已上传完成。",
            "现在可以直接点击中间的“生成结果文件”。程序会自动检查后再生成。",
            "success",
        )
        self._scan_cache_key = None
        self._folder_cache_key = None
        self.refresh_folder_overview()

    def create_next_month_folder(self) -> None:
        base_dir = Path(self.data_dir_var.get())
        base_dir.mkdir(parents=True, exist_ok=True)

        folder_name = self._get_next_month_folder_name()
        parsed = _parse_month_input(folder_name)
        if parsed:
            self._set_selected_month(*parsed)
        month_dir = base_dir / folder_name
        month_dir.mkdir(parents=True, exist_ok=True)

        guide_file = month_dir / "请把这2个文件放到这里.txt"
        guide_file.write_text(
            "\n".join(
                [
                    "请把下面 2 个文件放到这个文件夹里：",
                    "",
                    "1. 考勤打卡记录表.xls",
                    "2. 请假记录表.xls",
                    "",
                    "当前年假总数表请在程序中单独上传。",
                    "",
                    "注意：",
                    "- 文件名请保持不变",
                    "- 放好后，回到程序里点击“生成结果文件”",
                ]
            ),
            encoding="utf-8",
        )

        self.log(f"已创建月份文件夹：{month_dir}", SUCCESS)
        self.log("请把 2 个原始 Excel 文件放进去。", MUTED)
        self.summary_var.set(f"已创建 {folder_name} 文件夹")
        self.status_var.set("新文件夹已创建，请把该月的 2 个文件放进去。")
        self.set_notice(
            f"已经创建好 {folder_name} 文件夹。",
            "现在请把这个月的 2 个 Excel 文件放进去。年假表只在变化时单独更新。",
            "info",
        )
        self._scan_cache_key = None
        self._folder_cache_key = None
        self.refresh_folder_overview()
        _open_path(month_dir)

    def scan_bundles(self, preserve_log: bool = False) -> None:
        if self._closing:
            return
        started_at = time.perf_counter()
        data_dir = Path(self.data_dir_var.get())
        selected_year = None
        try:
            selected_year = self._get_selected_year()
        except ValueError:
            pass
        cache_key = self._build_scan_cache_key(data_dir, selected_year)
        if self._scan_cache_key == cache_key:
            self._last_scan_duration_seconds = time.perf_counter() - started_at
            self._last_scan_used_cache = True
            self._last_scan_feedback = f"自上次完整检查后未发现文件变化，本次直接复用结果。{self._format_scan_duration(self._last_scan_duration_seconds)}"
            self.last_scan_error_message = self._scan_cache_error
            self._refresh_annual_info()
            self._populate_bundles(self._scan_cache_bundles, full_scan=True)
            if not preserve_log:
                self.clear_log()
                self.log(f"放文件的文件夹：{data_dir}", MUTED)
                self.log("检查结果来自缓存，本次未重复读取 Excel。", SUCCESS)
                try:
                    annual_summary = get_current_annual_leave_summary(str(data_dir))
                    self.log(f"当前年假表：{annual_summary['file_name']} | 员工数={annual_summary['employee_count']}", SUCCESS)
                except Exception:
                    self.log("当前年假表：未上传", ERROR)
                for bundle in self._scan_cache_bundles:
                    self.log(
                        f"找到月份 {bundle.year}-{bundle.month:02d} | "
                        f"{bundle.attendance_file.name} | {bundle.leave_file.name}",
                        SUCCESS,
                    )
            return
        previous_full_scan_key = self._last_full_scan_key
        file_changed_since_last_scan = previous_full_scan_key is not None and previous_full_scan_key != cache_key
        try:
            bundles = discover_monthly_source_bundles(str(data_dir), target_year=selected_year, relaxed=True)
        except Exception as exc:
            friendly_message = _friendly_scan_error(exc)
            self._last_scan_duration_seconds = time.perf_counter() - started_at
            self._last_scan_used_cache = False
            prefix = "检测到文件变化，已重新检查。" if file_changed_since_last_scan else "已完成完整检查。"
            self._last_scan_feedback = f"{prefix}{self._format_scan_duration(self._last_scan_duration_seconds)}"
            self._scan_cache_key = cache_key
            self._scan_cache_bundles = []
            self._scan_cache_error = friendly_message
            self._last_full_scan_key = cache_key
            self.last_scan_error_message = friendly_message
            self._populate_bundles([], full_scan=True)
            self._refresh_annual_info()
            self.status_var.set("检查失败，请先处理文件夹中的问题。")
            self.set_notice(
                "文件检查失败。",
                friendly_message,
                "error",
            )
            if not preserve_log:
                self.clear_log()
            self.log(f"检查失败：{friendly_message}", ERROR)
            self.log(str(exc), MUTED)
            return

        if not bundles:
            ready_rows = [row for row in self._inspect_month_folders()[0] if row.get("tag") == "ok"]
            if ready_rows:
                try:
                    discover_monthly_source_bundles(str(data_dir), target_year=selected_year, relaxed=False)
                except Exception as exc:
                    friendly_message = _friendly_scan_error(exc)
                else:
                    friendly_message = "已检测到月份文件，但未能识别出可统计的月份。请检查考勤表里的年月、请假表配套关系和文件内容。"
                self._last_scan_duration_seconds = time.perf_counter() - started_at
                self._last_scan_used_cache = False
                prefix = "检测到文件变化，已重新检查。" if file_changed_since_last_scan else "已完成完整检查。"
                self._last_scan_feedback = f"{prefix}{self._format_scan_duration(self._last_scan_duration_seconds)}"
                self._scan_cache_key = cache_key
                self._scan_cache_bundles = []
                self._scan_cache_error = friendly_message
                self._last_full_scan_key = cache_key
                self.last_scan_error_message = friendly_message
                self._refresh_annual_info()
                self._populate_bundles([], full_scan=True)
                self.status_var.set("已发现月份文件，但当前还不能识别成可统计月份。")
                self.set_notice("当前文件还不能直接统计。", friendly_message, "error")
                if not preserve_log:
                    self.clear_log()
                self.log(f"检查结果：{friendly_message}", ERROR)
                return

        self._scan_cache_key = cache_key
        self._scan_cache_bundles = list(bundles)
        self._scan_cache_error = ""
        self._last_full_scan_key = cache_key
        self._last_scan_duration_seconds = time.perf_counter() - started_at
        self._last_scan_used_cache = False
        prefix = "检测到文件变化，已重新检查。" if file_changed_since_last_scan else "已完成完整检查。"
        self._last_scan_feedback = f"{prefix}{self._format_scan_duration(self._last_scan_duration_seconds)}"
        self.last_scan_error_message = ""

        if not preserve_log:
            self.clear_log()
            self.log(f"放文件的文件夹：{data_dir}", MUTED)
            try:
                annual_summary = get_current_annual_leave_summary(str(data_dir))
                self.log(f"当前年假表：{annual_summary['file_name']} | 员工数={annual_summary['employee_count']}", SUCCESS)
            except Exception:
                self.log("当前年假表：未上传", ERROR)
            for bundle in bundles:
                self.log(
                    f"找到月份 {bundle.year}-{bundle.month:02d} | "
                    f"{bundle.attendance_file.name} | {bundle.leave_file.name}",
                    SUCCESS,
                )
        self._refresh_annual_info()
        self._populate_bundles(bundles, full_scan=True)

    def run_report(self) -> None:
        if self._closing:
            return
        if self.run_thread and self.run_thread.is_alive():
            return

        data_dir = Path(self.data_dir_var.get())
        selected_year = self._get_selected_year()
        cache_key = self._build_scan_cache_key(data_dir, selected_year)
        generation_cache_reused = False
        if self._scan_cache_key == cache_key:
            generation_cache_reused = True
            self._last_scan_used_cache = True
            self.last_scan_error_message = self._scan_cache_error
            self.current_bundles = list(self._scan_cache_bundles)
        else:
            self.scan_bundles()
            generation_cache_reused = self._last_scan_used_cache

        if not self.current_bundles:
            issue_lines = self._selected_year_folder_issues()
            detail = self.last_scan_error_message or "请先上传当前年假表，并上传至少一个月份的 2 个表。"
            if issue_lines:
                detail = "当前选中年份还没有完整可统计的月份。\n\n" + "\n".join(issue_lines[:5])
            self.summary_var.set("还没有可统计的月份")
            self.status_var.set("请先上传当前年假表，并上传该月 2 个表。")
            self.set_notice(
                "还没有找到可统计的月份。",
                detail,
                "error",
            )
            messagebox.showinfo("无法生成", f"还没有找到可统计的月份。\n\n{detail}")
            return

        output_file = Path(self.output_file_var.get())
        output_file.parent.mkdir(parents=True, exist_ok=True)

        self.clear_log()
        self.log(f"开始生成统计。放文件的文件夹：{data_dir}", MUTED)
        self.log(f"结果会保存到：{output_file}", MUTED)
        if generation_cache_reused:
            self.log("生成前缓存有效，未重复执行完整检查。", SUCCESS)
        self.upload_button.configure(state="disabled")
        self.check_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.summary_var.set("正在生成，请稍等 10-30 秒")
        self.set_notice(
            "正在生成结果文件，请稍等。",
            "生成过程中不要重复点击按钮。完成后会自动提示你打开结果文件。",
            "info",
        )

        def worker() -> None:
            try:
                summary = generate_report(
                    str(data_dir),
                    str(output_file),
                    logger=lambda msg: self.log_queue.put(("log", msg)),
                    target_year=selected_year,
                    relaxed=True,
                )
                self.log_queue.put(("done", summary))
            except Exception:
                self.log_queue.put(("error", traceback.format_exc()))

        self.run_thread = threading.Thread(target=worker, daemon=True)
        self.run_thread.start()

    def _poll_log_queue(self) -> None:
        if self._closing:
            return
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "log":
                    self.log(str(payload), INK)
                elif kind == "done":
                    summary = payload
                    assert isinstance(summary, ReportSummary)
                    self.log("统计完成。", SUCCESS)
                    self.summary_var.set(f"已经生成完成，共统计 {len(summary.monthly_results)} 个月份")
                    self.status_var.set(f"结果文件已生成：{summary.output_file}")
                    self.set_notice(
                        "结果文件已生成完成。",
                        f"保存位置：{summary.output_file}",
                        "success",
                    )
                    self._show_generation_success_dialog(summary)
                    self.scan_bundles(preserve_log=True)
                elif kind == "error":
                    self.log("生成失败：", ERROR)
                    self.log(str(payload), ERROR)
                    self.summary_var.set("生成失败，请检查放进去的文件")
                    self.status_var.set("生成失败，请检查文件是否完整、文件名是否正确。")
                    self.set_notice(
                        "生成失败。",
                        "请检查当前年假表和各月份的 2 个文件是否完整、文件名是否正确。",
                        "error",
                    )
                    messagebox.showerror("生成失败", "统计过程中出现错误。\n请检查文件是否完整、文件名是否正确。")
                if kind in {"done", "error"}:
                    self.upload_button.configure(state="normal")
                    self.check_button.configure(state="normal")
                    self.run_button.configure(state="normal")
        except queue.Empty:
            pass
        if not self._closing:
            self._poll_after_id = self.after(120, self._poll_log_queue)


def main() -> None:
    try:
        app = AttendanceGui()
        app.mainloop()
    except Exception as exc:
        log_file = _write_startup_diagnostic(exc)
        print("GUI startup failed.", file=sys.stderr)
        print(f"Diagnostic log: {log_file}", file=sys.stderr)
        print(str(exc), file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
