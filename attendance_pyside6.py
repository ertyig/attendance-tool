#!/usr/bin/env python3
"""PySide6 版本的考勤统计桌面界面。"""

from __future__ import annotations

import os
import platform
import queue
import shutil
import subprocess
import sys
import threading
import time
import traceback
from datetime import date, datetime
from fnmatch import fnmatch
from pathlib import Path

if sys.platform == "darwin":
    os.environ.pop("SYSTEM_VERSION_COMPAT", None)

from PySide6 import QtCore, QtGui, QtWidgets

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

APP_NAME = "考勤统计助手"
APP_VERSION = "v1.0.0"
APP_BG = "#ECF2F6"
CARD_BG = "#FFFFFF"
CARD_ALT = "#F6F9FC"
HEADER_BG = "#123847"
TEXT = "#1E2A34"
TEXT_MUTED = "#617383"
PRIMARY = "#245B7A"
PRIMARY_DARK = "#1B4760"
SUCCESS = "#2D7A58"
SUCCESS_SOFT = "#EAF7F0"
WARNING = "#C56A3D"
WARNING_SOFT = "#FFF2EA"
DANGER = "#C95A4D"
DANGER_SOFT = "#FCEDEA"
INFO_SOFT = "#EEF5FA"
BORDER = "#D7E1E8"
RUNTIME_LOG_MAX_BYTES = 2 * 1024 * 1024
RUNTIME_LOG_BACKUP_COUNT = 1

FILE_KIND_CONFIG = {
    "attendance": {
        "label": "考勤打卡记录表",
        "patterns": ["*考勤*.xls", "*考勤*.xlsx", "a.xls", "a.xlsx"],
    },
    "leave": {
        "label": "请假记录表",
        "patterns": ["*请假*.xls", "*请假*.xlsx", "b.xls", "b.xlsx"],
    },
    "annual": {
        "label": "员工年假总数表",
        "patterns": ["*年假*.xls", "*年假*.xlsx", "c.xls", "c.xlsx"],
    },
}
CURRENT_ANNUAL_TARGET_STEM = "当前员工年假总数表"


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


def _app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _default_data_dir() -> Path:
    return (_app_root() / DATA_DIR).resolve()


def _default_output_file() -> Path:
    return (_app_root() / OUTPUT_FILE).resolve()


def _runtime_log_file() -> Path:
    return (_app_root() / "attendance_pyside6_runtime.log").resolve()


def _startup_log_file() -> Path:
    return (_app_root() / "attendance_pyside6_startup.log").resolve()


def _detailed_usage_file() -> Path:
    return (_app_root() / "使用说明-给同事看.txt").resolve()


def _fallback_usage_file() -> Path:
    return (_app_root() / "data" / "月度文件" / "使用说明.txt").resolve()


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


def _parse_month_input(text: str) -> tuple[int, int] | None:
    text = text.strip()
    candidates = [
        text,
        text.replace("/", "-"),
        text.replace("_", "-"),
        text.replace("年", "-").replace("月", ""),
    ]
    for candidate in candidates:
        parts = candidate.split("-")
        if len(parts) != 2:
            continue
        year, month = parts
        if year.isdigit() and month.isdigit():
            value = int(month)
            if 1 <= value <= 12:
                return int(year), value
    return None


class WorkerSignals(QtCore.QObject):
    message = QtCore.Signal(str)
    done = QtCore.Signal(object)
    error = QtCore.Signal(str)


class ReportWorker(QtCore.QRunnable):
    def __init__(self, data_dir: str, output_file: str, target_year: int) -> None:
        super().__init__()
        self.signals = WorkerSignals()
        self.data_dir = data_dir
        self.output_file = output_file
        self.target_year = target_year

    @QtCore.Slot()
    def run(self) -> None:
        try:
            summary = generate_report(
                self.data_dir,
                self.output_file,
                logger=lambda msg: self.signals.message.emit(str(msg)),
                target_year=self.target_year,
                relaxed=True,
            )
            self.signals.done.emit(summary)
        except Exception:
            self.signals.error.emit(traceback.format_exc())


class AttendancePySide6Window(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION} | PySide6")
        self.resize(1260, 940)
        self.setMinimumSize(1120, 860)

        self.data_dir = _default_data_dir()
        self.output_file = _default_output_file()
        self.runtime_log_path = _runtime_log_file()
        self.current_bundles: list[MonthlySourceBundle] = []
        self.month_issue_details: dict[str, str] = {}
        self.last_scan_error_message = ""
        self.runtime_logs: list[str] = []
        self._runtime_log_buffer: list[str] = []
        self._scan_cache_key: object | None = None
        self._scan_cache_bundles: list[MonthlySourceBundle] = []
        self._scan_cache_error = ""
        self._last_full_scan_key: object | None = None
        self._last_scan_used_cache = False
        self._last_scan_feedback = ""
        self._folder_cache_key: object | None = None
        self._folder_cache_inspections: list[MonthFolderInspection] = []
        self._last_scan_duration_seconds = 0.0
        self._closing = False
        self._thread_pool = QtCore.QThreadPool.globalInstance()
        self._current_worker: ReportWorker | None = None

        self._build_ui()
        self._apply_styles()
        self._start_runtime_log_session()
        self._refresh_annual_info()
        self._refresh_year_values()
        self._on_selected_month_changed()
        QtCore.QTimer.singleShot(200, self.refresh_folder_overview)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # noqa: N802
        self._closing = True
        self._flush_runtime_log_buffer()
        super().closeEvent(event)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            f"""
            QMainWindow {{ background: {APP_BG}; }}
            QWidget#header {{ background: {HEADER_BG}; }}
            QWidget#card {{ background: {CARD_BG}; border: 1px solid {BORDER}; border-radius: 16px; }}
            QWidget#softCard {{ background: {CARD_ALT}; border: 1px solid {BORDER}; border-radius: 14px; }}
            QLabel#title {{ color: #F7FBFD; font-size: 30px; font-weight: 700; }}
            QLabel#version {{ color: #D4E4EA; font-size: 12px; }}
            QLabel#sectionTitle {{ color: {TEXT}; font-size: 18px; font-weight: 700; }}
            QLabel#muted {{ color: {TEXT_MUTED}; font-size: 12px; }}
            QLabel#summary {{ color: {TEXT}; font-size: 18px; font-weight: 700; }}
            QLabel#status {{ color: {TEXT_MUTED}; font-size: 12px; }}
            QLabel#stateOk {{ color: {SUCCESS}; font-weight: 700; }}
            QLabel#stateBad {{ color: {DANGER}; font-weight: 700; }}
            QLabel#noticeTitle {{ color: {PRIMARY}; font-size: 14px; font-weight: 700; }}
            QLabel#noticeDesc {{ color: {TEXT_MUTED}; font-size: 11px; }}
            QPushButton {{
                border-radius: 10px; padding: 10px 16px; font-size: 13px; font-weight: 600;
                border: 1px solid {BORDER}; background: #E7EEF4; color: #41586F;
            }}
            QPushButton:hover {{ background: #DDE7EF; }}
            QPushButton#primaryAction {{ background: {PRIMARY}; color: white; border: none; }}
            QPushButton#primaryAction:hover {{ background: {PRIMARY_DARK}; }}
            QPushButton#successAction {{ background: {SUCCESS}; color: white; border: none; }}
            QPushButton#successAction:hover {{ background: #25664B; }}
            QPushButton#ghostAction {{ background: #EFF4F8; color: #4F657B; }}
            QPushButton#ghostAction:hover {{ background: #E3EBF1; }}
            QComboBox {{ background: white; border: 1px solid {BORDER}; border-radius: 10px; padding: 8px 12px; min-height: 18px; }}
            QTableWidget {{ background: white; border: 1px solid {BORDER}; border-radius: 12px; gridline-color: {BORDER}; }}
            QHeaderView::section {{ background: #F3F7FA; padding: 8px; border: none; border-bottom: 1px solid {BORDER}; font-weight: 700; }}
            QCheckBox {{ color: {TEXT_MUTED}; font-size: 12px; }}
            QTextEdit {{ background: #FBFCFE; border: 1px solid {BORDER}; border-radius: 12px; padding: 8px; color: {TEXT}; }}
            """
        )

    def _build_ui(self) -> None:
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(16)

        header = QtWidgets.QWidget(objectName="header")
        header.setFixedHeight(100)
        header_layout = QtWidgets.QVBoxLayout(header)
        header_layout.setContentsMargins(24, 18, 24, 18)
        title = QtWidgets.QLabel(APP_NAME, objectName="title")
        version = QtWidgets.QLabel(f"PySide6 版本  {APP_VERSION}", objectName="version")
        header_layout.addWidget(title)
        header_layout.addWidget(version)
        header_layout.addStretch(1)
        root.addWidget(header)

        notice = QtWidgets.QWidget(objectName="softCard")
        notice_layout = QtWidgets.QVBoxLayout(notice)
        notice_layout.setContentsMargins(16, 12, 16, 12)
        self.notice_title = QtWidgets.QLabel("请按顺序操作：先上传当前年假表，再选年月上传本月两个表，最后生成结果文件。", objectName="noticeTitle")
        self.notice_desc = QtWidgets.QLabel("年假表通常不用每月都传，只有员工年假信息发生变化时再更新。", objectName="noticeDesc")
        notice_layout.addWidget(self.notice_title)
        notice_layout.addWidget(self.notice_desc)
        root.addWidget(notice)

        self.summary_label = QtWidgets.QLabel("先上传当前年假表，再选择年月上传该月 2 个表", objectName="summary")
        self.status_label = QtWidgets.QLabel("准备就绪。", objectName="status")
        root.addWidget(self.summary_label)
        root.addWidget(self.status_label)

        body = QtWidgets.QHBoxLayout()
        body.setSpacing(16)
        root.addLayout(body, 1)

        left = QtWidgets.QVBoxLayout()
        left.setSpacing(16)
        body.addLayout(left, 3)

        right = QtWidgets.QVBoxLayout()
        right.setSpacing(16)
        body.addLayout(right, 2)

        left.addWidget(self._build_action_card())
        left.addWidget(self._build_month_card(), 1)
        right.addWidget(self._build_log_card(), 1)

    def _build_action_card(self) -> QtWidgets.QWidget:
        card = QtWidgets.QWidget(objectName="card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        section = QtWidgets.QLabel("本月操作区", objectName="sectionTitle")
        hint = QtWidgets.QLabel("先维护当前年假表，再选年月上传本月 2 个表，最后生成结果。", objectName="muted")
        layout.addWidget(section)
        layout.addWidget(hint)

        annual = QtWidgets.QWidget(objectName="softCard")
        annual_layout = QtWidgets.QGridLayout(annual)
        annual_layout.setContentsMargins(16, 12, 16, 12)
        annual_layout.setHorizontalSpacing(10)
        annual_layout.setVerticalSpacing(8)
        annual_layout.addWidget(QtWidgets.QLabel("当前年假表", objectName="sectionTitle"), 0, 0, 1, 2)
        self.annual_info_label = QtWidgets.QLabel("当前年假表：未上传")
        self.annual_info_label.setWordWrap(True)
        annual_layout.addWidget(QtWidgets.QLabel("年假总数表只有在员工年假信息变化时才需要更新。", objectName="muted"), 1, 0, 1, 2)
        annual_layout.addWidget(self.annual_info_label, 2, 0, 1, 2)
        btn_upload_annual = QtWidgets.QPushButton("上传当前年假表")
        btn_upload_annual.clicked.connect(self.upload_current_annual_file)
        btn_open_annual = QtWidgets.QPushButton("打开当前年假表")
        btn_open_annual.setObjectName("ghostAction")
        btn_open_annual.clicked.connect(self.open_current_annual_file)
        annual_layout.addWidget(btn_upload_annual, 0, 2)
        annual_layout.addWidget(btn_open_annual, 0, 3)
        layout.addWidget(annual)

        month_box = QtWidgets.QWidget(objectName="softCard")
        month_layout = QtWidgets.QVBoxLayout(month_box)
        month_layout.setContentsMargins(16, 12, 16, 12)
        month_layout.setSpacing(10)
        month_layout.addWidget(QtWidgets.QLabel("统计月份", objectName="sectionTitle"))
        month_layout.addWidget(QtWidgets.QLabel("选择的年月就是本次上传和统计的目标月份。", objectName="muted"))

        row = QtWidgets.QHBoxLayout()
        row.addStretch(1)
        row.addWidget(QtWidgets.QLabel("年份"))
        self.year_combo = QtWidgets.QComboBox()
        self.year_combo.currentTextChanged.connect(self._on_selected_month_changed)
        row.addWidget(self.year_combo)
        row.addSpacing(12)
        row.addWidget(QtWidgets.QLabel("月份"))
        self.month_combo = QtWidgets.QComboBox()
        self.month_combo.addItems([f"{m:02d}" for m in range(1, 13)])
        self.month_combo.currentTextChanged.connect(self._on_selected_month_changed)
        row.addWidget(self.month_combo)
        row.addStretch(1)
        month_layout.addLayout(row)

        button_row = QtWidgets.QHBoxLayout()
        self.upload_button = QtWidgets.QPushButton("上传所选月份2个表")
        self.upload_button.setObjectName("successAction")
        self.upload_button.clicked.connect(self.upload_all_monthly_files)
        self.check_button = QtWidgets.QPushButton("检查当前年份文件")
        self.check_button.clicked.connect(self.check_selected_year_files)
        self.run_button = QtWidgets.QPushButton("生成结果文件")
        self.run_button.setObjectName("primaryAction")
        self.run_button.clicked.connect(self.run_report)
        self.open_button = QtWidgets.QPushButton("打开生成好的 Excel")
        self.open_button.setObjectName("ghostAction")
        self.open_button.clicked.connect(self.open_output_file)
        for button in (self.upload_button, self.check_button, self.run_button, self.open_button):
            button_row.addWidget(button)
        month_layout.addLayout(button_row)

        status_card = QtWidgets.QWidget(objectName="softCard")
        status_layout = QtWidgets.QVBoxLayout(status_card)
        status_layout.setContentsMargins(14, 12, 14, 12)
        self.selected_month_state = QtWidgets.QLabel("当前月份状态：未检查")
        self.selected_month_state.setStyleSheet(f"font-weight:700; color:{TEXT};")
        self.selected_month_detail = QtWidgets.QLabel("选择年份和月份后，会显示该月是否已经有完整文件。")
        self.selected_month_detail.setWordWrap(True)
        self.selected_month_attendance = QtWidgets.QLabel("考勤文件：未检查")
        self.selected_month_leave = QtWidgets.QLabel("请假文件：未检查")
        status_layout.addWidget(self.selected_month_state)
        status_layout.addWidget(self.selected_month_detail)
        status_layout.addWidget(self.selected_month_attendance)
        status_layout.addWidget(self.selected_month_leave)
        month_layout.addWidget(status_card)

        helper_row = QtWidgets.QHBoxLayout()
        btn_open_month = QtWidgets.QPushButton("打开所选月份文件夹")
        btn_open_month.setObjectName("ghostAction")
        btn_open_month.clicked.connect(self.open_selected_month_folder)
        btn_open_data = QtWidgets.QPushButton("打开放文件夹")
        btn_open_data.setObjectName("ghostAction")
        btn_open_data.clicked.connect(self.open_data_dir)
        btn_log = QtWidgets.QPushButton("打开运行日志")
        btn_log.setObjectName("ghostAction")
        btn_log.clicked.connect(self.open_runtime_log)
        helper_row.addWidget(btn_open_month)
        helper_row.addWidget(btn_open_data)
        helper_row.addWidget(btn_log)
        helper_row.addStretch(1)
        month_layout.addLayout(helper_row)
        layout.addWidget(month_box)
        return card

    def _build_month_card(self) -> QtWidgets.QWidget:
        card = QtWidgets.QWidget(objectName="card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)

        title_row = QtWidgets.QHBoxLayout()
        title_row.addWidget(QtWidgets.QLabel("月份列表", objectName="sectionTitle"))
        title_row.addStretch(1)
        self.issue_only_checkbox = QtWidgets.QCheckBox("仅显示待处理月份")
        self.issue_only_checkbox.stateChanged.connect(self._on_issue_only_changed)
        title_row.addWidget(self.issue_only_checkbox)
        layout.addLayout(title_row)
        layout.addWidget(QtWidgets.QLabel("双击某个月份可以直接打开对应文件夹。", objectName="muted"))

        self.month_table = QtWidgets.QTableWidget(0, 5)
        self.month_table.setHorizontalHeaderLabels(["月份", "考勤文件", "请假文件", "状态", "问题说明"])
        self.month_table.horizontalHeader().setStretchLastSection(True)
        self.month_table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.month_table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.month_table.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        self.month_table.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        self.month_table.verticalHeader().setVisible(False)
        self.month_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.month_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.month_table.setAlternatingRowColors(False)
        self.month_table.cellDoubleClicked.connect(self.open_tree_selected_month_folder)
        layout.addWidget(self.month_table, 1)
        return card

    def _build_log_card(self) -> QtWidgets.QWidget:
        card = QtWidgets.QWidget(objectName="card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)
        layout.addWidget(QtWidgets.QLabel("运行日志", objectName="sectionTitle"))
        layout.addWidget(QtWidgets.QLabel("这里只显示最近日志，完整内容请点“打开运行日志”。", objectName="muted"))
        self.log_view = QtWidgets.QTextEdit()
        self.log_view.setReadOnly(True)
        layout.addWidget(self.log_view, 1)
        return card

    def _set_notice(self, title: str, desc: str, level: str = "info") -> None:
        if level == "success":
            self.notice_title.setStyleSheet(f"color:{SUCCESS}; font-size:14px; font-weight:700;")
            self.notice_desc.setStyleSheet(f"color:{TEXT_MUTED}; font-size:11px;")
        elif level == "error":
            self.notice_title.setStyleSheet(f"color:{DANGER}; font-size:14px; font-weight:700;")
            self.notice_desc.setStyleSheet(f"color:{TEXT_MUTED}; font-size:11px;")
        else:
            self.notice_title.setStyleSheet(f"color:{PRIMARY}; font-size:14px; font-weight:700;")
            self.notice_desc.setStyleSheet(f"color:{TEXT_MUTED}; font-size:11px;")
        self.notice_title.setText(title)
        self.notice_desc.setText(desc)

    def _available_years(self) -> list[int]:
        current_year = date.today().year
        years = set(range(current_year, current_year + 11))
        for bundle in self.current_bundles:
            years.add(bundle.year)
        if self.data_dir.exists():
            for child in self.data_dir.iterdir():
                if not child.is_dir():
                    continue
                parsed = _parse_month_input(child.name)
                if parsed:
                    years.add(parsed[0])
        if self.year_combo.currentText().isdigit():
            years.add(int(self.year_combo.currentText()))
        return sorted(years)

    def _refresh_year_values(self) -> None:
        values = [str(year) for year in self._available_years()]
        current = self.year_combo.currentText() or str(date.today().year)
        self.year_combo.blockSignals(True)
        self.year_combo.clear()
        self.year_combo.addItems(values)
        if current in values:
            self.year_combo.setCurrentText(current)
        elif values:
            self.year_combo.setCurrentText(values[-1])
        self.year_combo.blockSignals(False)
        if not self.month_combo.currentText():
            self.month_combo.setCurrentText(f"{date.today().month:02d}")

    def _get_selected_year_month(self) -> tuple[int, int]:
        year_text = self.year_combo.currentText().strip()
        month_text = self.month_combo.currentText().strip()
        if not year_text.isdigit():
            raise ValueError("请先选择年份。")
        if not month_text.isdigit():
            raise ValueError("请先选择月份。")
        year = int(year_text)
        month = int(month_text)
        if not 1 <= month <= 12:
            raise ValueError("月份不正确。")
        return year, month

    def _get_selected_year(self) -> int:
        return self._get_selected_year_month()[0]

    def _selected_month_folder_name(self) -> str:
        year, month = self._get_selected_year_month()
        return f"{year}-{month:02d}"

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

    def _get_folder_inspections(self, selected_year: int | None) -> list[MonthFolderInspection]:
        cache_key = ("folder", self._build_scan_cache_key(self.data_dir, selected_year))
        if self._folder_cache_key == cache_key:
            return list(self._folder_cache_inspections)
        inspections = inspect_month_source_folders(str(self.data_dir), target_year=selected_year)
        self._folder_cache_key = cache_key
        self._folder_cache_inspections = list(inspections)
        return list(inspections)

    def refresh_folder_overview(self) -> None:
        if self._closing:
            return
        self.last_scan_error_message = ""
        self._refresh_annual_info()
        self._populate_bundles([], full_scan=False)

    def _refresh_annual_info(self) -> None:
        try:
            summary = get_current_annual_leave_summary(str(self.data_dir))
        except Exception:
            self.annual_info_label.setText("当前年假表：未上传")
            return
        updated_at = summary["updated_at"].strftime("%Y-%m-%d %H:%M")
        self.annual_info_label.setText(f"{summary['file_name']} | 员工数 {summary['employee_count']} | 更新于 {updated_at}")

    def _selected_file_status_color(self, text: str) -> str:
        if "已找到" in text:
            return SUCCESS
        if "重复" in text or "未找到" in text:
            return DANGER
        return TEXT

    def _apply_selected_file_status_colors(self) -> None:
        self.selected_month_attendance.setStyleSheet(f"color:{self._selected_file_status_color(self.selected_month_attendance.text())}; font-weight:600;")
        self.selected_month_leave.setStyleSheet(f"color:{self._selected_file_status_color(self.selected_month_leave.text())}; font-weight:600;")

    def _on_selected_month_changed(self, _value: object = None) -> None:
        try:
            year, month = self._get_selected_year_month()
            folder_name = f"{year}-{month:02d}"
        except ValueError:
            return
        inspections = self._get_folder_inspections(year)
        selected_info = next((item for item in inspections if item.year == year and item.month == month), None)
        if selected_info is None or not selected_info.has_any_data:
            self.selected_month_state.setText(f"{folder_name}：未放文件")
            self.selected_month_detail.setText("这个月份文件夹里还没有识别到考勤或请假文件。")
            self.selected_month_attendance.setText("考勤文件：未找到")
            self.selected_month_leave.setText("请假文件：未找到")
            self.summary_label.setText(f"当前准备上传：{folder_name}")
            self.status_label.setText(f"{folder_name} 还没有放入文件。选择好该月后，点击“上传所选月份2个表”。")
        elif selected_info.ready:
            self.selected_month_state.setText(f"{folder_name}：已就绪")
            self.selected_month_detail.setText(f"考勤：{selected_info.attendance_files[0].name}；请假：{selected_info.leave_files[0].name}")
            self.selected_month_attendance.setText(f"考勤文件：已找到（{selected_info.attendance_files[0].name}）")
            self.selected_month_leave.setText(f"请假文件：已找到（{selected_info.leave_files[0].name}）")
            self.summary_label.setText(f"{folder_name} 已有完整文件")
            self.status_label.setText(f"{folder_name} 已识别到考勤和请假文件，可以直接生成结果。")
        else:
            self.selected_month_state.setText(f"{folder_name}：待处理")
            self.selected_month_detail.setText(selected_info.detail or "该月文件还不完整，请补齐后再生成。")
            if len(selected_info.attendance_files) == 1:
                self.selected_month_attendance.setText(f"考勤文件：已找到（{selected_info.attendance_files[0].name}）")
            elif len(selected_info.attendance_files) > 1:
                self.selected_month_attendance.setText("考勤文件：重复")
            else:
                self.selected_month_attendance.setText("考勤文件：未找到")
            if len(selected_info.leave_files) == 1:
                self.selected_month_leave.setText(f"请假文件：已找到（{selected_info.leave_files[0].name}）")
            elif len(selected_info.leave_files) > 1:
                self.selected_month_leave.setText("请假文件：重复")
            else:
                self.selected_month_leave.setText("请假文件：未找到")
            self.summary_label.setText(f"{folder_name} 已有部分文件")
            self.status_label.setText(f"{folder_name} 还需处理：{selected_info.detail}")
        self._apply_selected_file_status_colors()

    def _rotate_runtime_log_if_needed(self) -> None:
        try:
            if not self.runtime_log_path.exists() or self.runtime_log_path.stat().st_size < RUNTIME_LOG_MAX_BYTES:
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

    def _start_runtime_log_session(self) -> None:
        self._rotate_runtime_log_if_needed()
        header = [
            "",
            "=" * 72,
            f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {APP_NAME} {APP_VERSION} PySide6 session start",
            f"cwd: {Path.cwd()}",
            f"script_root: {_app_root()}",
            "=" * 72,
        ]
        self.runtime_log_path.parent.mkdir(parents=True, exist_ok=True)
        with self.runtime_log_path.open("a", encoding="utf-8") as handle:
            handle.write("\n".join(header) + "\n")

    def _append_runtime_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for line in (message.splitlines() or [""]):
            self._runtime_log_buffer.append(f"{timestamp} | {line}\n")

    def _flush_runtime_log_buffer(self) -> None:
        if not self._runtime_log_buffer:
            return
        self.runtime_log_path.parent.mkdir(parents=True, exist_ok=True)
        self._rotate_runtime_log_if_needed()
        with self.runtime_log_path.open("a", encoding="utf-8") as handle:
            handle.writelines(self._runtime_log_buffer)
        self._runtime_log_buffer.clear()

    def _format_scan_duration(self, seconds: float) -> str:
        return f"耗时 {seconds:.2f} 秒"

    def log(self, message: str) -> None:
        self.runtime_logs.append(message)
        if len(self.runtime_logs) > 300:
            self.runtime_logs = self.runtime_logs[-300:]
        self.log_view.setPlainText("\n".join(self.runtime_logs))
        cursor = self.log_view.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        self.log_view.setTextCursor(cursor)
        self._append_runtime_log(message)
        QtCore.QTimer.singleShot(250, self._flush_runtime_log_buffer)

    def clear_log(self) -> None:
        self.runtime_logs.clear()
        self.log_view.clear()

    def _current_annual_target_path(self, suffix: str = ".xlsx") -> Path:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        return self.data_dir / f"{CURRENT_ANNUAL_TARGET_STEM}{suffix}"

    def choose_data_dir(self) -> None:
        selected = QtWidgets.QFileDialog.getExistingDirectory(self, "选择放文件的文件夹", str(self.data_dir))
        if selected:
            self.data_dir = Path(selected)
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
        self.data_dir.mkdir(parents=True, exist_ok=True)
        _open_path(self.data_dir)

    def open_runtime_log(self) -> None:
        if not self.runtime_log_path.exists():
            self._start_runtime_log_session()
        self._flush_runtime_log_buffer()
        _open_path(self.runtime_log_path)

    def open_output_file(self) -> None:
        if not self.output_file.exists():
            QtWidgets.QMessageBox.information(self, "提示", "结果文件还没有生成。")
            return
        _open_path(self.output_file)

    def open_current_annual_file(self) -> None:
        try:
            summary = get_current_annual_leave_summary(str(self.data_dir))
        except Exception:
            QtWidgets.QMessageBox.information(self, "提示", "当前还没有上传年假总数表。")
            return
        _open_path(Path(summary["path"]))

    def upload_current_annual_file(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        selected, _ = QtWidgets.QFileDialog.getOpenFileName(self, "请选择当前员工年假总数表", str(self.data_dir), "Excel Files (*.xls *.xlsx)")
        if not selected:
            return
        source_path = Path(selected)
        suffix = source_path.suffix.lower()
        if suffix not in {".xls", ".xlsx"}:
            QtWidgets.QMessageBox.critical(self, "上传失败", "只支持 Excel 文件：.xls 或 .xlsx")
            return
        target_path = self._current_annual_target_path(suffix)
        existing_files = _find_matching_files(self.data_dir, FILE_KIND_CONFIG["annual"]["patterns"])
        if existing_files and QtWidgets.QMessageBox.question(self, "确认覆盖", "当前年假总数表已存在，是否用新文件覆盖？") != QtWidgets.QMessageBox.Yes:
            return
        for existing in existing_files:
            if existing.exists():
                existing.unlink()
        shutil.copy2(source_path, target_path)
        self._refresh_annual_info()
        self.log(f"已上传当前年假表 -> {target_path}")
        self._set_notice("当前年假表已更新。", "只有员工年假信息发生变化时，才需要再次上传。", "success")
        self._scan_cache_key = None
        self._folder_cache_key = None
        self.refresh_folder_overview()

    def _ensure_selected_month_dir(self) -> Path:
        year, month = self._get_selected_year_month()
        self.data_dir.mkdir(parents=True, exist_ok=True)
        month_dir = self.data_dir / f"{year}-{month:02d}"
        month_dir.mkdir(parents=True, exist_ok=True)
        return month_dir

    def _copy_monthly_file(self, month_dir: Path, kind: str, source_path: Path) -> Path:
        suffix = source_path.suffix.lower()
        if suffix not in {".xls", ".xlsx"}:
            raise ValueError("只支持 Excel 文件：.xls 或 .xlsx")
        existing_files = _find_matching_files(month_dir, FILE_KIND_CONFIG[kind]["patterns"])
        if existing_files and QtWidgets.QMessageBox.question(self, "确认覆盖", f"{month_dir.name} 文件夹里已经有“{FILE_KIND_CONFIG[kind]['label']}”。\n是否用新文件覆盖？") != QtWidgets.QMessageBox.Yes:
            raise RuntimeError("用户取消覆盖")
        for existing in existing_files:
            if existing.exists():
                existing.unlink()
        target_path = month_dir / f"{FILE_KIND_CONFIG[kind]['label']}{suffix}"
        shutil.copy2(source_path, target_path)
        return target_path

    def upload_all_monthly_files(self) -> None:
        try:
            month_dir = self._ensure_selected_month_dir()
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "请选择年月", str(exc))
            return
        uploaded = 0
        for kind in ("attendance", "leave"):
            label = FILE_KIND_CONFIG[kind]["label"]
            selected, _ = QtWidgets.QFileDialog.getOpenFileName(self, f"请选择{label}", str(month_dir), "Excel Files (*.xls *.xlsx)")
            if not selected:
                return
            try:
                target_path = self._copy_monthly_file(month_dir, kind, Path(selected))
            except RuntimeError:
                return
            except Exception as exc:
                QtWidgets.QMessageBox.critical(self, "上传失败", str(exc))
                return
            self.log(f"已上传 {label} -> {target_path}")
            uploaded += 1
        if uploaded:
            self.summary_label.setText(f"{month_dir.name} 的 2 个文件已上传完成")
            self.status_label.setText(f"{month_dir.name} 文件夹文件已齐，请点击“生成结果文件”。")
            self._set_notice(f"{month_dir.name} 的 2 个文件已上传完成。", "现在可以直接点击“生成结果文件”。程序会自动检查后再生成。", "success")
            self._scan_cache_key = None
            self._folder_cache_key = None
            self.refresh_folder_overview()

    def _inspect_month_folders(self) -> tuple[list[dict[str, object]], list[str]]:
        if not self.data_dir.exists():
            return [], ["放文件的文件夹不存在。"]
        rows: list[dict[str, object]] = []
        issues: list[str] = []
        try:
            target_year = self._get_selected_year()
        except ValueError:
            target_year = None
        inspections = self._get_folder_inspections(target_year)
        for item in inspections:
            attendance_count = len(item.attendance_files)
            leave_count = len(item.leave_files)
            if item.ready:
                status_text, detail_text, tag = "[已就绪]", "", "ok"
            elif attendance_count == 0 and leave_count == 0:
                status_text, detail_text, tag = "[未放文件]", item.detail or "还没有识别到考勤或请假文件", "empty"
            elif "重复" in item.detail:
                status_text, detail_text, tag = "[文件重复]", item.detail, "warn"
            else:
                status_text, detail_text, tag = "[已放1个文件]", item.detail or "该月文件还不完整", "partial"
            if detail_text:
                issues.append(f"{item.folder_name}: {detail_text}")
            rows.append({
                "month": item.folder_name,
                "attendance": item.attendance_files[0].name if attendance_count == 1 else ("未放" if attendance_count == 0 else f"{attendance_count}个文件"),
                "leave": item.leave_files[0].name if leave_count == 1 else ("未放" if leave_count == 0 else f"{leave_count}个文件"),
                "status": status_text,
                "detail": detail_text,
                "tag": tag,
            })
        return rows, issues

    def _populate_bundles(self, bundles: list[MonthlySourceBundle], full_scan: bool = True) -> None:
        folder_rows, issues = self._inspect_month_folders()
        self.current_bundles = list(bundles)
        self.month_issue_details = {str(row["month"]): str(row.get("detail", "")) for row in folder_rows if row.get("detail")}
        display_rows = [row for row in folder_rows if (not self.issue_only_checkbox.isChecked() or row.get("detail"))]
        self._refresh_year_values()
        self.month_table.setRowCount(0)
        self.month_table.setColumnHidden(4, not self.issue_only_checkbox.isChecked())
        palette = {
            "ok": QtGui.QColor("#EDF8F2"),
            "warn": QtGui.QColor("#FCEDEA"),
            "partial": QtGui.QColor("#FFF6E8"),
            "empty": QtGui.QColor("#F4F6F8"),
        }
        for row_data in display_rows:
            row = self.month_table.rowCount()
            self.month_table.insertRow(row)
            values = [row_data["month"], row_data["attendance"], row_data["leave"], row_data["status"], row_data["detail"]]
            for col, value in enumerate(values):
                item = QtWidgets.QTableWidgetItem(str(value))
                item.setBackground(palette.get(str(row_data["tag"]), QtGui.QColor("white")))
                if col in {0, 3}:
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.month_table.setItem(row, col, item)
        ready_folder_count = sum(1 for row in folder_rows if row.get("tag") == "ok")
        if issues and bundles:
            self.status_label.setText(f"已识别 {len(bundles)} 个月份，另有部分月份待处理。")
            self.summary_label.setText(f"当前可统计 {len(bundles)} 个月份，其余月份待处理")
            self._set_notice("部分月份还需要处理。", "已就绪的月份仍可先生成；显示 [需处理] 的月份可以稍后补齐。", "info")
        elif issues:
            self.status_label.setText("有月份文件夹需要处理，请先补齐或去重。")
            self.summary_label.setText("当前还没有完整月份可统计")
            self._set_notice("有月份文件夹还需要处理。", "请看月份列表的状态标签，显示 [需处理] 的月份先处理。", "error")
        elif not full_scan and ready_folder_count:
            self.status_label.setText(f"当前年份发现 {ready_folder_count} 个月份文件已就绪。")
            self.summary_label.setText("已完成轻量检查，点击“检查当前年份文件”可做完整识别")
            self._set_notice("已完成轻量检查。", f"当前年份发现 {ready_folder_count} 个已就绪月份。需要完整识别时，请点击“检查当前年份文件”。", "info")
        elif bundles:
            self.status_label.setText(f"已找到 {len(bundles)} 个月份，年度合计会自动一起计算。")
            self.summary_label.setText(f"已识别 {len(bundles)} 个月份，可以直接生成")
            self._set_notice("检查通过，可以直接生成结果文件。", f"当前共识别到 {len(bundles)} 个月份。点击“生成结果文件”即可。", "success")
        else:
            self._on_selected_month_changed()
            self.status_label.setText("没有找到可统计的月份。")
            self.summary_label.setText("请先上传当前年假表，并上传某个月份的 2 个表")
            self._set_notice("还没有找到可以统计的月份。", "请先上传当前年假表，再选好年份和月份，点击“上传所选月份2个表”。", "info")

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
        lines.append(f"已识别月份：{'、'.join(ready_months) if ready_months else '无'}")
        lines.append(f"待处理月份：{'、'.join(issue_months) if issue_months else '无'}")
        if issue_months:
            lines.append("")
            lines.extend(issues[:6])
        QtWidgets.QMessageBox.information(self, "检查结果", "\n".join(lines))

    def check_selected_year_files(self) -> None:
        self.scan_bundles()
        self._show_year_check_result()

    def _show_generation_success_dialog(self, summary: ReportSummary) -> None:
        processed_months = [f"{item.bundle.month:02d}" for item in summary.monthly_results]
        rows, issues = self._inspect_month_folders()
        skipped_months = [str(row["month"]) for row in rows if row.get("detail")]
        dialog = QtWidgets.QMessageBox(self)
        dialog.setWindowTitle("生成完成")
        dialog.setIcon(QtWidgets.QMessageBox.Information)
        dialog.setText("结果文件已生成")
        dialog.setInformativeText(
            "\n".join(
                [
                    f"保存位置：{summary.output_file}",
                    f"已统计月份：{'、'.join(processed_months) if processed_months else '无'}",
                    f"待处理月份：{'、'.join(skipped_months) if skipped_months else '无'}",
                    *(issues[:5] if issues else []),
                ]
            )
        )
        open_excel = dialog.addButton("打开 Excel", QtWidgets.QMessageBox.AcceptRole)
        open_folder = dialog.addButton("打开结果文件夹", QtWidgets.QMessageBox.ActionRole)
        dialog.addButton("关闭", QtWidgets.QMessageBox.RejectRole)
        dialog.exec()
        clicked = dialog.clickedButton()
        if clicked == open_excel:
            _open_path(summary.output_file)
        elif clicked == open_folder:
            _open_path(summary.output_file.parent)

    def scan_bundles(self, preserve_log: bool = False) -> None:
        if self._closing:
            return
        started_at = time.perf_counter()
        try:
            selected_year = self._get_selected_year()
        except ValueError:
            selected_year = None
        cache_key = self._build_scan_cache_key(self.data_dir, selected_year)
        if self._scan_cache_key == cache_key:
            self._last_scan_duration_seconds = time.perf_counter() - started_at
            self._last_scan_used_cache = True
            self._last_scan_feedback = f"自上次完整检查后未发现文件变化，本次直接复用结果。{self._format_scan_duration(self._last_scan_duration_seconds)}"
            self.last_scan_error_message = self._scan_cache_error
            self._refresh_annual_info()
            self._populate_bundles(self._scan_cache_bundles, full_scan=True)
            if not preserve_log:
                self.clear_log()
                self.log(f"放文件的文件夹：{self.data_dir}")
                self.log("检查结果来自缓存，本次未重复读取 Excel。")
            return
        previous_key = self._last_full_scan_key
        file_changed = previous_key is not None and previous_key != cache_key
        try:
            bundles = discover_monthly_source_bundles(str(self.data_dir), target_year=selected_year, relaxed=True)
        except Exception as exc:
            friendly = _friendly_scan_error(exc)
            self._last_scan_duration_seconds = time.perf_counter() - started_at
            prefix = "检测到文件变化，已重新检查。" if file_changed else "已完成完整检查。"
            self._last_scan_feedback = f"{prefix}{self._format_scan_duration(self._last_scan_duration_seconds)}"
            self._last_scan_used_cache = False
            self._scan_cache_key = cache_key
            self._scan_cache_bundles = []
            self._scan_cache_error = friendly
            self._last_full_scan_key = cache_key
            self.last_scan_error_message = friendly
            self._refresh_annual_info()
            self._populate_bundles([], full_scan=True)
            self.status_label.setText("检查失败，请先处理文件夹中的问题。")
            self._set_notice("文件检查失败。", friendly, "error")
            if not preserve_log:
                self.clear_log()
            self.log(f"检查失败：{friendly}")
            self.log(str(exc))
            return
        self._scan_cache_key = cache_key
        self._scan_cache_bundles = list(bundles)
        self._scan_cache_error = ""
        self._last_full_scan_key = cache_key
        self._last_scan_duration_seconds = time.perf_counter() - started_at
        prefix = "检测到文件变化，已重新检查。" if file_changed else "已完成完整检查。"
        self._last_scan_feedback = f"{prefix}{self._format_scan_duration(self._last_scan_duration_seconds)}"
        self._last_scan_used_cache = False
        self.last_scan_error_message = ""
        if not preserve_log:
            self.clear_log()
            self.log(f"放文件的文件夹：{self.data_dir}")
            for bundle in bundles:
                self.log(f"找到月份 {bundle.year}-{bundle.month:02d} | {bundle.attendance_file.name} | {bundle.leave_file.name}")
        self._refresh_annual_info()
        self._populate_bundles(bundles, full_scan=True)

    def run_report(self) -> None:
        if self._current_worker is not None:
            return
        try:
            selected_year = self._get_selected_year()
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "请选择年月", str(exc))
            return
        cache_key = self._build_scan_cache_key(self.data_dir, selected_year)
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
            detail = self.last_scan_error_message or "请先上传当前年假表，并上传至少一个月份的 2 个表。"
            QtWidgets.QMessageBox.information(self, "无法生成", f"还没有找到可统计的月份。\n\n{detail}")
            return
        self.output_file.parent.mkdir(parents=True, exist_ok=True)
        self.clear_log()
        self.log(f"开始生成统计。放文件的文件夹：{self.data_dir}")
        self.log(f"结果会保存到：{self.output_file}")
        if generation_cache_reused:
            self.log("生成前缓存有效，未重复执行完整检查。")
        for button in (self.upload_button, self.check_button, self.run_button):
            button.setEnabled(False)
        self.summary_label.setText("正在生成，请稍等 10-30 秒")
        self._set_notice("正在生成结果文件，请稍等。", "生成过程中不要重复点击按钮。完成后会自动提示你打开结果文件。", "info")

        worker = ReportWorker(str(self.data_dir), str(self.output_file), selected_year)
        self._current_worker = worker
        worker.signals.message.connect(self.log)
        worker.signals.done.connect(self._on_report_done)
        worker.signals.error.connect(self._on_report_error)
        self._thread_pool.start(worker)

    def _on_report_done(self, summary: object) -> None:
        assert isinstance(summary, ReportSummary)
        self.log("统计完成。")
        self.summary_label.setText(f"已经生成完成，共统计 {len(summary.monthly_results)} 个月份")
        self.status_label.setText(f"结果文件已生成：{summary.output_file}")
        self._set_notice("结果文件已生成完成。", f"保存位置：{summary.output_file}", "success")
        self._show_generation_success_dialog(summary)
        self._current_worker = None
        for button in (self.upload_button, self.check_button, self.run_button):
            button.setEnabled(True)
        self.scan_bundles(preserve_log=True)

    def _on_report_error(self, payload: str) -> None:
        self.log("生成失败：")
        self.log(str(payload))
        self.summary_label.setText("生成失败，请检查放进去的文件")
        self.status_label.setText("生成失败，请检查文件是否完整、文件名是否正确。")
        self._set_notice("生成失败。", "请检查当前年假表和各月份的 2 个文件是否完整、文件名是否正确。", "error")
        QtWidgets.QMessageBox.critical(self, "生成失败", "统计过程中出现错误。\n请检查文件是否完整、文件名是否正确。")
        self._current_worker = None
        for button in (self.upload_button, self.check_button, self.run_button):
            button.setEnabled(True)

    def open_selected_month_folder(self) -> None:
        try:
            month_dir = self._ensure_selected_month_dir()
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "请选择年月", str(exc))
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
                ]),
                encoding="utf-8",
            )
        _open_path(month_dir)

    def open_tree_selected_month_folder(self, row: int | None = None, _column: int | None = None) -> None:
        current_row = row if row is not None else self.month_table.currentRow()
        if current_row < 0:
            QtWidgets.QMessageBox.information(self, "提示", "请先在月份列表里选中一个月份。")
            return
        month_name = self.month_table.item(current_row, 0).text()
        month_dir = self.data_dir / month_name
        month_dir.mkdir(parents=True, exist_ok=True)
        issue_detail = self.month_issue_details.get(month_name, "")
        if issue_detail:
            QtWidgets.QMessageBox.information(self, "该月份需要处理", f"{month_name} 当前问题：\n{issue_detail}")
        _open_path(month_dir)

    def _on_issue_only_changed(self, _state: object = None) -> None:
        self._populate_bundles(self.current_bundles)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    window = AttendancePySide6Window()
    window.show()
    return app.exec()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception as exc:  # pragma: no cover
        log_file = _write_startup_diagnostic(exc)
        print(f"Startup failed. See log: {log_file}", file=sys.stderr)
        raise
