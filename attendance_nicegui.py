#!/usr/bin/env python3
"""NiceGUI version of the attendance tool."""

from __future__ import annotations

import os
import re
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from attendance_report import (
    DATA_DIR,
    OUTPUT_FILE,
    MonthlySourceBundle,
    ReportSummary,
    discover_monthly_source_bundles,
    generate_report,
)

try:
    from nicegui import events, run, ui
except Exception as exc:  # pragma: no cover
    raise SystemExit(
        "NiceGUI is not installed. Run: python -m pip install nicegui"
    ) from exc

try:
    _original_nicegui_setup = run.setup

    def _safe_nicegui_setup() -> None:
        try:
            _original_nicegui_setup()
        except PermissionError:
            # Some sandboxed macOS environments block NiceGUI's ProcessPool setup.
            # This app only uses run.io_bound, so it is safe to continue without cpu_bound support.
            return

    run.setup = _safe_nicegui_setup
except Exception:
    pass


FILE_KIND_CONFIG = {
    "attendance": {
        "label": "考勤打卡记录表",
        "file_names": ["考勤打卡记录表.xls", "考勤打卡记录表.xlsx"],
        "accept": ".xls,.xlsx",
    },
    "leave": {
        "label": "请假记录表",
        "file_names": ["请假记录表.xls", "请假记录表.xlsx"],
        "accept": ".xls,.xlsx",
    },
    "annual": {
        "label": "员工年假总数表",
        "file_names": ["员工年假总数表.xls", "员工年假总数表.xlsx"],
        "accept": ".xls,.xlsx",
    },
}

PAGE_CSS = """
body {
  background: linear-gradient(180deg, #f6f1e7 0%, #efe8da 100%);
  color: #1f2a2c;
}
.nice-shell {
  max-width: 1440px;
  margin: 0 auto;
}
.metric-card {
  background: rgba(255,255,255,0.78);
  border: 1px solid #ddd4c5;
  border-radius: 20px;
  box-shadow: 0 12px 30px rgba(41, 33, 24, 0.08);
}
.panel-card {
  background: rgba(255,255,255,0.88);
  border: 1px solid #ddd4c5;
  border-radius: 24px;
  box-shadow: 0 14px 36px rgba(41, 33, 24, 0.08);
}
.hero-card {
  background: linear-gradient(135deg, #183a37 0%, #24514c 100%);
  color: #f9f3ea;
  border-radius: 28px;
  box-shadow: 0 20px 50px rgba(24, 58, 55, 0.22);
}
.action-button {
  min-height: 64px;
  border-radius: 18px;
  font-weight: 700;
  font-size: 18px;
}
.soft-button {
  border-radius: 16px;
  font-weight: 600;
}
.step-chip {
  background: rgba(255,255,255,0.12);
  border: 1px solid rgba(255,255,255,0.18);
  border-radius: 999px;
  padding: 10px 14px;
}
.wizard-step {
  background: #fbf8f1;
  border: 1px solid #e6ddcf;
  border-radius: 16px;
  box-shadow: none;
  transition: all 0.2s ease;
}
.wizard-step-active {
  background: #fff2e6;
  border: 1px solid #e8b183;
  box-shadow: 0 10px 24px rgba(201, 107, 59, 0.12);
}
.wizard-step-done {
  background: #edf6ee;
  border: 1px solid #b9d7bf;
}
.wizard-step-error {
  background: #fbeceb;
  border: 1px solid #e5b8b4;
}
.upload-button-active {
  background: #c96b3b !important;
  color: #ffffff !important;
  box-shadow: 0 12px 24px rgba(201, 107, 59, 0.22);
}
.upload-button-done {
  background: #2f6b3b !important;
  color: #ffffff !important;
}
.upload-button-error {
  background: #a33b2b !important;
  color: #ffffff !important;
}
.check-button-active {
  background: #1f5a33 !important;
  color: #ffffff !important;
  box-shadow: 0 12px 24px rgba(47, 107, 59, 0.24);
}
.run-button-active {
  background: #c96b3b !important;
  color: #ffffff !important;
  box-shadow: 0 14px 28px rgba(201, 107, 59, 0.28);
}
.open-button-active {
  background: #24514c !important;
  color: #ffffff !important;
  box-shadow: 0 12px 24px rgba(36, 81, 76, 0.24);
}
.log-box textarea {
  min-height: 240px !important;
  font-family: Consolas, Monaco, monospace;
}
.footer-note {
  opacity: 0.72;
  font-size: 12px;
}
"""


def _app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _default_data_dir() -> Path:
    return (_app_root() / DATA_DIR).resolve()


def _default_output_file() -> Path:
    return (_app_root() / OUTPUT_FILE).resolve()


def _usage_file() -> Path:
    return (_app_root() / "使用说明-给同事看.txt").resolve()


def _open_local_path(path: Path) -> None:
    target = path if path.exists() else path.parent
    if sys.platform.startswith("win"):
        os.startfile(str(target))  # type: ignore[attr-defined]
        return
    if sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=False)
        return
    subprocess.run(["xdg-open", str(target)], check=False)


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


@dataclass
class DashboardState:
    data_dir: Path = field(default_factory=_default_data_dir)
    output_file: Path = field(default_factory=_default_output_file)
    upload_month: str = field(default_factory=lambda: f"{date.today().year}-{date.today().month:02d}")
    selected_month: str = ""
    current_bundles: list[MonthlySourceBundle] = field(default_factory=list)
    log_lines: list[str] = field(default_factory=list)

    def log(self, message: str) -> None:
        self.log_lines.append(message)
        if len(self.log_lines) > 400:
            self.log_lines = self.log_lines[-400:]


state = DashboardState()


def _normalize_month_folder_name(raw_text: str) -> str:
    parsed = _parse_month_input(raw_text)
    if parsed is None:
        if not raw_text.strip():
            return _next_month_name()
        raise ValueError("月份请按 2026-03、2026/03 或 2026年3月 这样的格式填写。")
    year, month = parsed
    return f"{year}-{month:02d}"


def _next_month_name() -> str:
    if state.current_bundles:
        latest = max(state.current_bundles, key=lambda item: (item.year, item.month))
        year, month = latest.year, latest.month + 1
        if month > 12:
            year += 1
            month = 1
        return f"{year}-{month:02d}"
    today = date.today()
    return f"{today.year}-{today.month:02d}"


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
        return "还没有找到考勤文件。请先把每个月的 3 个文件放进月份文件夹。"
    return message


def _find_existing_month_files(month_dir: Path, kind: str) -> list[Path]:
    return [
        month_dir / file_name
        for file_name in FILE_KIND_CONFIG[kind]["file_names"]
        if (month_dir / file_name).exists()
    ]


def inspect_month_folders(base_dir: Path) -> tuple[list[dict[str, str]], list[str]]:
    if not base_dir.exists():
        return [], ["放文件的文件夹不存在。"]

    rows: list[dict[str, str]] = []
    issues: list[str] = []
    month_dirs = sorted([path for path in base_dir.iterdir() if path.is_dir()])

    for month_dir in month_dirs:
        attendance_files = _find_existing_month_files(month_dir, "attendance")
        leave_files = _find_existing_month_files(month_dir, "leave")
        annual_files = _find_existing_month_files(month_dir, "annual")

        problems = []
        if not attendance_files:
            problems.append("缺少考勤文件")
        elif len(attendance_files) > 1:
            problems.append("重复考勤文件")
        if not leave_files:
            problems.append("缺少请假文件")
        elif len(leave_files) > 1:
            problems.append("重复请假文件")
        if not annual_files:
            problems.append("缺少年假文件")
        elif len(annual_files) > 1:
            problems.append("重复年假文件")

        status_text = "文件齐全" if not problems else "；".join(problems)
        status_color = "positive" if not problems else "negative"
        if problems:
            issues.append(f"{month_dir.name}: {', '.join(problems)}")

        rows.append(
            {
                "month": month_dir.name,
                "attendance": attendance_files[0].name if len(attendance_files) == 1 else ("未放" if not attendance_files else f"{len(attendance_files)}个文件"),
                "leave": leave_files[0].name if len(leave_files) == 1 else ("未放" if not leave_files else f"{len(leave_files)}个文件"),
                "annual": annual_files[0].name if len(annual_files) == 1 else ("未放" if not annual_files else f"{len(annual_files)}个文件"),
                "status": status_text,
                "status_color": status_color,
                "actions": "",
            }
        )

    return rows, issues


def inspect_single_month(month_dir: Path) -> dict[str, object]:
    status_map: dict[str, dict[str, str]] = {}
    ready_count = 0
    next_missing_label = "可以先上传考勤打卡记录表"
    next_missing_kind = "attendance"
    first_issue_found = False

    for kind in ("attendance", "leave", "annual"):
        files = _find_existing_month_files(month_dir, kind)
        if len(files) == 1:
            status_map[kind] = {
                "text": "已就绪",
                "color": "positive",
                "detail": files[0].name,
            }
            ready_count += 1
        elif len(files) > 1:
            status_map[kind] = {
                "text": "重复文件",
                "color": "negative",
                "detail": f"{len(files)} 个同类文件",
            }
            if not first_issue_found:
                next_missing_label = f"请先处理重复的 {FILE_KIND_CONFIG[kind]['label']}"
                next_missing_kind = kind
                first_issue_found = True
        else:
            status_map[kind] = {
                "text": "未上传",
                "color": "warning",
                "detail": f"缺少 {FILE_KIND_CONFIG[kind]['label']}",
            }
            if not first_issue_found:
                next_missing_label = f"下一步建议上传：{FILE_KIND_CONFIG[kind]['label']}"
                next_missing_kind = kind
                first_issue_found = True

    all_ready = ready_count == 3 and all(item["text"] == "已就绪" for item in status_map.values())
    if all_ready:
        next_missing_label = "当前月份 3 个文件都已到位，可以点击“先检查文件”"
        next_missing_kind = ""

    return {
        "ready_count": ready_count,
        "progress": ready_count / 3,
        "status_map": status_map,
        "next_label": next_missing_label,
        "next_kind": next_missing_kind,
        "all_ready": all_ready,
    }


async def ensure_month_dir() -> Path:
    sync_inputs()
    base_dir = state.data_dir
    base_dir.mkdir(parents=True, exist_ok=True)
    folder_name = _normalize_month_folder_name(state.upload_month)
    state.upload_month = folder_name
    month_dir = base_dir / folder_name
    month_dir.mkdir(parents=True, exist_ok=True)
    return month_dir


def _write_guide_file(month_dir: Path) -> None:
    guide_file = month_dir / "请把这3个文件放到这里.txt"
    if guide_file.exists():
        return
    guide_file.write_text(
        "\n".join(
            [
                "请把下面 3 个文件放到这个文件夹里：",
                "",
                "1. 考勤打卡记录表.xls",
                "2. 请假记录表.xls",
                "3. 员工年假总数表.xlsx",
            ]
        ),
        encoding="utf-8",
    )


def _store_uploaded_file(month_dir: Path, kind: str, filename: str, content: bytes) -> Path:
    suffix = Path(filename).suffix.lower()
    if suffix not in {".xls", ".xlsx"}:
        raise ValueError("只支持 .xls 或 .xlsx 文件。")
    for existing in _find_existing_month_files(month_dir, kind):
        existing.unlink()
    target = month_dir / f"{FILE_KIND_CONFIG[kind]['label']}{suffix}"
    target.write_bytes(content)
    return target


async def save_upload(kind: str, event: events.UploadEventArguments) -> None:
    try:
        month_dir = await ensure_month_dir()
        content = event.content.read()
        target = await run.io_bound(_store_uploaded_file, month_dir, kind, event.name, content)
        state.log(f"已上传 {FILE_KIND_CONFIG[kind]['label']} -> {target}")
        ui.notify(f"{FILE_KIND_CONFIG[kind]['label']} 已上传", color="positive")
        await scan_files(show_notify=False)
        refresh_upload_wizard()
    except Exception as exc:
        ui.notify(str(exc), color="negative")


async def scan_files(show_notify: bool = True) -> None:
    sync_inputs()
    base_dir = state.data_dir
    rows, issues = inspect_month_folders(base_dir)
    notice_payload: tuple[str, str, str] | None = None
    try:
        bundles = await run.io_bound(discover_monthly_source_bundles, str(base_dir))
    except Exception as exc:
        bundles = []
        friendly_message = _friendly_scan_error(exc)
        state.log(f"检查失败：{friendly_message}")
        state.log(str(exc))
        notice_payload = ("文件检查失败", friendly_message, "negative")
        if show_notify:
            ui.notify(friendly_message, color="negative")
    else:
        state.current_bundles = list(bundles)
        if bundles:
            state.upload_month = _next_month_name()
            state.log(f"已识别 {len(bundles)} 个月份")
    update_dashboard(rows, issues, notice_payload)


async def create_next_month_folder() -> None:
    try:
        sync_inputs()
        state.upload_month = _next_month_name()
        month_dir = await ensure_month_dir()
        await run.io_bound(_write_guide_file, month_dir)
        state.log(f"已创建月份文件夹：{month_dir}")
        _open_local_path(month_dir)
        ui.notify(f"已创建 {month_dir.name}", color="positive")
        await scan_files(show_notify=False)
        refresh_upload_wizard()
    except Exception as exc:
        ui.notify(str(exc), color="negative")


async def generate_excel() -> None:
    sync_inputs()
    rows, issues = inspect_month_folders(state.data_dir)
    if not rows:
        ui.notify("还没有可统计的月份，请先放文件。", color="warning")
        return
    if issues:
        ui.notify("还有缺少或重复文件的月份，请先处理。", color="warning")
        return

    set_notice("正在生成 Excel", "生成过程中请不要重复点击按钮。", "warning")
    state.log(f"开始生成：data={state.data_dir} output={state.output_file}")
    try:
        logs: list[str] = []
        summary = await run.io_bound(
            generate_report,
            str(state.data_dir),
            str(state.output_file),
            lambda msg: logs.append(str(msg)),
        )
        for line in logs:
            state.log(line)
        assert isinstance(summary, ReportSummary)
        state.log(f"统计完成：{summary.output_file}")
        set_notice("统计已完成", f"已生成 {summary.output_file.name}，共统计 {len(summary.monthly_results)} 个月份。下一步请点击“打开生成好的 Excel”。", "positive")
        ui.notify("Excel 已生成完成", color="positive")
        await scan_files(show_notify=False)
        refresh_open_result_button(summary.output_file.exists())
        log_box.value = "\n".join(state.log_lines)
    except Exception as exc:
        state.log(f"生成失败：{exc}")
        set_notice("生成失败", str(exc), "negative")
        ui.notify(str(exc), color="negative")


def update_dashboard(
    rows: list[dict[str, str]],
    issues: list[str],
    notice_payload: tuple[str, str, str] | None = None,
) -> None:
    month_table.rows = rows
    month_table.update()
    detected_label.set_text(str(len(rows)))
    ready_label.set_text(str(sum(1 for row in rows if row["status"] == "文件齐全")))
    issue_label.set_text(str(len(issues)))
    output_label.set_text(state.output_file.name)
    month_input.value = state.upload_month
    data_dir_input.value = str(state.data_dir)
    output_input.value = str(state.output_file)
    selected_month_label.set_text(state.selected_month or "未选择")
    log_box.value = "\n".join(state.log_lines)
    refresh_upload_wizard()
    refresh_run_button(rows, issues)
    refresh_open_result_button(state.output_file.exists())
    if notice_payload:
        set_notice(*notice_payload)
    elif issues:
        set_notice("有月份文件夹需要处理", "状态列里显示缺少或重复文件的月份，先处理后再生成。", "negative")
    elif rows:
        set_notice("文件检查通过", f"已识别 {len(rows)} 个月份。下一步请点击“生成 Excel 统计表”。", "positive")
    else:
        set_notice("还没有找到月份", "请先创建月份文件夹并放入 3 个文件。", "warning")


def set_notice(title: str, desc: str, level: str) -> None:
    notice_title.text = title
    notice_desc.text = desc
    hero_status_label.text = title
    if level == "positive":
        notice_box.classes(replace="panel-card q-pa-md row items-start q-col-gutter-md bg-green-1")
    elif level == "negative":
        notice_box.classes(replace="panel-card q-pa-md row items-start q-col-gutter-md bg-red-1")
    else:
        notice_box.classes(replace="panel-card q-pa-md row items-start q-col-gutter-md bg-orange-1")


def on_data_dir_change() -> None:
    state.data_dir = Path(data_dir_input.value).expanduser().resolve()


def on_output_change() -> None:
    state.output_file = Path(output_input.value).expanduser().resolve()
    output_label.set_text(state.output_file.name)


def sync_inputs() -> None:
    state.upload_month = month_input.value or state.upload_month
    on_data_dir_change()
    on_output_change()


def select_month(month_name: str) -> None:
    state.selected_month = month_name
    state.upload_month = month_name
    month_input.value = month_name
    selected_month_label.set_text(month_name)
    refresh_upload_wizard()
    ui.notify(f"已选中 {month_name}", color="primary")


def open_current_month_folder() -> None:
    try:
        month_name = _normalize_month_folder_name(month_input.value or state.upload_month)
    except ValueError as exc:
        ui.notify(str(exc), color="negative")
        return
    month_dir = state.data_dir / month_name
    month_dir.mkdir(parents=True, exist_ok=True)
    _write_guide_file(month_dir)
    _open_local_path(month_dir)


def open_selected_month_folder() -> None:
    if not state.selected_month:
        ui.notify("请先在月份列表中点选一个月份。", color="warning")
        return
    month_dir = state.data_dir / state.selected_month
    month_dir.mkdir(parents=True, exist_ok=True)
    _open_local_path(month_dir)


def handle_table_use_month(event: events.GenericEventArguments) -> None:
    month_name = str(event.args["month"])
    select_month(month_name)


def handle_table_open_month(event: events.GenericEventArguments) -> None:
    month_name = str(event.args["month"])
    state.selected_month = month_name
    month_dir = state.data_dir / month_name
    month_dir.mkdir(parents=True, exist_ok=True)
    _open_local_path(month_dir)


def open_output_file() -> None:
    if not state.output_file.exists():
        ui.notify("结果文件还没有生成。", color="warning")
        return
    _open_local_path(state.output_file)


def reset_default_paths() -> None:
    state.data_dir = _default_data_dir()
    state.output_file = _default_output_file()
    data_dir_input.value = str(state.data_dir)
    output_input.value = str(state.output_file)
    output_label.set_text(state.output_file.name)
    refresh_upload_wizard()
    ui.notify("已恢复默认路径。", color="positive")


def clear_logs() -> None:
    state.log_lines.clear()
    log_box.value = ""
    ui.notify("日志已清空。", color="primary")


def refresh_upload_wizard() -> None:
    try:
        month_name = _normalize_month_folder_name(month_input.value or state.upload_month)
    except ValueError:
        upload_wizard_month_label.set_text("月份格式不正确")
        upload_wizard_hint.set_text("请先把月份改成 2026-03 这样的格式。")
        upload_progress.value = 0
        for kind in ("attendance", "leave", "annual"):
            upload_step_badges[kind].set_text("待处理")
            upload_step_badges[kind].props("color=grey-6")
            upload_step_details[kind].set_text(FILE_KIND_CONFIG[kind]["label"])
        return

    month_dir = state.data_dir / month_name
    month_dir.mkdir(parents=True, exist_ok=True)
    info = inspect_single_month(month_dir)
    upload_wizard_month_label.set_text(month_name)
    upload_progress.value = float(info["progress"])
    upload_wizard_hint.set_text(str(info["next_label"]))

    status_map = info["status_map"]
    for kind in ("attendance", "leave", "annual"):
        item = status_map.get(kind, {"text": "待处理", "color": "grey-6", "detail": FILE_KIND_CONFIG[kind]["label"]})
        upload_step_badges[kind].set_text(str(item["text"]))
        upload_step_badges[kind].props(f"color={item['color']}")
        upload_step_details[kind].set_text(str(item["detail"]))
        if str(item["text"]) == "重复文件":
            upload_step_cards[kind].classes(replace="col q-pa-sm wizard-step wizard-step-error")
            upload_action_buttons[kind].classes(replace="w-full soft-button upload-button-error")
        elif str(item["text"]) == "已就绪":
            upload_step_cards[kind].classes(replace="col q-pa-sm wizard-step wizard-step-done")
            upload_action_buttons[kind].classes(replace="w-full soft-button upload-button-done")
        elif str(info["next_kind"]) == kind and not bool(info["all_ready"]):
            upload_step_cards[kind].classes(replace="col q-pa-sm wizard-step wizard-step-active")
            upload_action_buttons[kind].classes(replace="w-full soft-button upload-button-active")
        else:
            upload_step_cards[kind].classes(replace="col q-pa-sm wizard-step")
            upload_action_buttons[kind].classes(replace="w-full soft-button")

    if bool(info["all_ready"]):
        hero_check_button.classes(replace="action-button check-button-active")
        upload_check_button.classes(replace="soft-button col check-button-active")
    else:
        hero_check_button.classes(replace="action-button")
        upload_check_button.classes(replace="soft-button col")


def refresh_run_button(rows: list[dict[str, str]], issues: list[str]) -> None:
    if rows and not issues:
        hero_run_button.classes(replace="action-button run-button-active")
    else:
        hero_run_button.classes(replace="action-button")


def refresh_open_result_button(ready: bool) -> None:
    if ready:
        hero_open_button.classes(replace="soft-button open-button-active")
        table_open_button.classes(replace="soft-button open-button-active")
    else:
        hero_open_button.classes(replace="soft-button")
        table_open_button.classes(replace="soft-button")


async def upload_attendance(event: events.UploadEventArguments) -> None:
    await save_upload("attendance", event)


async def upload_leave(event: events.UploadEventArguments) -> None:
    await save_upload("leave", event)


async def upload_annual(event: events.UploadEventArguments) -> None:
    await save_upload("annual", event)


ui.add_head_html(f"<style>{PAGE_CSS}</style>")
ui.colors(primary="#183A37", secondary="#C96B3B", positive="#2F6B3B", negative="#A33B2B", warning="#A66500")
ui.page_title("财务公司考勤统计助手")

with ui.column().classes('nice-shell w-full q-pa-lg gap-4'):
    with ui.row().classes('w-full items-stretch q-col-gutter-lg'):
        with ui.column().classes('col-12 col-md-7 hero-card q-pa-xl'):
            ui.label('财务公司考勤统计助手').classes('text-h4 text-weight-bold')
            ui.label('这是 NiceGUI 版本。界面更现代，但流程仍然只有三步：放文件、检查文件、生成 Excel。').classes('text-body1 opacity-90')
            with ui.row().classes('q-gutter-sm q-mt-md'):
                ui.label('1. 放入本月 3 个文件').classes('step-chip text-body2')
                ui.label('2. 点击先检查文件').classes('step-chip text-body2')
                ui.label('3. 点击生成 Excel').classes('step-chip text-body2')
            with ui.row().classes('q-gutter-sm q-mt-md'):
                hero_check_button = ui.button('先检查文件', on_click=scan_files, color='green-8').classes('action-button')
                hero_run_button = ui.button('生成 Excel 统计表', on_click=generate_excel, color='deep-orange-6').classes('action-button')
                ui.button('自动创建下个月文件夹', on_click=create_next_month_folder, color='blue-grey-7').classes('soft-button')
            with ui.row().classes('q-gutter-sm q-mt-sm'):
                hero_open_button = ui.button('打开生成好的 Excel', on_click=open_output_file, color='brown-6').classes('soft-button')
                ui.button('打开使用说明', on_click=lambda: _open_local_path(_usage_file()), color='grey-8').classes('soft-button')
        with ui.column().classes('col-12 col-md-5 gap-3'):
            with ui.card().classes('metric-card q-pa-md'):
                ui.label('当前状态').classes('text-caption text-grey-7')
                hero_status_label = ui.label('等待检查').classes('text-h6 text-weight-bold')
                ui.label('建议先检查文件，再生成结果。').classes('text-body2 text-grey-7')
            with ui.row().classes('q-col-gutter-sm'):
                with ui.card().classes('metric-card q-pa-md col'):
                    ui.label('已识别月份').classes('text-caption text-grey-7')
                    detected_label = ui.label('0').classes('text-h4 text-weight-bold')
                with ui.card().classes('metric-card q-pa-md col'):
                    ui.label('文件齐全').classes('text-caption text-grey-7')
                    ready_label = ui.label('0').classes('text-h4 text-weight-bold')
                with ui.card().classes('metric-card q-pa-md col'):
                    ui.label('待处理').classes('text-caption text-grey-7')
                    issue_label = ui.label('0').classes('text-h4 text-weight-bold')
            with ui.card().classes('metric-card q-pa-md'):
                ui.label('结果文件').classes('text-caption text-grey-7')
                output_label = ui.label(state.output_file.name).classes('text-subtitle1 text-weight-bold')

    with ui.card().classes('panel-card w-full q-pa-md') as notice_box:
        with ui.column().classes('gap-1'):
            notice_title = ui.label('请先放文件，再检查文件。').classes('text-subtitle1 text-weight-bold')
            notice_desc = ui.label('右边先上传，再看月份列表。').classes('text-body2 text-grey-8')

    with ui.tabs().classes('w-full') as tabs:
        upload_tab = ui.tab('文件上传')
        month_tab = ui.tab('月份和日志')

    with ui.tab_panels(tabs, value=upload_tab).classes('w-full'):
        with ui.tab_panel(upload_tab):
            with ui.row().classes('w-full q-col-gutter-lg items-start'):
                with ui.card().classes('panel-card col-12 col-lg-8 q-pa-lg'):
                    ui.label('本月上传区').classes('text-h6 text-weight-bold')
                    ui.label('先填月份，再上传。大多数人只需要用这一块。高级设置默认隐藏。').classes('text-body2 text-grey-7')
                    month_input = ui.input('月份', value=state.upload_month, placeholder='2026-03').classes('w-full q-mt-md')
                    month_input.on('blur', lambda _e: (setattr(state, 'upload_month', month_input.value), refresh_upload_wizard()))
                    with ui.card().classes('metric-card q-pa-md q-mt-md'):
                        with ui.row().classes('w-full items-center justify-between'):
                            with ui.column().classes('gap-1'):
                                ui.label('上传向导').classes('text-subtitle1 text-weight-bold')
                                upload_wizard_month_label = ui.label(state.upload_month).classes('text-body1 text-weight-bold')
                            upload_progress = ui.linear_progress(value=0).classes('w-40')
                        upload_wizard_hint = ui.label('下一步建议上传：考勤打卡记录表').classes('text-body2 text-grey-7 q-mt-sm')
                        with ui.row().classes('w-full q-col-gutter-md q-mt-sm'):
                            upload_step_cards = {}
                            upload_step_badges = {}
                            upload_step_details = {}
                            for kind in ('attendance', 'leave', 'annual'):
                                with ui.card().classes('col q-pa-sm wizard-step') as step_card:
                                    upload_step_cards[kind] = step_card
                                    ui.label(FILE_KIND_CONFIG[kind]['label']).classes('text-caption text-grey-7')
                                    upload_step_badges[kind] = ui.badge('待处理', color='grey-6')
                                    upload_step_details[kind] = ui.label(FILE_KIND_CONFIG[kind]['label']).classes('text-body2 q-mt-xs')
                    with ui.row().classes('w-full q-col-gutter-md q-mt-sm'):
                        ui.button('创建并打开这个月份文件夹', on_click=create_next_month_folder, color='blue-grey-7').classes('soft-button col')
                        ui.button('打开当前月份文件夹', on_click=open_current_month_folder, color='brown-6').classes('soft-button col')
                        upload_check_button = ui.button('先检查文件', on_click=scan_files, color='green-8').classes('soft-button col')
                    ui.separator().classes('q-my-md')
                    ui.label('直接上传文件').classes('text-subtitle1 text-weight-bold')
                    with ui.row().classes('w-full q-col-gutter-md items-start'):
                        upload_action_buttons = {}
                        with ui.column().classes('col-12 col-md-4'):
                            ui.label('考勤打卡记录表').classes('text-body2 text-grey-8')
                            upload_action_buttons['attendance'] = ui.upload(on_upload=upload_attendance, auto_upload=True).props('accept=.xls,.xlsx').classes('w-full soft-button')
                        with ui.column().classes('col-12 col-md-4'):
                            ui.label('请假记录表').classes('text-body2 text-grey-8')
                            upload_action_buttons['leave'] = ui.upload(on_upload=upload_leave, auto_upload=True).props('accept=.xls,.xlsx').classes('w-full soft-button')
                        with ui.column().classes('col-12 col-md-4'):
                            ui.label('员工年假总数表').classes('text-body2 text-grey-8')
                            upload_action_buttons['annual'] = ui.upload(on_upload=upload_annual, auto_upload=True).props('accept=.xls,.xlsx').classes('w-full soft-button')
                with ui.column().classes('col-12 col-lg-4 gap-4'):
                    with ui.card().classes('panel-card q-pa-lg'):
                        ui.label('提示').classes('text-h6 text-weight-bold')
                        ui.markdown('- 先上传 3 个文件\n- 再检查月份\n- 最后生成 Excel\n- 如果状态列显示缺少或重复，先处理再生成')
                    with ui.expansion('高级设置（一般不用改）', icon='settings').classes('panel-card q-pa-sm'):
                        with ui.column().classes('q-pa-md gap-3'):
                            data_dir_input = ui.input('放文件的文件夹', value=str(state.data_dir)).classes('w-full')
                            data_dir_input.on('blur', lambda _e: on_data_dir_change())
                            output_input = ui.input('生成后的 Excel', value=str(state.output_file)).classes('w-full')
                            output_input.on('blur', lambda _e: on_output_change())
                            with ui.row().classes('q-col-gutter-sm'):
                                ui.button('打开放文件夹', on_click=lambda: _open_local_path(state.data_dir), color='grey-8').classes('soft-button')
                                ui.button('打开结果文件夹', on_click=lambda: _open_local_path(state.output_file.parent), color='grey-8').classes('soft-button')
                            ui.button('恢复默认路径', on_click=reset_default_paths, color='brown-5').classes('soft-button')

        with ui.tab_panel(month_tab):
            with ui.row().classes('w-full q-col-gutter-lg items-start'):
                with ui.card().classes('panel-card col-12 col-lg-7 q-pa-lg'):
                    ui.label('月份列表').classes('text-h6 text-weight-bold')
                    ui.label('状态列会提示缺少或重复的文件。').classes('text-body2 text-grey-7')
                    with ui.row().classes('items-center q-col-gutter-md q-mb-md'):
                        ui.label('当前选中月份').classes('text-body2 text-grey-7')
                        selected_month_label = ui.label('未选择').classes('text-subtitle1 text-weight-bold')
                        ui.button('打开选中月份文件夹', on_click=open_selected_month_folder, color='brown-6').classes('soft-button')
                    month_table = ui.table(
                        columns=[
                            {'name': 'month', 'label': '月份', 'field': 'month', 'align': 'left'},
                            {'name': 'attendance', 'label': '考勤文件', 'field': 'attendance', 'align': 'left'},
                            {'name': 'leave', 'label': '请假文件', 'field': 'leave', 'align': 'left'},
                            {'name': 'annual', 'label': '年假文件', 'field': 'annual', 'align': 'left'},
                            {'name': 'status', 'label': '状态', 'field': 'status', 'align': 'left'},
                            {'name': 'actions', 'label': '操作', 'field': 'actions', 'align': 'left'},
                        ],
                        rows=[],
                        row_key='month',
                        pagination=8,
                    ).classes('w-full')
                    month_table.add_slot(
                        'body-cell-status',
                        r'''
                        <q-td :props="props">
                          <q-badge :color="props.row.status_color" text-color="white" rounded :label="props.row.status" />
                        </q-td>
                        ''',
                    )
                    month_table.add_slot(
                        'body-cell-actions',
                        r'''
                        <q-td :props="props">
                          <div class="row q-gutter-sm">
                            <q-btn dense flat color="primary" icon="edit_calendar" label="设为当前月" @click="$parent.$emit('use-month', {month: props.row.month})" />
                            <q-btn dense flat color="secondary" icon="folder_open" label="打开文件夹" @click="$parent.$emit('open-month', {month: props.row.month})" />
                          </div>
                        </q-td>
                        ''',
                    )
                    month_table.on('rowClick', lambda e: select_month(str(e.args['row']['month'])))
                    month_table.on('use-month', handle_table_use_month)
                    month_table.on('open-month', handle_table_open_month)
                    with ui.row().classes('q-col-gutter-sm q-mt-md'):
                        ui.button('重新检查文件', on_click=scan_files, color='green-8').classes('soft-button')
                        table_open_button = ui.button('打开结果 Excel', on_click=open_output_file, color='deep-orange-6').classes('soft-button')
                with ui.card().classes('panel-card col-12 col-lg-5 q-pa-lg log-box'):
                    ui.label('运行记录').classes('text-h6 text-weight-bold')
                    ui.label('默认只看月份列表即可。需要排查问题时，再展开下面的详细日志。').classes('text-body2 text-grey-7')
                    with ui.expansion('查看详细日志', icon='article').classes('w-full q-mt-md'):
                        with ui.column().classes('q-pa-sm gap-2'):
                            log_box = ui.textarea(value='').props('readonly autogrow').classes('w-full')
                            ui.button('清空日志', on_click=clear_logs, color='grey-7').classes('soft-button')

    ui.label('桌面化简洁模式：默认只显示常用操作，路径设置和详细日志已折叠。').classes('footer-note self-end q-px-sm')

set_notice('请先放文件，再检查文件。', '你可以先在“文件上传”页放入本月的 3 个文件。', 'warning')


async def _initial_scan() -> None:
    await scan_files(show_notify=False)


ui.timer(0.2, _initial_scan, once=True)
ui.run(title='财务公司考勤统计助手', favicon='📊', reload=False, native=False, show=True, host='127.0.0.1', port=8188)
