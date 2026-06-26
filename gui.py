from __future__ import annotations

import os
import queue
import random
import re
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk
from typing import Iterable

from main import (
    DEFAULT_TEMPLATE,
    SYNC_DEFAULT_OUTPUT,
    CrawlCancelled,
    CrawlEvent,
    ExportRecord,
    MatchInfo,
    changed_handicap_schedule_ids,
    changed_schedule_ids,
    crawl_complete_company_schedule_records,
    crawl_historical_date_records,
    fetch_company_schedule_snapshot,
    filter_matches_by_date_window,
    historical_output_path,
    historical_week_output_path,
    handicap_change_signatures_by_schedule,
    record_signatures_by_schedule,
    read_export_records_from_workbook,
    read_selection_matches_from_workbook,
    sync_workbook,
)
# from startup_guard import StartupBlocked, enforce_startup_time_limit


PRIMARY = "#1E40AF"
SECONDARY = "#3B82F6"
CTA = "#F59E0B"
BG = "#F8FAFC"
TEXT = "#0F172A"
MUTED = "#475569"
BORDER = "#CBD5E1"
DEFAULT_MONITOR_INTERVAL_MINUTES = 1.0
DEFAULT_RANDOM_DELAY_SECONDS = 10.0
HANDICAP_ALERT_THRESHOLD = 3
HANDICAP_DANGER_THRESHOLD = 4
HANDICAP_ALERT_BLUE = "#1E3A8A"
HANDICAP_ALERT_RED = "#DC2626"
DEFAULT_HANDICAP_CHANGE_BACKGROUND = "#DBEAFE"
CHECKED_MARK = "☑"
UNCHECKED_MARK = "☐"
SELECTED_DATA_COLUMNS = [
    ("league", "联赛", 70),
    ("event_time", "时间", 78),
    ("status", "状态", 52),
    ("home", "比赛球队", 140),
    ("score", "比分", 52),
    ("away", "比赛球队", 140),
    ("asian_index_up", "指数", 58),
    ("asian_current_line", "盘口", 70),
    ("asian_index_down", "指数", 58),
    ("asian_time", "时间节点", 88),
    ("asian_up", "上盘", 58),
    ("asian_line", "盘", 82),
    ("asian_down", "下盘", 58),
    ("asian_odds_change", "赔率变动", 70),
    ("asian_change", "盘口变动", 70),
    ("total_time", "时间节点", 88),
    ("big", "大球", 58),
    ("total_line", "盘口", 70),
    ("small", "小球", 58),
    ("total_odds_change", "赔率变动", 70),
    ("total_change", "盘口变动", 70),
]


def parse_float_setting(value: str, default: float, minimum: float) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return default
    if number < minimum:
        return default
    return number


def format_delay(seconds: float) -> str:
    if seconds >= 60:
        minutes = int(seconds // 60)
        remainder = int(round(seconds % 60))
        return f"{minutes} 分 {remainder} 秒"
    return f"{int(round(seconds))} 秒"


def parse_history_date_input(value: str) -> date:
    text = (value or "").strip()
    for fmt in ("%Y%m%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError("日期格式错误，请使用 YYYYMMDD 或 YYYY-MM-DD")


def alert_schedule_ids(updated_ids: list[str], selected_ids: list[str]) -> list[str]:
    selected = set(selected_ids)
    return [schedule_id for schedule_id in updated_ids if schedule_id in selected]


def handicap_alert_details(
    records: Iterable[ExportRecord],
    schedule_ids: Iterable[str],
    threshold: int = HANDICAP_ALERT_THRESHOLD,
) -> list[dict[str, int | str]]:
    selected = {schedule_id for schedule_id in schedule_ids if schedule_id}
    if not selected:
        return []
    by_schedule: dict[str, dict[str, int | str]] = {}
    for record in records:
        if record.row_type != "match" or record.schedule_id not in selected:
            continue
        asian_count = change_count_number(record.row_values[14] if len(record.row_values) > 14 else None)
        total_count = change_count_number(record.row_values[20] if len(record.row_values) > 20 else None)
        if asian_count >= threshold or total_count >= threshold:
            by_schedule[record.schedule_id] = {
                "schedule_id": record.schedule_id,
                "asian_count": asian_count,
                "total_count": total_count,
            }
    return [by_schedule[schedule_id] for schedule_id in schedule_ids if schedule_id in by_schedule]


def is_pending_match(match: MatchInfo) -> bool:
    return (match.status or "").strip() == "未"


def is_match_on_date(match: MatchInfo, target_date: date | None = None) -> bool:
    target_date = target_date or datetime.now().date()
    return match.match_time.date() == target_date


def auto_monitor_ids_for_today(matches: Iterable[MatchInfo], target_date: date | None = None) -> list[str]:
    return [match.schedule_id for match in matches if is_match_on_date(match, target_date) and is_pending_match(match)]


def keep_manual_monitor_match(match: MatchInfo, target_date: date | None = None) -> bool:
    target_date = target_date or datetime.now().date()
    return match.match_time.date() >= target_date


def merge_monitor_alert_ids(auto_ids: Iterable[str], manual_ids: Iterable[str]) -> list[str]:
    merged: list[str] = []
    for schedule_id in list(auto_ids) + list(manual_ids):
        if schedule_id and schedule_id not in merged:
            merged.append(schedule_id)
    return merged


def merge_selection_matches(primary: list[MatchInfo], secondary: list[MatchInfo]) -> list[MatchInfo]:
    return primary if primary else secondary


def selected_match_display_rows(records: Iterable[ExportRecord], selected_ids: Iterable[str]) -> list[tuple[str, ...]]:
    selected = {schedule_id for schedule_id in selected_ids if schedule_id}
    if not selected:
        return []
    rows: list[tuple[str, ...]] = []
    for record in records:
        if record.schedule_id not in selected or record.row_type == "separator":
            continue
        rows.append(tuple(format_display_cell(value) for value in record.row_values))
    return rows


def selected_match_info_display_rows(matches: dict[str, MatchInfo], selected_ids: Iterable[str]) -> list[tuple[str, ...]]:
    rows: list[tuple[str, ...]] = []
    for schedule_id in selected_ids:
        match = matches.get(schedule_id)
        if match is None:
            continue
        rows.append(
            (
                match.league,
                match.event_time,
                match.status,
                format_team_with_rank(match.home_team, match.home_rank),
                match.score,
                format_team_with_rank(match.away_team, match.away_rank),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            )
        )
    return rows


def match_change_count_display_values(records: Iterable[ExportRecord]) -> dict[str, tuple[str, str, str, str]]:
    values: dict[str, tuple[str, str, str, str]] = {}
    for record in records:
        if record.row_type != "match":
            continue
        if is_current_selected_data_row(record.row_values):
            odds_count = format_change_count_value(record.row_values[13] if len(record.row_values) > 13 else "")
            handicap_count = format_change_count_value(record.row_values[14] if len(record.row_values) > 14 else "")
            total_odds_count = format_change_count_value(record.row_values[19] if len(record.row_values) > 19 else "")
            total_handicap_count = format_change_count_value(record.row_values[20] if len(record.row_values) > 20 else "")
        elif is_previous_split_selected_data_row(record.row_values):
            odds_count = format_change_count_value(record.row_values[15] if len(record.row_values) > 15 else "")
            handicap_count = format_change_count_value(record.row_values[16] if len(record.row_values) > 16 else "")
            total_odds_count = format_change_count_value(record.row_values[21] if len(record.row_values) > 21 else "")
            total_handicap_count = "0"
        elif is_previous_rank_selected_data_row(record.row_values):
            odds_count = format_change_count_value(record.row_values[15] if len(record.row_values) > 15 else "")
            handicap_count = "0"
            total_odds_count = format_change_count_value(record.row_values[20] if len(record.row_values) > 20 else "")
            total_handicap_count = "0"
        else:
            odds_count = format_change_count_value(record.row_values[13] if len(record.row_values) > 13 else "")
            handicap_count = "0"
            total_odds_count = format_change_count_value(record.row_values[18] if len(record.row_values) > 18 else "")
            total_handicap_count = "0"
        values[record.schedule_id] = (odds_count, handicap_count, total_odds_count, total_handicap_count)
    return values


def is_current_selected_data_row(row_values: tuple[object | None, ...]) -> bool:
    if len(row_values) < 21:
        return False
    old_total_time = row_values[14] if len(row_values) > 14 else None
    new_total_time = row_values[15] if len(row_values) > 15 else None
    if old_total_time and not new_total_time and not is_change_count_cell(old_total_time):
        return False
    if new_total_time and is_change_count_cell(new_total_time):
        return False
    return True


def is_previous_split_selected_data_row(row_values: tuple[object | None, ...]) -> bool:
    return len(row_values) >= 22 and not is_current_selected_data_row(row_values)


def is_previous_rank_selected_data_row(row_values: tuple[object | None, ...]) -> bool:
    return len(row_values) >= 21 and not is_current_selected_data_row(row_values) and not is_previous_split_selected_data_row(row_values)


def is_change_count_cell(value: object | None) -> bool:
    if value in (None, ""):
        return True
    if isinstance(value, (int, float)):
        return True
    return re.fullmatch(r"\d+(?:\.0)?", str(value).strip()) is not None


def format_change_count_value(value: object | None) -> str:
    if value in (None, ""):
        return "0"
    return format_display_cell(value)


def change_count_number(value: object | None) -> int:
    try:
        return int(float(str(value)))
    except (TypeError, ValueError):
        return 0


def change_count_at_least(value: object | None, threshold: int = HANDICAP_ALERT_THRESHOLD) -> bool:
    return change_count_number(value) >= threshold


def handicap_change_count_color(value: object | None) -> str:
    return handicap_change_count_color_with_palette(value, HANDICAP_ALERT_BLUE, HANDICAP_ALERT_RED)


def handicap_change_count_color_with_palette(value: object | None, low_color: str, high_color: str) -> str:
    if value in (None, ""):
        return ""
    count = change_count_number(value)
    return high_color if count >= HANDICAP_DANGER_THRESHOLD else low_color


def normalized_hex_color(value: object | None, fallback: str) -> str:
    text = str(value or "").strip()
    if re.fullmatch(r"#[0-9A-Fa-f]{6}", text):
        return text.upper()
    return fallback


def selected_data_row_needs_handicap_warning(row: tuple[str, ...]) -> bool:
    return (
        (len(row) > 14 and change_count_at_least(row[14]))
        or (len(row) > 20 and change_count_at_least(row[20]))
    )


def match_tree_values_need_handicap_warning(values: list[str] | tuple[str, ...]) -> bool:
    return (
        (len(values) > 7 and change_count_at_least(values[7]))
        or (len(values) > 9 and change_count_at_least(values[9]))
    )


class TreeviewCellColorOverlay:
    def __init__(self, tree: ttk.Treeview, columns: Iterable[str], color_for_value, background_for_value=None) -> None:
        self.tree = tree
        self.columns = tuple(columns)
        self.color_for_value = color_for_value
        self.background_for_value = background_for_value
        self.labels: dict[tuple[str, str], tk.Label] = {}
        self.refresh_pending = False
        for event in ("<Configure>", "<Expose>", "<MouseWheel>", "<ButtonRelease-1>", "<KeyRelease>"):
            self.tree.bind(event, self.schedule_refresh, add="+")

    def yview(self, *args) -> None:
        self.tree.yview(*args)
        self.schedule_refresh()

    def xview(self, *args) -> None:
        self.tree.xview(*args)
        self.schedule_refresh()

    def schedule_refresh(self, *_args) -> None:
        if self.refresh_pending:
            return
        self.refresh_pending = True
        self.tree.after_idle(self.refresh)

    def refresh(self) -> None:
        self.refresh_pending = False
        visible_keys: set[tuple[str, str]] = set()
        for item in self.tree.get_children(""):
            for column in self.columns:
                bbox = self.tree.bbox(item, column)
                if not bbox:
                    continue
                value = self.tree.set(item, column)
                color = self.color_for_value(value)
                if not color:
                    continue
                key = (item, column)
                visible_keys.add(key)
                label = self.labels.get(key)
                if label is None:
                    label = tk.Label(self.tree, bd=0, padx=0, pady=0, anchor="center", bg="white")
                    self.labels[key] = label
                background = self.background_for_value(value) if self.background_for_value else "white"
                label.configure(text=value, fg=color, bg=background or "white")
                label.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
                label.lift()

        for key, label in list(self.labels.items()):
            if key not in visible_keys:
                if not self.tree.exists(key[0]):
                    label.destroy()
                    del self.labels[key]
                else:
                    label.place_forget()


def format_team_with_rank(team: str, rank: str) -> str:
    rank = (rank or "").strip()
    team = team or ""
    return f"[{rank}] {team}" if rank else team


def format_display_cell(value: object | None) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%m-%d %H:%M")
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def default_output_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


@dataclass(frozen=True)
class GuiConfig:
    output_dir: Path = field(default_factory=default_output_dir)
    output_name: str = SYNC_DEFAULT_OUTPUT
    template_path: Path = Path(DEFAULT_TEMPLATE)
    company_id: int = 1
    workers: int = 6
    monitor_interval_minutes: float = DEFAULT_MONITOR_INTERVAL_MINUTES
    random_delay_seconds: float = DEFAULT_RANDOM_DELAY_SECONDS

    def resolve_output_path(self, now: datetime | None = None) -> Path:
        name = self.output_name.strip() or SYNC_DEFAULT_OUTPUT
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        return self.output_dir / name

    def next_monitor_delay_seconds(self, random_value: float | None = None) -> float:
        random_part = random.random() if random_value is None else random_value
        interval_seconds = max(0.01, self.monitor_interval_minutes) * 60
        jitter_seconds = max(0.0, self.random_delay_seconds) * max(0.0, min(1.0, random_part))
        return interval_seconds + jitter_seconds


class GuiEventFormatter:
    def status_text(self, event: CrawlEvent) -> str:
        if event.type == "start":
            return "运行中"
        if event.type == "match_done":
            return f"{event.completed}/{event.total}"
        if event.type == "complete":
            return "完成"
        if event.type == "error":
            return "部分失败"
        if event.type == "cancelled":
            return "已取消"
        return "就绪"

    def log_text(self, event: CrawlEvent) -> str:
        if event.type == "start":
            return f"开始抓取完整赛事：{event.total} 场"
        if event.type == "match_done":
            return f"已完成 {event.completed}/{event.total}：{event.schedule_id}"
        if event.type == "complete":
            return f"导出数据准备完成：{event.total} 场，{event.rows} 行"
        if event.type == "error":
            return f"盘口抓取失败 {event.schedule_id}：{event.message}"
        if event.type == "cancelled":
            return "任务已取消"
        return event.message or event.type


class Titan007ExporterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Titan007 数据导出工具")
        self.root.geometry("1120x720")
        self.root.minsize(980, 640)
        self.formatter = GuiEventFormatter()
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.monitor_cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.monitor_worker: threading.Thread | None = None
        self.match_loader: threading.Thread | None = None
        self.last_output_path: Path | None = None
        self.today_matches: dict[str, MatchInfo] = {}
        self.auto_monitor_alert_ids: list[str] = []
        self.monitor_alert_ids: list[str] = []

        self.output_dir_var = tk.StringVar(value=str(default_output_dir()))
        self.history_date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.history_week_var = tk.BooleanVar(value=False)
        self.interval_minutes_var = tk.StringVar(value=str(int(DEFAULT_MONITOR_INTERVAL_MINUTES)))
        self.random_delay_seconds_var = tk.StringVar(value=str(int(DEFAULT_RANDOM_DELAY_SECONDS)))
        self.handicap_bg_var = tk.StringVar(value=DEFAULT_HANDICAP_CHANGE_BACKGROUND)
        self.handicap_low_text_color_var = tk.StringVar(value=HANDICAP_ALERT_BLUE)
        self.handicap_high_text_color_var = tk.StringVar(value=HANDICAP_ALERT_RED)
        self.status_var = tk.StringVar(value="就绪")
        self.monitor_status_var = tk.StringVar(value="未监控")
        self.matches_var = tk.StringVar(value="-")
        self.rows_var = tk.StringVar(value="-")
        self.failures_var = tk.StringVar(value="0")
        self.result_var = tk.StringVar(value="等待导出")
        self.progress_var = tk.DoubleVar(value=0)

        self.configure_styles()
        self.build_layout()
        for color_var in (
            self.handicap_bg_var,
            self.handicap_low_text_color_var,
            self.handicap_high_text_color_var,
        ):
            color_var.trace_add("write", lambda *_args: self.refresh_handicap_change_overlays())
        self.root.after(100, self.drain_messages)

    def configure_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background="white", relief="solid", borderwidth=1)
        style.configure("Title.TLabel", background=PRIMARY, foreground="white", font=("Microsoft YaHei UI", 14, "bold"))
        style.configure("HeaderSub.TLabel", background=PRIMARY, foreground="#DBEAFE", font=("Microsoft YaHei UI", 9))
        style.configure("CardTitle.TLabel", background="white", foreground=TEXT, font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Body.TLabel", background="white", foreground=TEXT, font=("Microsoft YaHei UI", 9))
        style.configure("Muted.TLabel", background="white", foreground=MUTED, font=("Microsoft YaHei UI", 8))
        style.configure("Metric.TLabel", background="white", foreground=PRIMARY, font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("TProgressbar", troughcolor="#E2E8F0", background=SECONDARY)

    def build_layout(self) -> None:
        self.root.configure(bg=BG)
        main = ttk.Frame(self.root, style="App.TFrame")
        main.pack(fill="both", expand=True)

        header = tk.Frame(main, bg=PRIMARY, height=76)
        header.pack(fill="x")
        header.pack_propagate(False)
        title_area = tk.Frame(header, bg=PRIMARY)
        title_area.pack(side="left", padx=18, pady=10)
        ttk.Label(title_area, text="Titan007 数据导出工具", style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_area, text="完整赛事 · 澳*盘口 · Excel 导出", style="HeaderSub.TLabel").pack(anchor="w")

        action_area = tk.Frame(header, bg=PRIMARY)
        action_area.pack(side="right", padx=18)
        self.header_status = tk.Label(action_area, textvariable=self.status_var, bg=PRIMARY, fg="white", font=("Microsoft YaHei UI", 9))
        self.header_status.pack(side="left", padx=(0, 12))
        self.start_button = tk.Button(
            action_area,
            text="开始抓取",
            command=self.start_export,
            bg=CTA,
            fg="#111827",
            activebackground="#D97706",
            relief="flat",
            padx=16,
            pady=8,
            cursor="hand2",
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        self.start_button.pack(side="left")
        self.cancel_button = tk.Button(
            action_area,
            text="取消",
            command=self.cancel_export,
            state="disabled",
            bg="#E2E8F0",
            fg=TEXT,
            relief="flat",
            padx=12,
            pady=8,
            cursor="hand2",
            font=("Microsoft YaHei UI", 9),
        )
        self.cancel_button.pack(side="left", padx=(8, 0))

        content = ttk.Frame(main, style="App.TFrame", padding=14)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=0, minsize=430)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        self.build_config_panel(content).grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        self.build_run_panel(content).grid(row=0, column=1, sticky="nsew")

    def build_config_panel(self, parent) -> ttk.Frame:
        panel = ttk.Frame(parent, style="Card.TFrame", padding=14)
        self.add_path_field(panel, "保存目录", self.output_dir_var, self.choose_output_dir)
        history_frame = ttk.Frame(panel, style="Card.TFrame")
        history_frame.pack(fill="x", pady=(0, 12))
        history_frame.columnconfigure(1, weight=1)
        ttk.Label(history_frame, text="历史日期", style="Body.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(history_frame, textvariable=self.history_date_var, width=12).grid(row=0, column=1, sticky="ew")
        ttk.Button(history_frame, text="抓取指定日期", command=self.start_history_export).grid(row=0, column=2, sticky="e", padx=(8, 0))
        ttk.Checkbutton(history_frame, text="按周获取", variable=self.history_week_var).grid(row=1, column=1, columnspan=2, sticky="w", pady=(6, 0))
        settings = ttk.Frame(panel, style="Card.TFrame")
        settings.pack(fill="x", pady=(0, 12))
        settings.columnconfigure(0, weight=1)
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(2, weight=1)
        self.add_small_entry(settings, "间隔(分钟)", self.interval_minutes_var, 0)
        self.add_small_entry(settings, "随机延迟(秒)", self.random_delay_seconds_var, 1)
        self.add_color_entry(settings, "盘口背景色", self.handicap_bg_var, self.choose_handicap_bg_color, 2)
        self.add_color_entry(settings, "盘口<4文字", self.handicap_low_text_color_var, self.choose_handicap_low_text_color, 0, row=1)
        self.add_color_entry(settings, "盘口>=4文字", self.handicap_high_text_color_var, self.choose_handicap_high_text_color, 1, row=1)

        ttk.Label(panel, text="可选赛事", style="CardTitle.TLabel").pack(anchor="w", pady=(2, 8))
        bottom_controls = ttk.Frame(panel, style="Card.TFrame")
        bottom_controls.pack(side="bottom", fill="x")
        match_actions = ttk.Frame(bottom_controls, style="Card.TFrame")
        match_actions.pack(fill="x", pady=(0, 10))
        self.load_matches_button = ttk.Button(match_actions, text="加载可选赛事", command=self.load_today_matches)
        self.load_matches_button.pack(side="left", fill="x", expand=True)
        self.add_monitor_button = ttk.Button(match_actions, text="添加赛事监控", command=self.add_checked_matches_to_monitor)
        self.add_monitor_button.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.monitor_button = ttk.Button(match_actions, text="启动监控", command=self.start_monitor)
        self.monitor_button.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.stop_monitor_button = ttk.Button(match_actions, text="停止监控", command=self.stop_monitor, state="disabled")
        self.stop_monitor_button.pack(side="left", fill="x", expand=True, padx=(8, 0))
        ttk.Label(bottom_controls, textvariable=self.monitor_status_var, style="Muted.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Button(bottom_controls, text="打开输出目录", command=self.open_output_dir).pack(fill="x", pady=(6, 6))
        ttk.Button(bottom_controls, text="打开最新 Excel", command=self.open_latest_file).pack(fill="x")

        match_frame = ttk.Frame(panel, style="Card.TFrame")
        match_frame.pack(fill="both", expand=True, pady=(0, 10))
        match_frame.rowconfigure(0, weight=1)
        match_frame.columnconfigure(0, weight=1)
        self.match_tree = ttk.Treeview(
            match_frame,
            columns=(
                "check",
                "time",
                "league",
                "home",
                "score",
                "away",
                "odds_change_count",
                "handicap_change_count",
                "total_odds_change_count",
                "total_handicap_change_count",
            ),
            show="headings",
            selectmode="none",
            height=12,
        )
        for column, title, width in [
            ("check", "监控", 46),
            ("time", "时间", 58),
            ("league", "联赛", 72),
            ("home", "主队", 92),
            ("score", "比分", 44),
            ("away", "客队", 92),
            ("odds_change_count", "赔率变动", 72),
            ("handicap_change_count", "盘口变动", 72),
            ("total_odds_change_count", "大小球赔率变动", 96),
            ("total_handicap_change_count", "大小球盘口变动", 96),
        ]:
            self.match_tree.heading(column, text=title)
            self.match_tree.column(column, width=width, anchor="center", stretch=column in {"home", "away"})
        match_scroll = ttk.Scrollbar(match_frame, orient="vertical")
        self.match_tree.configure(yscrollcommand=match_scroll.set)
        self.match_change_overlay = TreeviewCellColorOverlay(
            self.match_tree,
            ("handicap_change_count", "total_handicap_change_count"),
            self.handicap_change_text_color,
            lambda _value: self.handicap_change_background_color(),
        )
        match_scroll.configure(command=self.match_change_overlay.yview)
        self.match_tree.grid(row=0, column=0, sticky="nsew")
        match_scroll.grid(row=0, column=1, sticky="ns")
        self.match_tree.bind("<Button-1>", self.on_match_tree_click)
        return panel

    def build_run_panel(self, parent) -> ttk.Frame:
        panel = ttk.Frame(parent, style="App.TFrame")
        panel.rowconfigure(1, weight=1)
        panel.columnconfigure(0, weight=1)
        metrics = ttk.Frame(panel, style="App.TFrame")
        metrics.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        for index in range(4):
            metrics.columnconfigure(index, weight=1)
        self.metric_card(metrics, "赛事数", self.matches_var).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.metric_card(metrics, "数据行", self.rows_var).grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.metric_card(metrics, "失败盘口", self.failures_var).grid(row=0, column=2, sticky="ew", padx=(0, 10))
        self.metric_card(metrics, "状态", self.status_var).grid(row=0, column=3, sticky="ew")

        self.detail_tabs = ttk.Notebook(panel)
        self.detail_tabs.grid(row=1, column=0, sticky="nsew")
        progress_tab = ttk.Frame(self.detail_tabs, style="Card.TFrame", padding=14)
        progress_tab.rowconfigure(3, weight=1)
        progress_tab.columnconfigure(0, weight=1)
        ttk.Label(progress_tab, text="运行进度", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.progress_label = ttk.Label(progress_tab, text="等待开始", style="Muted.TLabel")
        self.progress_label.grid(row=0, column=0, sticky="e")
        self.progressbar = ttk.Progressbar(progress_tab, variable=self.progress_var, maximum=100)
        self.progressbar.grid(row=1, column=0, sticky="ew", pady=(10, 14))
        self.log_text = self.text_panel(progress_tab, "任务日志")
        self.log_text.master.grid(row=3, column=0, sticky="nsew")

        selected_tab = ttk.Frame(self.detail_tabs, style="Card.TFrame", padding=14)
        selected_tab.rowconfigure(0, weight=1)
        selected_tab.columnconfigure(0, weight=1)
        self.selected_data_panel(selected_tab).grid(row=0, column=0, sticky="nsew")
        self.detail_tabs.add(progress_tab, text="运行进度")
        self.detail_tabs.add(selected_tab, text="选中赛事数据")
        return panel

    def add_path_field(self, parent, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label, style="Muted.TLabel").pack(anchor="w")
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", pady=(4, 12))
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="选择", command=command).pack(side="left", padx=(6, 0))

    def add_small_entry(self, parent, label: str, variable: tk.StringVar, column: int, row: int = 0) -> None:
        cell = ttk.Frame(parent, style="Card.TFrame")
        cell.grid(row=row, column=column, sticky="ew", padx=(0, 8) if column == 0 else (8, 0), pady=(6, 0) if row else 0)
        ttk.Label(cell, text=label, style="Muted.TLabel").pack(anchor="w")
        ttk.Entry(cell, textvariable=variable, width=10).pack(fill="x", pady=(4, 0))

    def add_color_entry(self, parent, label: str, variable: tk.StringVar, command, column: int, row: int = 0) -> None:
        cell = ttk.Frame(parent, style="Card.TFrame")
        cell.grid(row=row, column=column, sticky="ew", padx=(0, 8) if column == 0 else (8, 0), pady=(6, 0) if row else 0)
        ttk.Label(cell, text=label, style="Muted.TLabel").pack(anchor="w")
        row = ttk.Frame(cell, style="Card.TFrame")
        row.pack(fill="x", pady=(4, 0))
        ttk.Entry(row, textvariable=variable, width=9).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="选", width=3, command=command).pack(side="left", padx=(4, 0))

    def metric_card(self, parent, title: str, variable: tk.StringVar) -> ttk.Frame:
        card = ttk.Frame(parent, style="Card.TFrame", padding=12)
        ttk.Label(card, text=title, style="Muted.TLabel").pack(anchor="w")
        ttk.Label(card, textvariable=variable, style="Metric.TLabel").pack(anchor="w", pady=(6, 0))
        return card

    def text_panel(self, parent, title: str) -> tk.Text:
        frame = ttk.Frame(parent, style="Card.TFrame", padding=0)
        ttk.Label(frame, text=title, style="CardTitle.TLabel", padding=8).pack(fill="x")
        text = tk.Text(frame, height=16, wrap="word", relief="flat", bg="white", fg="#334155", font=("Consolas", 10), padx=10, pady=8)
        text.pack(fill="both", expand=True)
        text.configure(state="disabled")
        return text

    def selected_data_panel(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame", padding=0)
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text="选中赛事数据", style="CardTitle.TLabel", padding=8).grid(row=0, column=0, sticky="ew")
        table_frame = ttk.Frame(frame, style="Card.TFrame")
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        columns = [column for column, _, _ in SELECTED_DATA_COLUMNS]
        self.selected_data_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
            height=16,
        )
        for column, title, width in SELECTED_DATA_COLUMNS:
            self.selected_data_tree.heading(column, text=title)
            self.selected_data_tree.column(column, width=width, minwidth=width, anchor="center", stretch=False)
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal")
        self.selected_data_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.selected_change_overlay = TreeviewCellColorOverlay(
            self.selected_data_tree,
            ("asian_change", "total_change"),
            self.handicap_change_text_color,
            lambda _value: self.handicap_change_background_color(),
        )
        y_scroll.configure(command=self.selected_change_overlay.yview)
        x_scroll.configure(command=self.selected_change_overlay.xview)
        self.selected_data_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        return frame

    def choose_output_dir(self) -> None:
        directory = filedialog.askdirectory(initialdir=self.output_dir_var.get() or str(default_output_dir()))
        if directory:
            self.output_dir_var.set(directory)

    def choose_handicap_bg_color(self) -> None:
        _rgb, color = colorchooser.askcolor(color=self.handicap_change_background_color(), parent=self.root)
        if color:
            self.handicap_bg_var.set(color.upper())

    def choose_handicap_low_text_color(self) -> None:
        _rgb, color = colorchooser.askcolor(color=self.handicap_change_low_text_color(), parent=self.root)
        if color:
            self.handicap_low_text_color_var.set(color.upper())

    def choose_handicap_high_text_color(self) -> None:
        _rgb, color = colorchooser.askcolor(color=self.handicap_change_high_text_color(), parent=self.root)
        if color:
            self.handicap_high_text_color_var.set(color.upper())

    def handicap_change_background_color(self) -> str:
        return normalized_hex_color(self.handicap_bg_var.get(), DEFAULT_HANDICAP_CHANGE_BACKGROUND)

    def handicap_change_low_text_color(self) -> str:
        return normalized_hex_color(self.handicap_low_text_color_var.get(), HANDICAP_ALERT_BLUE)

    def handicap_change_high_text_color(self) -> str:
        return normalized_hex_color(self.handicap_high_text_color_var.get(), HANDICAP_ALERT_RED)

    def handicap_change_text_color(self, value: object | None) -> str:
        return handicap_change_count_color_with_palette(
            value,
            self.handicap_change_low_text_color(),
            self.handicap_change_high_text_color(),
        )

    def refresh_handicap_change_overlays(self) -> None:
        if hasattr(self, "match_change_overlay"):
            self.match_change_overlay.schedule_refresh()
        if hasattr(self, "selected_change_overlay"):
            self.selected_change_overlay.schedule_refresh()

    def load_today_matches(self) -> None:
        if self.match_loader and self.match_loader.is_alive():
            return
        self.load_matches_button.configure(state="disabled")
        self.monitor_status_var.set("正在加载可选赛事")
        self.match_loader = threading.Thread(target=self.run_load_today_matches, daemon=True)
        self.match_loader.start()

    def run_load_today_matches(self) -> None:
        try:
            config = self.current_config()
            excel_matches: list[MatchInfo] = []
            try:
                excel_matches = read_selection_matches_from_workbook(config.resolve_output_path())
            except Exception as exc:
                self.messages.put(("monitor_log", f"本地 Excel 赛事读取失败：{exc}"))

            try:
                snapshot = fetch_company_schedule_snapshot(company_id=config.company_id)
                web_matches = filter_matches_by_date_window(snapshot.matches, snapshot.schedule_ids)
            except Exception:
                if not excel_matches:
                    raise
                web_matches = []
                self.messages.put(("monitor_log", "网页赛事加载失败，已使用本地 Excel 可识别赛事"))

            matches = merge_selection_matches(web_matches, excel_matches)
            self.messages.put(("today_matches", matches))
        except Exception as exc:
            self.messages.put(("monitor_error", exc))
        finally:
            self.messages.put(("matches_idle", None))

    def apply_today_matches(self, matches: list[MatchInfo]) -> None:
        self.today_matches = {match.schedule_id: match for match in matches}
        self.auto_monitor_alert_ids = auto_monitor_ids_for_today(matches)
        self.monitor_alert_ids = [
            schedule_id
            for schedule_id in self.monitor_alert_ids
            if (match := self.today_matches.get(schedule_id)) is not None and keep_manual_monitor_match(match)
        ]
        visible_matches = [match for match in matches if is_pending_match(match)]
        self.clear_selected_data_table()
        self.populate_match_tree(visible_matches)
        self.refresh_selected_match_table()
        self.monitor_status_var.set(f"已加载 {len(visible_matches)} 场可选赛事")
        self.append_log(f"可选赛事加载完成：显示未开始赛事 {len(visible_matches)} 场；隐藏已开赛/结束赛事 {len(matches) - len(visible_matches)} 场")

    def populate_match_tree(self, matches: list[MatchInfo]) -> None:
        checked_ids = set(self.checked_schedule_ids())
        for item in self.match_tree.get_children():
            self.match_tree.delete(item)
        for match in matches:
            self.match_tree.insert(
                "",
                "end",
                iid=match.schedule_id,
                values=(
                    CHECKED_MARK if match.schedule_id in checked_ids else UNCHECKED_MARK,
                    match.event_time,
                    match.league,
                    format_team_with_rank(match.home_team, match.home_rank),
                    match.score,
                    format_team_with_rank(match.away_team, match.away_rank),
                    "",
                    "",
                    "",
                    "",
                ),
            )
        self.match_change_overlay.schedule_refresh()

    def on_match_tree_click(self, event) -> str | None:
        if self.match_tree.identify_column(event.x) != "#1":
            return None
        schedule_id = self.match_tree.identify_row(event.y)
        if schedule_id:
            self.toggle_match_check(schedule_id)
        return "break"

    def toggle_match_check(self, schedule_id: str) -> None:
        values = list(self.match_tree.item(schedule_id, "values"))
        if not values:
            return
        values[0] = UNCHECKED_MARK if values[0] == CHECKED_MARK else CHECKED_MARK
        self.match_tree.item(schedule_id, values=values)

    def checked_schedule_ids(self) -> list[str]:
        checked_ids = []
        for schedule_id in self.match_tree.get_children():
            values = self.match_tree.item(schedule_id, "values")
            if values and values[0] == CHECKED_MARK:
                checked_ids.append(schedule_id)
        return checked_ids

    def add_checked_matches_to_monitor(self) -> None:
        self.monitor_alert_ids = [
            schedule_id
            for schedule_id in self.checked_schedule_ids()
            if (match := self.today_matches.get(schedule_id)) is not None and keep_manual_monitor_match(match)
        ]
        self.populate_match_tree([match for match in self.today_matches.values() if is_pending_match(match)])
        self.refresh_selected_match_table()
        active_count = len(self.active_monitor_alert_ids())
        if active_count:
            self.monitor_status_var.set(f"已添加 {active_count} 场提示赛事")
            self.append_log(f"提示赛事：自动 {len(self.auto_monitor_alert_ids)} 场，手动 {len(self.monitor_alert_ids)} 场")
        else:
            self.monitor_status_var.set("未添加提示赛事")
            self.append_log("未添加提示赛事：范围赛事仍可同步更新，但不会播放提示音")

    def active_monitor_alert_ids(self) -> list[str]:
        return merge_monitor_alert_ids(self.auto_monitor_alert_ids, self.monitor_alert_ids)

    def start_monitor(self) -> None:
        if self.monitor_worker and self.monitor_worker.is_alive():
            return
        alert_ids = list(self.monitor_alert_ids)
        active_alert_ids = self.active_monitor_alert_ids()
        if not self.today_matches:
            messagebox.showinfo("请先加载赛事", "请先加载可选赛事，再开始监控")
            return
        if not active_alert_ids:
            self.append_log("未选择提示赛事：范围赛事仍会同步更新，但不会播放提示音")
        self.monitor_cancel_event.clear()
        self.set_monitor_running(True)
        config = self.current_config()
        self.monitor_worker = threading.Thread(target=self.run_monitor, args=(config, alert_ids), daemon=True)
        self.monitor_worker.start()

    def run_monitor(self, config: GuiConfig, alert_ids: list[str]) -> None:
        previous_signatures = None
        previous_handicap_signatures = None
        cycle = 0
        output_path = config.resolve_output_path()
        try:
            while not self.monitor_cancel_event.is_set():
                cycle += 1
                snapshot = fetch_company_schedule_snapshot(company_id=config.company_id)
                range_matches = filter_matches_by_date_window(snapshot.matches, snapshot.schedule_ids)
                self.auto_monitor_alert_ids = auto_monitor_ids_for_today(range_matches)
                active_alert_ids = merge_monitor_alert_ids(self.auto_monitor_alert_ids, alert_ids)
                monitor_ids = [match.schedule_id for match in range_matches]
                self.messages.put(("monitor_log", f"开始第 {cycle} 次监控检查，范围赛事 {len(monitor_ids)} 场，提示赛事 {len(active_alert_ids)} 场"))
                records = crawl_complete_company_schedule_records(
                    company_id=config.company_id,
                    selected_schedule_ids=monitor_ids,
                    schedule_snapshot=snapshot,
                    workers=config.workers,
                    log_to_console=False,
                )
                sync_workbook(records, output_path, template_path=config.template_path)
                display_records = read_export_records_from_workbook(output_path) or records
                self.messages.put(("selected_records", (display_records, active_alert_ids)))
                self.messages.put(("monitor_synced", output_path))

                if previous_signatures is None:
                    self.messages.put(("monitor_log", "监控基准已建立，后续变化会提示"))
                else:
                    updated_ids = changed_schedule_ids(previous_signatures, records)
                    handicap_updated_ids = changed_handicap_schedule_ids(previous_handicap_signatures or {}, records)
                    alert_updates = alert_schedule_ids(handicap_updated_ids, active_alert_ids)
                    alert_details = handicap_alert_details(records, alert_updates)
                    if alert_details:
                        self.messages.put(("monitor_alert", alert_details))
                    elif updated_ids:
                        self.messages.put(("monitor_log", f"已同步更新 {len(updated_ids)} 场，未触发盘口变动提示音"))
                previous_signatures = record_signatures_by_schedule(records)
                previous_handicap_signatures = handicap_change_signatures_by_schedule(records)

                delay = config.next_monitor_delay_seconds()
                self.messages.put(("monitor_log", f"下次检查约 {format_delay(delay)} 后执行"))
                if self.monitor_cancel_event.wait(delay):
                    break
        except Exception as exc:
            self.messages.put(("monitor_error", exc))
        finally:
            self.messages.put(("monitor_idle", None))

    def stop_monitor(self) -> None:
        self.monitor_cancel_event.set()
        self.monitor_status_var.set("正在停止监控")

    def current_config(self) -> GuiConfig:
        return GuiConfig(
            output_dir=Path(self.output_dir_var.get() or default_output_dir()),
            monitor_interval_minutes=parse_float_setting(
                self.interval_minutes_var.get(),
                DEFAULT_MONITOR_INTERVAL_MINUTES,
                0.01,
            ),
            random_delay_seconds=parse_float_setting(
                self.random_delay_seconds_var.get(),
                DEFAULT_RANDOM_DELAY_SECONDS,
                0.0,
            ),
        )

    def start_export(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        self.cancel_event.clear()
        self.set_running(True)
        self.reset_output()
        config = self.current_config()
        self.worker = threading.Thread(target=self.run_export, args=(config,), daemon=True)
        self.worker.start()

    def start_history_export(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        try:
            history_date = parse_history_date_input(self.history_date_var.get())
        except ValueError as exc:
            messagebox.showerror("日期错误", str(exc))
            return
        self.cancel_event.clear()
        self.set_running(True)
        self.reset_output()
        config = self.current_config()
        self.worker = threading.Thread(target=self.run_history_export, args=(config, history_date, self.history_week_var.get()), daemon=True)
        self.worker.start()

    def run_export(self, config: GuiConfig) -> None:
        try:
            records = crawl_complete_company_schedule_records(
                company_id=config.company_id,
                workers=config.workers,
                progress_callback=lambda event: self.messages.put(("event", event)),
                cancel_event=self.cancel_event,
                log_to_console=False,
            )
            output_path = config.resolve_output_path()
            sync_workbook(records, output_path, template_path=config.template_path)
            display_records = read_export_records_from_workbook(output_path) or records
            self.messages.put(("selected_records", (display_records, self.active_monitor_alert_ids())))
            self.messages.put(("done", output_path))
        except CrawlCancelled:
            self.messages.put(("event", CrawlEvent("cancelled")))
            self.messages.put(("idle", None))
        except Exception as exc:
            self.messages.put(("fatal", exc))

    def run_history_export(self, config: GuiConfig, history_date: date, weekly: bool = False) -> None:
        try:
            records: list[ExportRecord] = []
            dates = [history_date + timedelta(days=offset) for offset in range(7)] if weekly else [history_date]
            for index, target_date in enumerate(dates, start=1):
                self.messages.put(("event", CrawlEvent("message", message=f"开始抓取历史日期 {target_date:%Y-%m-%d} ({index}/{len(dates)})")))
                records.extend(
                    crawl_historical_date_records(
                        target_date,
                        company_id=config.company_id,
                        workers=config.workers,
                        progress_callback=lambda event: self.messages.put(("event", event)),
                        cancel_event=self.cancel_event,
                        log_to_console=False,
                    )
                )
            output_path = historical_week_output_path(config.output_dir, history_date) if weekly else historical_output_path(config.output_dir, history_date)
            sync_workbook(records, output_path, template_path=config.template_path)
            display_records = read_export_records_from_workbook(output_path) or records
            self.messages.put(("selected_records", (display_records, [record.schedule_id for record in records if record.row_type == "match"])))
            self.messages.put(("done", output_path))
        except CrawlCancelled:
            self.messages.put(("event", CrawlEvent("cancelled")))
            self.messages.put(("idle", None))
        except Exception as exc:
            self.messages.put(("fatal", exc))

    def cancel_export(self) -> None:
        self.cancel_event.set()
        self.status_var.set("取消中")
        self.append_log("正在请求取消，当前已发出的请求会先结束")

    def drain_messages(self) -> None:
        while True:
            try:
                kind, payload = self.messages.get_nowait()
            except queue.Empty:
                break
            if kind == "event":
                self.apply_event(payload)
            elif kind == "done":
                self.apply_done(payload)
            elif kind == "fatal":
                self.apply_fatal(payload)
            elif kind == "today_matches":
                self.apply_today_matches(payload)
            elif kind == "monitor_log":
                self.append_log(payload)
            elif kind == "monitor_synced":
                self.apply_monitor_synced(payload)
            elif kind == "monitor_alert":
                self.apply_monitor_alert(payload)
            elif kind == "selected_records":
                records, selected_ids = payload
                self.apply_selected_records(records, selected_ids)
            elif kind == "monitor_error":
                self.apply_monitor_error(payload)
            elif kind == "matches_idle":
                self.load_matches_button.configure(state="normal")
            elif kind == "monitor_idle":
                self.set_monitor_running(False)
            elif kind == "idle":
                self.set_running(False)
        self.root.after(100, self.drain_messages)

    def apply_event(self, event: CrawlEvent) -> None:
        self.append_log(self.formatter.log_text(event))
        self.status_var.set(self.formatter.status_text(event))
        if event.type == "start":
            self.matches_var.set(str(event.total or 0))
            self.rows_var.set("-")
            self.progress_var.set(0)
            self.progress_label.configure(text=f"0 / {event.total or 0}")
        elif event.type == "match_done" and event.total:
            self.progress_var.set(((event.completed or 0) / event.total) * 100)
            self.progress_label.configure(text=f"{event.completed} / {event.total}")
        elif event.type == "error":
            current = int(self.failures_var.get() or "0")
            self.failures_var.set(str(current + 1))
        elif event.type == "complete":
            self.rows_var.set(str(event.rows or 0))
            self.progress_var.set(100)
            self.progress_label.configure(text=f"{event.total or 0} / {event.total or 0}")
        elif event.type == "cancelled":
            self.progress_label.configure(text="已取消")

    def apply_done(self, output_path: Path) -> None:
        self.last_output_path = output_path
        self.status_var.set("完成")
        self.append_log(f"Excel 写入完成：{output_path}")
        self.set_running(False)

    def apply_monitor_synced(self, output_path: Path) -> None:
        self.last_output_path = output_path
        self.append_log(f"监控同步：{output_path}")

    def apply_selected_records(self, records: list[ExportRecord], selected_ids: list[str]) -> None:
        self.update_match_change_counts(records)
        rows = selected_match_display_rows(records, selected_ids)
        if rows:
            self.populate_selected_data_table(rows)
        elif not selected_ids:
            self.refresh_selected_match_table()
        else:
            self.append_log("未找到选中赛事的同步记录，保留当前选中赛事数据")

    def update_match_change_counts(self, records: list[ExportRecord]) -> None:
        for schedule_id, display_values in match_change_count_display_values(records).items():
            if not self.match_tree.exists(schedule_id):
                continue
            values = list(self.match_tree.item(schedule_id, "values"))
            while len(values) < 10:
                values.append("")
            values[6], values[7], values[8], values[9] = display_values
            self.match_tree.item(schedule_id, values=values, tags=())
        self.match_change_overlay.schedule_refresh()

    def refresh_selected_match_table(self) -> None:
        self.populate_selected_data_table(selected_match_info_display_rows(self.today_matches, self.active_monitor_alert_ids()))

    def populate_selected_data_table(self, rows: list[tuple[str, ...]]) -> None:
        self.clear_selected_data_table()
        for row in rows:
            self.selected_data_tree.insert("", "end", values=row, tags=())
        self.selected_change_overlay.schedule_refresh()

    def clear_selected_data_table(self) -> None:
        for item in self.selected_data_tree.get_children():
            self.selected_data_tree.delete(item)
        self.selected_change_overlay.schedule_refresh()

    def apply_monitor_alert(self, details: list[dict[str, int | str]]) -> None:
        labels = [self.describe_match(str(detail.get("schedule_id", ""))) for detail in details]
        message = "监控更新：" + "；".join(labels)
        self.monitor_status_var.set(f"发现更新 {len(details)} 场")
        self.append_log(message)
        self.show_monitor_alert_popup(details)
        self.play_alert_sound()

    def show_monitor_alert_popup(self, details: list[dict[str, int | str]]) -> None:
        if not details:
            return
        popup = tk.Toplevel(self.root)
        popup.title("盘口变动提醒")
        popup.transient(self.root)
        popup.resizable(False, False)

        container = ttk.Frame(popup, padding=16)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="盘口变动提醒", font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w")
        ttk.Label(container, text=f"当前时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", foreground=MUTED).pack(anchor="w", pady=(4, 10))

        for index, detail in enumerate(details):
            schedule_id = str(detail.get("schedule_id", ""))
            match = self.today_matches.get(schedule_id)
            title = self.describe_match(schedule_id)
            event_time = match.event_time if match else ""
            status = match.status if match else ""
            if index:
                ttk.Separator(container).pack(fill="x", pady=10)
            ttk.Label(container, text=title, font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
            ttk.Label(container, text=f"比赛时间：{event_time}    状态：{status}", foreground=MUTED).pack(anchor="w", pady=(2, 6))
            counts = ttk.Frame(container)
            counts.pack(anchor="w")
            ttk.Label(counts, text="亚盘盘口变动：").pack(side="left")
            asian_count = str(detail.get("asian_count", 0))
            ttk.Label(counts, text=asian_count, foreground=self.handicap_change_text_color(asian_count), font=("Microsoft YaHei UI", 10, "bold")).pack(side="left")
            ttk.Label(counts, text="    大小球盘口变动：").pack(side="left")
            total_count = str(detail.get("total_count", 0))
            ttk.Label(counts, text=total_count, foreground=self.handicap_change_text_color(total_count), font=("Microsoft YaHei UI", 10, "bold")).pack(side="left")

        ttk.Button(container, text="知道了", command=popup.destroy).pack(anchor="e", pady=(14, 0))
        popup.lift()

    def apply_monitor_error(self, exc: Exception) -> None:
        self.monitor_status_var.set("监控异常")
        self.append_log(f"监控异常：{exc}")

    def apply_fatal(self, exc: Exception) -> None:
        self.status_var.set("失败")
        self.append_log(f"任务失败：{exc}")
        messagebox.showerror("导出失败", str(exc))
        self.set_running(False)

    def set_running(self, running: bool) -> None:
        self.start_button.configure(state="disabled" if running else "normal")
        self.cancel_button.configure(state="normal" if running else "disabled")

    def set_monitor_running(self, running: bool) -> None:
        self.load_matches_button.configure(state="disabled" if running else "normal")
        self.add_monitor_button.configure(state="disabled" if running else "normal")
        self.monitor_button.configure(state="disabled" if running else "normal")
        self.stop_monitor_button.configure(state="normal" if running else "disabled")
        if running:
            self.monitor_status_var.set("监控中")
        elif self.monitor_status_var.get() == "正在停止监控":
            self.monitor_status_var.set("已停止")

    def reset_output(self) -> None:
        self.status_var.set("运行中")
        self.matches_var.set("-")
        self.rows_var.set("-")
        self.failures_var.set("0")
        self.progress_var.set(0)
        self.progress_label.configure(text="等待数据")
        self.clear_text(self.log_text)

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.append_text(self.log_text, f"{timestamp}  {message}\n")

    def append_text(self, widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.insert("end", text)
        widget.see("end")
        widget.configure(state="disabled")

    def clear_text(self, widget: tk.Text) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    def open_output_dir(self) -> None:
        path = Path(self.output_dir_var.get() or default_output_dir())
        path.mkdir(parents=True, exist_ok=True)
        open_path(path)

    def open_latest_file(self) -> None:
        if self.last_output_path and self.last_output_path.exists():
            open_path(self.last_output_path)
            return
        messagebox.showinfo("没有文件", "当前还没有可打开的导出文件")

    def describe_match(self, schedule_id: str) -> str:
        match = self.today_matches.get(schedule_id)
        if match is None:
            return schedule_id
        return f"{match.league} {match.home_team} vs {match.away_team}"

    def play_alert_sound(self) -> None:
        if sys.platform.startswith("win"):
            try:
                import winsound

                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
                return
            except Exception:
                pass
        self.root.bell()


def open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def main() -> None:
    # try:
    #     enforce_startup_time_limit()
    # except StartupBlocked as exc:
    #     root = tk.Tk()
    #     root.withdraw()
    #     messagebox.showerror("启动失败", str(exc))
    #     root.destroy()
    #     return

    root = tk.Tk()
    Titan007ExporterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
