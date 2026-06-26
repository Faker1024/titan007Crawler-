import tempfile
import sys
import tkinter as tk
from tkinter import ttk
import unittest
from datetime import datetime, timedelta
from pathlib import Path

import gui as gui_module
from main import DEFAULT_TEMPLATE, CrawlCancelled, CrawlEvent, CurrentOdds, ExportRecord, MatchInfo, OddsRecord, SYNC_DEFAULT_OUTPUT, ScheduleSnapshot, combine_export_records, crawl_complete_company_schedule, sync_workbook
from gui import CHECKED_MARK, DEFAULT_HANDICAP_CHANGE_BACKGROUND, HANDICAP_ALERT_BLUE, HANDICAP_ALERT_RED, UNCHECKED_MARK, GuiConfig, GuiEventFormatter, Titan007ExporterApp, alert_schedule_ids, default_output_dir, handicap_alert_details, handicap_change_count_color, match_change_count_display_values, merge_selection_matches, parse_history_date_input, selected_match_display_rows, selected_match_info_display_rows


class GuiSupportTests(unittest.TestCase):
    def make_app(self):
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk is not available: {exc}")
        root.withdraw()
        self.addCleanup(root.destroy)
        return Titan007ExporterApp(root)

    def iter_widgets(self, widget):
        for child in widget.winfo_children():
            yield child
            yield from self.iter_widgets(child)

    def widget_texts(self, widget):
        texts = []
        for child in self.iter_widgets(widget):
            try:
                text = child.cget("text")
            except tk.TclError:
                continue
            if text:
                texts.append(text)
        return texts

    def test_gui_config_uses_explicit_output_path(self):
        config = GuiConfig(output_dir=Path("out"), output_name="custom.xlsx")

        self.assertEqual(config.resolve_output_path(now=datetime(2026, 5, 19, 14, 50, 1)), Path("out") / "custom.xlsx")

    def test_gui_config_uses_fixed_sync_default_output_path(self):
        config = GuiConfig(output_dir=Path("out"), output_name="")

        self.assertEqual(config.resolve_output_path(now=datetime(2026, 5, 19, 14, 50, 1)), Path("out") / SYNC_DEFAULT_OUTPUT)

    def test_gui_config_defaults_to_fixed_sync_workbook(self):
        config = GuiConfig()

        self.assertEqual(config.output_name, SYNC_DEFAULT_OUTPUT)
        self.assertEqual(config.template_path, Path(DEFAULT_TEMPLATE))
        self.assertEqual(config.monitor_interval_minutes, 1.0)
        self.assertEqual(config.random_delay_seconds, 10.0)
        self.assertEqual(config.resolve_output_path(now=datetime(2026, 5, 19, 14, 50, 1)), Path.cwd() / SYNC_DEFAULT_OUTPUT)

    def test_default_output_dir_uses_exe_directory_when_frozen(self):
        original_frozen = getattr(sys, "frozen", None)
        had_frozen = hasattr(sys, "frozen")
        original_executable = sys.executable
        try:
            sys.frozen = True
            sys.executable = str(Path("C:/Tools/Titan007Exporter.exe"))

            self.assertEqual(default_output_dir(), Path("C:/Tools"))
            self.assertEqual(GuiConfig().output_dir, Path("C:/Tools"))
        finally:
            sys.executable = original_executable
            if had_frozen:
                sys.frozen = original_frozen
            else:
                delattr(sys, "frozen")

    def test_parse_history_date_input_accepts_compact_and_dash_formats(self):
        self.assertEqual(parse_history_date_input("20260526"), datetime(2026, 5, 26).date())
        self.assertEqual(parse_history_date_input("2026-05-26"), datetime(2026, 5, 26).date())

    def test_config_panel_exposes_only_actionable_inputs(self):
        app = self.make_app()

        entries = [widget for widget in self.iter_widgets(app.root) if isinstance(widget, ttk.Entry)]
        texts = self.widget_texts(app.root)

        self.assertEqual(len(entries), 7)
        self.assertIn("\u52a0\u8f7d\u53ef\u9009\u8d5b\u4e8b", texts)
        self.assertNotIn("\u52a0\u8f7d\u524d\u540e7\u5929\u8d5b\u4e8b", texts)
        self.assertNotIn("\u52a0\u8f7d\u4eca\u65e5\u8d5b\u4e8b", texts)
        self.assertIn("保存目录", texts)
        self.assertIn("历史日期", texts)
        self.assertIn("抓取指定日期", texts)
        self.assertIn("间隔(分钟)", texts)
        self.assertIn("随机延迟(秒)", texts)
        self.assertIn("添加赛事监控", texts)
        self.assertIn("启动监控", texts)
        self.assertFalse(hasattr(app, "template_var"))
        self.assertFalse(hasattr(app, "output_name_var"))

    def test_handicap_change_background_color_is_customizable(self):
        app = self.make_app()

        self.assertEqual(app.handicap_change_background_color(), DEFAULT_HANDICAP_CHANGE_BACKGROUND)
        app.handicap_bg_var.set("#fee2e2")
        self.assertEqual(app.handicap_change_background_color(), "#FEE2E2")
        app.handicap_bg_var.set("bad-color")
        self.assertEqual(app.handicap_change_background_color(), DEFAULT_HANDICAP_CHANGE_BACKGROUND)

    def test_handicap_change_text_colors_are_customizable(self):
        app = self.make_app()

        self.assertEqual(app.handicap_change_text_color("3"), HANDICAP_ALERT_BLUE)
        self.assertEqual(app.handicap_change_text_color("4"), HANDICAP_ALERT_RED)

        app.handicap_low_text_color_var.set("#047857")
        app.handicap_high_text_color_var.set("#7f1d1d")
        self.assertEqual(app.handicap_change_text_color("3"), "#047857")
        self.assertEqual(app.handicap_change_text_color("4"), "#7F1D1D")

        app.handicap_low_text_color_var.set("bad-color")
        app.handicap_high_text_color_var.set("bad-color")
        self.assertEqual(app.handicap_change_text_color("3"), HANDICAP_ALERT_BLUE)
        self.assertEqual(app.handicap_change_text_color("4"), HANDICAP_ALERT_RED)

    def test_header_subtitle_has_enough_height_for_complete_match_label(self):
        app = self.make_app()
        app.root.deiconify()
        app.root.update()
        subtitle = None
        for widget in self.iter_widgets(app.root):
            try:
                text = widget.cget("text")
            except tk.TclError:
                continue
            if text == "\u5b8c\u6574\u8d5b\u4e8b \u00b7 \u6fb3*\u76d8\u53e3 \u00b7 Excel \u5bfc\u51fa":
                subtitle = widget
                break
        self.assertIsNotNone(subtitle)

        self.assertGreaterEqual(subtitle.winfo_height(), subtitle.winfo_reqheight())

    def test_config_panel_keeps_match_action_buttons_visible_at_default_size(self):
        app = self.make_app()
        app.root.deiconify()
        app.root.update()

        for button in (
            app.load_matches_button,
            app.add_monitor_button,
            app.monitor_button,
            app.stop_monitor_button,
        ):
            self.assertGreaterEqual(button.winfo_height(), button.winfo_reqheight())

    def test_run_panel_uses_tabs_for_progress_and_selected_match_data(self):
        app = self.make_app()

        notebooks = [widget for widget in self.iter_widgets(app.root) if isinstance(widget, ttk.Notebook)]

        self.assertEqual(len(notebooks), 1)
        self.assertEqual(
            [notebooks[0].tab(tab_id, "text") for tab_id in notebooks[0].tabs()],
            ["\u8fd0\u884c\u8fdb\u5ea6", "\u9009\u4e2d\u8d5b\u4e8b\u6570\u636e"],
        )

    def test_current_config_reads_monitor_timing_settings(self):
        app = self.make_app()
        app.interval_minutes_var.set("2.5")
        app.random_delay_seconds_var.set("12")

        config = app.current_config()

        self.assertEqual(config.monitor_interval_minutes, 2.5)
        self.assertEqual(config.random_delay_seconds, 12.0)
        self.assertEqual(config.next_monitor_delay_seconds(random_value=0.5), 156.0)

    def test_monitor_crawls_all_range_matches_even_when_one_match_is_selected_for_alerts(self):
        app = self.make_app()
        called_schedule_ids = []
        today = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        first = MatchInfo("1001", "联赛A", "今日19:30", "主队A", "-", "客队A", "-", "2.5", today, status="未")
        second = MatchInfo("1002", "联赛B", "今日20:30", "主队B", "-", "客队B", "-", "2.5", today.replace(hour=20), status="未")

        first = MatchInfo("1001", "A", "5-16 19:30", "H1", "-", "A1", "-", "2.5", today - timedelta(days=7), status="未")
        second = MatchInfo("1002", "B", "5-30 20:30", "H2", "-", "A2", "-", "2.5", today + timedelta(days=7), status="未")
        outside = MatchInfo("1003", "C", "5-31 20:30", "H3", "-", "A3", "-", "2.5", today + timedelta(days=8), status="未")

        class StopAfterOneCycle:
            def is_set(self):
                return False

            def wait(self, delay):
                return True

        def fake_snapshot(company_id=1):
            return ScheduleSnapshot(matches={"1001": first, "1002": second, "1003": outside}, schedule_ids=["1003", "1001", "1002"])

        def fake_crawl(**kwargs):
            called_schedule_ids.extend(kwargs["selected_schedule_ids"])
            return []

        original_snapshot = gui_module.fetch_company_schedule_snapshot
        original_crawl = gui_module.crawl_complete_company_schedule_records
        original_sync = gui_module.sync_workbook
        try:
            gui_module.fetch_company_schedule_snapshot = fake_snapshot
            gui_module.crawl_complete_company_schedule_records = fake_crawl
            gui_module.sync_workbook = lambda records, output_path, template_path=None: output_path
            app.monitor_cancel_event = StopAfterOneCycle()

            app.run_monitor(GuiConfig(), ["1001"])
        finally:
            gui_module.fetch_company_schedule_snapshot = original_snapshot
            gui_module.crawl_complete_company_schedule_records = original_crawl
            gui_module.sync_workbook = original_sync

        self.assertEqual(called_schedule_ids, ["1001", "1002"])

    def test_alert_schedule_ids_only_keeps_selected_updated_matches(self):
        self.assertEqual(alert_schedule_ids(["1001", "1002", "1003"], ["1002", "1004"]), ["1002"])

    def test_handicap_alert_details_requires_change_count_at_least_three(self):
        match = MatchInfo("1001", "League", "今日19:30", "Home", "-", "Away", "-", "2.5", datetime.now(), status="未")
        low_records = combine_export_records(
            {"1001": match},
            ["1001"],
            {
                "1001": [
                    OddsRecord(datetime(2026, 6, 8, 12, 0), 0.82, "A", 1.0),
                    OddsRecord(datetime(2026, 6, 8, 13, 0), 0.84, "B", 0.98),
                    OddsRecord(datetime(2026, 6, 8, 14, 0), 0.86, "C", 0.96),
                ]
            },
            {},
        )
        high_records = combine_export_records(
            {"1001": match},
            ["1001"],
            {
                "1001": [
                    OddsRecord(datetime(2026, 6, 8, 12, 0), 0.82, "A", 1.0),
                    OddsRecord(datetime(2026, 6, 8, 13, 0), 0.84, "B", 0.98),
                    OddsRecord(datetime(2026, 6, 8, 14, 0), 0.86, "C", 0.96),
                    OddsRecord(datetime(2026, 6, 8, 15, 0), 0.88, "D", 0.94),
                ]
            },
            {},
        )

        self.assertEqual(handicap_alert_details(low_records, ["1001"]), [])
        self.assertEqual(
            handicap_alert_details(high_records, ["1001"]),
            [{"schedule_id": "1001", "asian_count": 3, "total_count": 0}],
        )

    def test_handicap_change_count_color_is_blue_until_four_then_red(self):
        self.assertEqual(handicap_change_count_color(""), "")
        self.assertEqual(handicap_change_count_color("0"), HANDICAP_ALERT_BLUE)
        self.assertEqual(handicap_change_count_color("3"), HANDICAP_ALERT_BLUE)
        self.assertEqual(handicap_change_count_color("4"), HANDICAP_ALERT_RED)

    def test_selected_match_display_rows_include_latest_odds_columns(self):
        selected = MatchInfo("1001", "League", "5-23 19:30", "Home", "1-0", "Away", "0-0", "2.5", datetime(2026, 5, 23, 19, 30), status="未")
        skipped = MatchInfo("1002", "Other", "5-23 20:30", "Skip", "-", "Away2", "-", "2.5", datetime(2026, 5, 23, 20, 30), status="未")
        records = combine_export_records(
            {"1001": selected, "1002": skipped},
            ["1001", "1002"],
            {"1001": [OddsRecord(datetime(2026, 5, 23, 12, 0), 0.82, "A", 1.0)]},
            {
                "1001": [
                    OddsRecord(datetime(2026, 5, 23, 12, 5), 0.9, "2.5", 0.92),
                    OddsRecord(datetime(2026, 5, 23, 13, 5), 0.91, "3", 0.88),
                ]
            },
            current_odds={"1001": CurrentOdds("A", 0.81, 1.01, "2.5", 0.89, 0.93)},
        )

        rows = selected_match_display_rows(records, ["1001"])

        self.assertEqual(
            rows,
            [
                (
                    "League",
                    "5-23 19:30",
                    selected.status,
                    "Home",
                    "1-0",
                    "Away",
                    "0.81",
                    "A",
                    "1.01",
                    "05-23 12:00",
                    "0.82",
                    "A",
                    "1",
                    "0",
                    "0",
                    "05-23 12:05",
                    "0.9",
                    "2.5",
                    "0.92",
                    "1",
                    "1",
                )
                ,
                (
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
                    "05-23 13:05",
                    "0.91",
                    "3",
                    "0.88",
                    "",
                    "",
                )
            ],
        )

    def test_selected_match_data_table_refreshes_from_crawl_records(self):
        app = self.make_app()
        selected = MatchInfo("1001", "League", "5-23 19:30", "Home", "1-0", "Away", "0-0", "2.5", datetime(2026, 5, 23, 19, 30), status="未")
        records = combine_export_records(
            {"1001": selected},
            ["1001"],
            {"1001": [OddsRecord(datetime(2026, 5, 23, 12, 0), 0.82, "A", 1.0)]},
            {
                "1001": [
                    OddsRecord(datetime(2026, 5, 23, 12, 5), 0.9, "2.5", 0.92),
                    OddsRecord(datetime(2026, 5, 23, 13, 5), 0.91, "3", 0.88),
                ]
            },
            current_odds={"1001": CurrentOdds("A", 0.81, 1.01, "2.5", 0.89, 0.93)},
        )

        app.apply_selected_records(records, ["1001"])

        children = app.selected_data_tree.get_children()
        self.assertEqual(len(children), 2)
        self.assertEqual(app.selected_data_tree.item(children[0], "values")[9:21], ("05-23 12:00", "0.82", "A", "1", "0", "0", "05-23 12:05", "0.9", "2.5", "0.92", "1", "1"))
        self.assertEqual(app.selected_data_tree.item(children[1], "values")[15:21], ("05-23 13:05", "0.91", "3", "0.88", "", ""))

    def test_selectable_matches_use_web_source_before_excel_fallback(self):
        web = MatchInfo("1001", "A", "5-23 19:30", "H1", "-", "A1", "-", "2.5", datetime.now())
        excel_only = MatchInfo("1002", "B", "5-22 19:30", "H2", "-", "A2", "-", "2.5", datetime.now())

        self.assertEqual(merge_selection_matches([web], [excel_only]), [web])
        self.assertEqual(merge_selection_matches([], [excel_only]), [excel_only])

    def test_today_match_rows_use_checkbox_column_for_alert_selection(self):
        app = self.make_app()
        match = MatchInfo("1001", "联赛A", "今日19:30", "主队A", "-", "客队A", "-", "2.5", datetime.now(), status="未")

        app.apply_today_matches([match])

        self.assertEqual(app.match_tree.heading("odds_change_count")["text"], "赔率变动")
        self.assertEqual(app.match_tree.heading("handicap_change_count")["text"], "盘口变动")
        self.assertEqual(app.match_tree.heading("total_odds_change_count")["text"], "大小球赔率变动")
        self.assertEqual(app.match_tree.heading("total_handicap_change_count")["text"], "大小球盘口变动")
        self.assertEqual(app.match_tree.item("1001", "values")[0], UNCHECKED_MARK)
        self.assertEqual(app.match_tree.item("1001", "values")[6], "")
        self.assertEqual(app.match_tree.item("1001", "values")[7], "")
        self.assertEqual(app.match_tree.item("1001", "values")[8], "")
        self.assertEqual(app.match_tree.item("1001", "values")[9], "")
        app.toggle_match_check("1001")
        self.assertEqual(app.match_tree.item("1001", "values")[0], CHECKED_MARK)

    def test_today_matches_are_auto_monitored_without_checking_left_table(self):
        app = self.make_app()
        today = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        today_match = MatchInfo("1001", "联赛A", "今日19:30", "主队A", "-", "客队A", "-", "2.5", today, status="未")
        other_day = MatchInfo("1002", "联赛B", "明日19:30", "主队B", "-", "客队B", "-", "2.5", today + timedelta(days=1), status="未")

        app.apply_today_matches([today_match, other_day])

        self.assertEqual(app.auto_monitor_alert_ids, ["1001"])
        self.assertEqual(app.monitor_alert_ids, [])
        self.assertEqual(app.match_tree.item("1001", "values")[0], UNCHECKED_MARK)
        self.assertEqual(app.match_tree.item("1002", "values")[0], UNCHECKED_MARK)
        children = app.selected_data_tree.get_children()
        self.assertEqual(len(children), 1)
        self.assertEqual(app.selected_data_tree.item(children[0], "values")[:6], ("联赛A", "今日19:30", today_match.status, "主队A", "-", "客队A"))

    def test_selectable_match_table_shows_only_pending_matches(self):
        app = self.make_app()
        today = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        today_pending = MatchInfo("1001", "A", "今日19:30", "H1", "-", "A1", "-", "2.5", today, status="未")
        today_started = MatchInfo("1002", "B", "今日20:30", "H2", "0-0", "A2", "0-0", "2.5", today.replace(hour=20), status="即")
        future_unselected = MatchInfo("1003", "C", "明日19:30", "H3", "-", "A3", "-", "2.5", today + timedelta(days=1), status="未")

        app.apply_today_matches([today_pending, today_started, future_unselected])

        self.assertEqual(app.match_tree.get_children(), ("1001", "1003"))

    def test_started_today_match_is_not_auto_selected_or_left_selectable(self):
        app = self.make_app()
        today = datetime.now().replace(hour=20, minute=0, second=0, microsecond=0)
        started = MatchInfo("1001", "国际友谊", "今日20:00", "主队A", "0-0", "客队A", "0-0", "2.5", today, status="即")

        app.apply_today_matches([started])

        self.assertFalse(app.match_tree.exists("1001"))
        self.assertEqual(app.auto_monitor_alert_ids, [])
        children = app.selected_data_tree.get_children()
        self.assertEqual(len(children), 0)

    def test_loading_matches_removes_previous_day_from_auto_monitoring(self):
        app = self.make_app()
        now = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        old_today = MatchInfo("1001", "联赛A", "昨日19:30", "主队A", "-", "客队A", "-", "2.5", now, status="未")
        app.apply_today_matches([old_today])
        self.assertEqual(app.auto_monitor_alert_ids, ["1001"])

        yesterday = MatchInfo("1001", "联赛A", "昨日19:30", "主队A", "-", "客队A", "-", "2.5", now - timedelta(days=1), status="未")
        new_today = MatchInfo("1002", "联赛B", "今日19:30", "主队B", "-", "客队B", "-", "2.5", now, status="未")
        app.apply_today_matches([yesterday, new_today])

        self.assertEqual(app.auto_monitor_alert_ids, ["1002"])

    def test_selected_records_update_selectable_match_change_count_columns(self):
        app = self.make_app()
        match_time = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        match = MatchInfo("1001", "League", "今日19:30", "Home", "1-0", "Away", "0-0", "2.5", match_time, status="未")
        app.apply_today_matches([match])
        records = combine_export_records(
            {"1001": match},
            ["1001"],
            {
                "1001": [
                    OddsRecord(datetime(2026, 5, 23, 12, 0), 0.82, "A", 1.0),
                    OddsRecord(datetime(2026, 5, 23, 13, 0), 0.84, "B", 0.98),
                    OddsRecord(datetime(2026, 5, 23, 14, 0), 0.86, "C", 0.96),
                ]
            },
            {
                "1001": [
                    OddsRecord(datetime(2026, 5, 23, 12, 5), 0.9, "2.5", 0.92),
                    OddsRecord(datetime(2026, 5, 23, 13, 5), 0.91, "3", 0.88),
                ]
            },
            current_odds={"1001": CurrentOdds("A", 0.81, 1.01, "2.5", 0.89, 0.93)},
        )

        app.apply_selected_records(records, ["1001"])

        self.assertEqual(app.match_tree.item("1001", "values")[6:10], ("2", "2", "1", "1"))

    def test_selected_records_keeps_existing_table_when_selected_sync_rows_are_missing(self):
        app = self.make_app()
        app.populate_selected_data_table([("League", "今日20:00", "未", "Home", "-", "Away")])

        app.apply_selected_records([], ["1001"])

        children = app.selected_data_tree.get_children()
        self.assertEqual(len(children), 1)
        self.assertEqual(app.selected_data_tree.item(children[0], "values")[:6], ("League", "今日20:00", "未", "Home", "-", "Away"))

    def test_change_count_display_values_read_previous_rank_workbook_layout(self):
        old_row = (
            "League", "5-23 19:30", "", "", "Home", "1-0", "", "Away", "0.81", "A", "1.01",
            "05-23 12:00", "0.82", "A", "1", "2",
            "05-23 12:05", "0.9", "2.5", "0.92", "1",
        )
        records = [
            ExportRecord("1001", "1001|row|0", "match", 0, 0, old_row),
        ]

        self.assertEqual(match_change_count_display_values(records)["1001"], ("2", "0", "1", "0"))

    def test_change_count_display_values_keep_current_layout_without_total_rows(self):
        current_row = (
            "League", "5-23 19:30", "", "", "Home", "1-0", "", "Away", "0.81", "A", "1.01",
            "05-23 12:00", "0.82", "A", "1", "2", "2",
            "", "", "", "", "", "",
        )
        records = [
            ExportRecord("1001", "1001|row|0", "match", 0, 0, current_row),
        ]

        self.assertEqual(match_change_count_display_values(records)["1001"], ("2", "2", "0", "0"))

    def test_add_checked_matches_to_monitor_is_separate_from_starting_monitor(self):
        app = self.make_app()
        now = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        first = MatchInfo("1001", "联赛A", "今日19:30", "主队A", "-", "客队A", "-", "2.5", now, status="未")
        second = MatchInfo("1002", "联赛B", "今日20:30", "主队B", "-", "客队B", "-", "2.5", now.replace(hour=20), status="未")

        app.apply_today_matches([first, second])
        app.toggle_match_check("1001")
        app.add_checked_matches_to_monitor()
        app.toggle_match_check("1002")

        self.assertEqual(app.monitor_alert_ids, ["1001"])

    def test_add_checked_matches_keeps_today_and_future_manual_alerts(self):
        app = self.make_app()
        now = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        today_match = MatchInfo("1001", "联赛A", "今日19:30", "主队A", "-", "客队A", "-", "2.5", now, status="未")
        future_match = MatchInfo("1002", "联赛B", "明日19:30", "主队B", "-", "客队B", "-", "2.5", now + timedelta(days=1), status="未")

        app.monitor_alert_ids = ["1002"]
        app.apply_today_matches([today_match, future_match])
        app.toggle_match_check("1001")
        app.toggle_match_check("1002")
        app.add_checked_matches_to_monitor()

        self.assertEqual(app.auto_monitor_alert_ids, ["1001"])
        self.assertEqual(app.monitor_alert_ids, ["1001", "1002"])

    def test_selected_match_info_rows_match_selected_data_column_count(self):
        match = MatchInfo("1001", "League", "5-23 19:30", "Home", "1-0", "Away", "0-0", "2.5", datetime(2026, 5, 23, 19, 30), status="未")

        rows = selected_match_info_display_rows({"1001": match}, ["1001"])

        self.assertEqual(len(rows), 1)
        self.assertEqual(len(rows[0]), 21)

    def test_add_checked_matches_populates_selected_match_data_before_crawl(self):
        app = self.make_app()
        match = MatchInfo("1001", "联赛A", "今日19:30", "主队A", "-", "客队A", "-", "2.5", datetime.now(), status="未")

        app.apply_today_matches([match])
        app.toggle_match_check("1001")
        app.add_checked_matches_to_monitor()

        children = app.selected_data_tree.get_children()
        self.assertEqual(len(children), 1)
        self.assertEqual(app.selected_data_tree.item(children[0], "values")[:6], ("联赛A", "今日19:30", match.status, "主队A", "-", "客队A"))

    def test_start_monitor_uses_added_monitor_matches_not_current_checkbox_state(self):
        app = self.make_app()
        app.today_matches = {"1001": object(), "1002": object()}
        app.monitor_alert_ids = ["1001"]
        app.auto_monitor_alert_ids = ["1002"]
        captured = {}

        class FakeThread:
            def __init__(self, target, args=(), daemon=None):
                captured["args"] = args

            def start(self):
                captured["started"] = True

            def is_alive(self):
                return False

        original_thread = gui_module.threading.Thread
        try:
            gui_module.threading.Thread = FakeThread
            app.start_monitor()
        finally:
            gui_module.threading.Thread = original_thread

        self.assertTrue(captured["started"])
        self.assertEqual(captured["args"][1], ["1001"])

    def test_run_monitor_uses_auto_and_manual_alert_ids_for_selected_data_and_alerts(self):
        app = self.make_app()
        now = datetime.now().replace(hour=19, minute=30, second=0, microsecond=0)
        today_match = MatchInfo("1001", "A", "今日19:30", "H1", "-", "A1", "-", "2.5", now, status="未")
        manual_match = MatchInfo("1002", "B", "明日19:30", "H2", "-", "A2", "-", "2.5", now + timedelta(days=1), status="未")
        payloads = []

        class StopAfterOneCycle:
            def is_set(self):
                return False

            def wait(self, delay):
                return True

        def fake_snapshot(company_id=1):
            return ScheduleSnapshot(matches={"1001": today_match, "1002": manual_match}, schedule_ids=["1001", "1002"])

        def fake_crawl(**kwargs):
            return combine_export_records(
                {"1001": today_match, "1002": manual_match},
                ["1001", "1002"],
                {},
                {},
            )

        original_snapshot = gui_module.fetch_company_schedule_snapshot
        original_crawl = gui_module.crawl_complete_company_schedule_records
        original_sync = gui_module.sync_workbook
        original_put = app.messages.put
        try:
            gui_module.fetch_company_schedule_snapshot = fake_snapshot
            gui_module.crawl_complete_company_schedule_records = fake_crawl
            gui_module.sync_workbook = lambda records, output_path, template_path=None: output_path
            app.messages.put = lambda item: payloads.append(item)
            app.monitor_cancel_event = StopAfterOneCycle()

            app.run_monitor(GuiConfig(), ["1002"])
        finally:
            gui_module.fetch_company_schedule_snapshot = original_snapshot
            gui_module.crawl_complete_company_schedule_records = original_crawl
            gui_module.sync_workbook = original_sync
            app.messages.put = original_put

        selected_payload = next(payload for kind, payload in payloads if kind == "selected_records")
        self.assertEqual(selected_payload[1], ["1001", "1002"])

    def test_run_monitor_refreshes_selected_data_from_synced_workbook_records(self):
        app = self.make_app()
        match = MatchInfo("1001", "League", "今日20:00", "Home", "-", "Away", "-", "2.5", datetime.now().replace(hour=20, minute=0, second=0, microsecond=0), status="未")
        old_records = combine_export_records(
            {"1001": match},
            ["1001"],
            {
                "1001": [
                    OddsRecord(datetime(2026, 6, 3, 7, 0), 0.82, "A", 1.0),
                    OddsRecord(datetime(2026, 6, 3, 8, 0), 0.84, "B", 0.98),
                ]
            },
            {
                "1001": [
                    OddsRecord(datetime(2026, 6, 3, 7, 5), 0.9, "2.5", 0.92),
                    OddsRecord(datetime(2026, 6, 3, 8, 5), 0.91, "3", 0.88),
                ]
            },
        )
        new_records = combine_export_records({"1001": match}, ["1001"], {}, {})
        payloads = []

        class StopAfterOneCycle:
            def is_set(self):
                return False

            def wait(self, delay):
                return True

        def fake_snapshot(company_id=1):
            return ScheduleSnapshot(matches={"1001": match}, schedule_ids=["1001"])

        with tempfile.TemporaryDirectory() as tmp:
            output_path = Path(tmp) / SYNC_DEFAULT_OUTPUT
            sync_workbook(old_records, output_path, template_path=None)

            original_snapshot = gui_module.fetch_company_schedule_snapshot
            original_crawl = gui_module.crawl_complete_company_schedule_records
            original_put = app.messages.put
            try:
                gui_module.fetch_company_schedule_snapshot = fake_snapshot
                gui_module.crawl_complete_company_schedule_records = lambda **kwargs: new_records
                app.messages.put = lambda item: payloads.append(item)
                app.monitor_cancel_event = StopAfterOneCycle()

                app.run_monitor(GuiConfig(output_dir=Path(tmp)), ["1001"])
            finally:
                gui_module.fetch_company_schedule_snapshot = original_snapshot
                gui_module.crawl_complete_company_schedule_records = original_crawl
                app.messages.put = original_put

        selected_payload = next(payload for kind, payload in payloads if kind == "selected_records")
        display_records, selected_ids = selected_payload
        self.assertEqual(selected_ids, ["1001"])
        self.assertGreater(len(display_records), len(new_records))
        self.assertTrue(any(record.row_type == "detail" for record in display_records))

    def test_weekly_history_export_fetches_seven_days_into_one_workbook(self):
        app = self.make_app()
        calls = []
        saved = {}

        def fake_crawl(match_date, **kwargs):
            calls.append(match_date)
            match = MatchInfo(f"{match_date:%Y%m%d}", "League", match_date.strftime("%m-%d"), "Home", "1-0", "Away", "0-0", "2.5", datetime.combine(match_date, datetime.min.time()), status="完")
            return combine_export_records({match.schedule_id: match}, [match.schedule_id], {}, {})

        def fake_sync(records, output_path, template_path=None):
            saved["path"] = output_path
            saved["match_ids"] = [record.schedule_id for record in records if record.row_type == "match"]
            return output_path

        original_crawl = gui_module.crawl_historical_date_records
        original_sync = gui_module.sync_workbook
        original_read = gui_module.read_export_records_from_workbook
        original_put = app.messages.put
        try:
            gui_module.crawl_historical_date_records = fake_crawl
            gui_module.sync_workbook = fake_sync
            gui_module.read_export_records_from_workbook = lambda path: []
            app.messages.put = lambda item: None

            app.run_history_export(GuiConfig(output_dir=Path("out")), datetime(2026, 5, 26).date(), weekly=True)
        finally:
            gui_module.crawl_historical_date_records = original_crawl
            gui_module.sync_workbook = original_sync
            gui_module.read_export_records_from_workbook = original_read
            app.messages.put = original_put

        self.assertEqual(calls, [datetime(2026, 5, 26).date() + timedelta(days=offset) for offset in range(7)])
        self.assertEqual(saved["path"], Path("out") / "titan007_data_20260526_20260601.xlsx")
        self.assertEqual(saved["match_ids"], ["20260526", "20260527", "20260528", "20260529", "20260530", "20260531", "20260601"])

    def test_config_panel_hides_fixed_readonly_details(self):
        app = self.make_app()

        texts = self.widget_texts(app.root)

        for fixed_text in [
            "抓取配置",
            "赛事范围",
            "完整赛事",
            "盘口公司",
            "澳*（company_id=1）",
            "同步文件",
            SYNC_DEFAULT_OUTPUT,
            "参考模板",
            str(Path(DEFAULT_TEMPLATE)),
            "并发数",
        ]:
            self.assertNotIn(fixed_text, texts)

    def test_gui_event_formatter_formats_progress_and_complete_events(self):
        formatter = GuiEventFormatter()

        self.assertEqual(formatter.status_text(CrawlEvent("start", total=35)), "运行中")
        self.assertEqual(formatter.log_text(CrawlEvent("match_done", completed=3, total=35, schedule_id="2986245")), "已完成 3/35：2986245")
        self.assertEqual(formatter.status_text(CrawlEvent("complete", total=35, rows=149)), "完成")
        self.assertEqual(formatter.log_text(CrawlEvent("complete", total=35, rows=149)), "导出数据准备完成：35 场，149 行")

    def test_crawl_progress_callback_receives_start_match_done_and_complete(self):
        events = []

        def fake_fetch(session, schedule_id, match):
            return [], []

        rows = crawl_complete_company_schedule(
            limit=2,
            fetch_histories=fake_fetch,
            progress_callback=events.append,
            log_to_console=False,
        )

        self.assertTrue(rows)
        self.assertEqual(events[0].type, "start")
        self.assertEqual(events[0].total, 2)
        self.assertEqual([event.type for event in events if event.type == "match_done"], ["match_done", "match_done"])
        self.assertEqual(events[-1].type, "complete")

    def test_crawl_cancel_event_stops_before_fetching_matches(self):
        class CancelEvent:
            def is_set(self):
                return True

        with self.assertRaises(CrawlCancelled):
            crawl_complete_company_schedule(
                limit=1,
                fetch_histories=lambda session, schedule_id, match: ([], []),
                cancel_event=CancelEvent(),
                log_to_console=False,
            )


if __name__ == "__main__":
    unittest.main()
