import unittest
from datetime import datetime, timezone

import gui as gui_module
import main as crawler_main
from startup_guard import StartupBlocked, enforce_startup_time_limit, parse_http_date


class StartupGuardTests(unittest.TestCase):
    def test_allows_start_before_20260613_china_time(self):
        current_time = datetime(2026, 6, 12, 15, 59, 59, tzinfo=timezone.utc)

        self.assertEqual(enforce_startup_time_limit(now_fetcher=lambda: current_time), current_time)

    def test_blocks_start_at_20260613_china_time(self):
        current_time = datetime(2026, 6, 12, 16, 0, 0, tzinfo=timezone.utc)

        with self.assertRaises(StartupBlocked):
            enforce_startup_time_limit(now_fetcher=lambda: current_time)

    def test_blocks_start_when_network_time_is_unavailable(self):
        with self.assertRaises(StartupBlocked):
            enforce_startup_time_limit(now_fetcher=lambda: None)

    def test_parse_http_date_returns_timezone_aware_datetime(self):
        parsed = parse_http_date("Sat, 23 May 2026 15:59:59 GMT")

        self.assertEqual(parsed, datetime(2026, 5, 23, 15, 59, 59, tzinfo=timezone.utc))

    def test_cli_main_stops_before_parsing_args_when_startup_is_blocked(self):
        calls = []

        def blocked_guard():
            raise StartupBlocked("blocked")

        original_guard = crawler_main.enforce_startup_time_limit
        original_parse_args = crawler_main.parse_args
        try:
            crawler_main.enforce_startup_time_limit = blocked_guard
            crawler_main.parse_args = lambda: calls.append("parse_args")

            with self.assertRaises(SystemExit):
                crawler_main.main()
        finally:
            crawler_main.enforce_startup_time_limit = original_guard
            crawler_main.parse_args = original_parse_args

        self.assertEqual(calls, [])

    def test_gui_main_shows_error_and_does_not_build_app_when_startup_is_blocked(self):
        calls = []

        class FakeRoot:
            def withdraw(self):
                calls.append("withdraw")

            def destroy(self):
                calls.append("destroy")

        def blocked_guard():
            raise StartupBlocked("blocked")

        original_guard = gui_module.enforce_startup_time_limit
        original_tk = gui_module.tk.Tk
        original_showerror = gui_module.messagebox.showerror
        original_app = gui_module.Titan007ExporterApp
        try:
            gui_module.enforce_startup_time_limit = blocked_guard
            gui_module.tk.Tk = lambda: FakeRoot()
            gui_module.messagebox.showerror = lambda title, message: calls.append((title, message))
            gui_module.Titan007ExporterApp = lambda root: calls.append("app")

            gui_module.main()
        finally:
            gui_module.enforce_startup_time_limit = original_guard
            gui_module.tk.Tk = original_tk
            gui_module.messagebox.showerror = original_showerror
            gui_module.Titan007ExporterApp = original_app

        self.assertNotIn("app", calls)
        self.assertIn(("启动失败", "blocked"), calls)


if __name__ == "__main__":
    unittest.main()
