"""Exercise the real split workflow with synthetic workbooks and captured dialogs."""

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import openpyxl

from app.ui.template_tabs import TemplateTabsMixin


class SplitRoutingNotificationsTest(unittest.TestCase):
    def run_split(self, source_tag, xing_rule="", bo_rule=""):
        with tempfile.TemporaryDirectory(prefix="split-routing-qa-") as directory:
            source = Path(directory) / "synthetic.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet["A1"], sheet["B1"] = "場次", "61"
            sheet["C1"], sheet["D1"] = "院所", "TEST"
            sheet["H4"], sheet["J4"] = "姓名", "身分證"
            sheet["H5"] = "測試受檢者"
            # A recognised anchor keeps dynamic column detection consistent for all tag cases.
            sheet["L1"], sheet["L3"], sheet["L4"] = "杏2", "ANCHOR", "ANCHOR"
            sheet["M1"], sheet["M3"], sheet["M4"], sheet["M5"] = source_tag, "105", "GLU", "@"
            workbook.save(source)
            workbook.close()
            host = SimpleNamespace(
                entry_split_file=SimpleNamespace(get=lambda: str(source)),
                inst_var=SimpleNamespace(get=lambda: "TEST - 測試院所"),
                inst_db={"TEST": {"force_to_xinglian": xing_rule, "force_to_boren": bo_rule}},
            )
            with patch("app.ui.template_tabs.messagebox") as messages:
                TemplateTabsMixin.run_split_macro(host)
                messages.showerror.assert_not_called()
                messages.showinfo.assert_called_once()
                self.assertEqual(messages.showinfo.call_args.args[0], "拆分成功")
                summary = messages.showinfo.call_args.args[1]
                destinations = set()
                for name in ("杏聯", "博仁"):
                    result = openpyxl.load_workbook(source.with_name(f"synthetic_{name}.xlsx"))
                    if "105" in [cell.value for cell in result.active[1]]:
                        destinations.add(name)
                    result.close()
                return summary, messages.showwarning.call_count, destinations

    def test_no_rule_or_absent_item_is_silent(self):
        for rule in ("", "系統:999"):
            with self.subTest(rule=rule):
                summary, warnings, destinations = self.run_split("杏2", bo_rule=rule)
                self.assertNotIn("強制變更分派", summary)
                self.assertEqual(warnings, 0)
                self.assertEqual(destinations, {"杏聯"})

    def test_matching_current_destination_is_not_a_change(self):
        summary, warnings, destinations = self.run_split("博2", bo_rule="系統:105")
        self.assertNotIn("強制變更分派", summary)
        self.assertEqual(warnings, 0)
        self.assertEqual(destinations, {"博仁"})

    def test_real_changes_are_reported_once_after_split(self):
        for tag, xing, bo, target, description in (
            ("杏2", "", "系統:105,系統:999", "博仁", "杏聯 → 博仁"),
            ("博2", "系統:105", "", "杏聯", "博仁 → 杏聯"),
            ("C2", "", "系統:105", "博仁", "杏聯、博仁 → 博仁"),
            ("E", "系統:105", "", "杏聯", "不分派 → 杏聯"),
        ):
            with self.subTest(tag=tag):
                summary, warnings, destinations = self.run_split(tag, xing, bo)
                self.assertIn("強制變更分派 1 項", summary)
                self.assertIn(description, summary)
                self.assertEqual(warnings, 0)
                self.assertEqual(destinations, {target})

    def test_conflicting_rules_still_warn(self):
        _, warnings, destinations = self.run_split("博2", "系統:105", "系統:105")
        self.assertEqual(warnings, 1)
        self.assertEqual(destinations, {"杏聯"})


if __name__ == "__main__":
    unittest.main()
