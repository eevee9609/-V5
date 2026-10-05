"""Read-only layout smoke check; run with a working Tcl/Tk runtime."""

import json
import os
import sys
import tkinter as tk
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import customtkinter as ctk
from app.application import HealthCheckGuiApp
from app.data import all_item_names, find_db_item_robust


def fail_callback(_self, exc_type, exc_value, traceback_obj):
    import traceback
    traceback.print_exception(exc_type, exc_value, traceback_obj)
    sys.stderr.flush()
    os._exit(1)


tk.Tk.report_callback_exception = fail_callback
click_time = 1000


def children(widget):
    for child in widget.winfo_children():
        yield child
        yield from children(child)


def in_scroll_content(widget):
    current = widget.master
    while current is not None:
        if isinstance(current, ctk.CTkScrollableFrame):
            return True
        current = getattr(current, "master", None)
    return False


def inspect(window):
    global click_time
    issues = []
    lists = []
    for widget in children(window):
        if not isinstance(widget, (ctk.CTkButton, ctk.CTkLabel, ctk.CTkEntry, tk.Listbox)):
            continue
        if in_scroll_content(widget):
            continue
        if not widget.winfo_ismapped():
            parent = widget
            expected = True
            while parent is not window and parent is not None:
                if not parent.winfo_manager():
                    expected = False
                    break
                parent = parent.master
            if expected:
                issues.append({"type": type(widget).__name__, "text": widget.cget("text") if not isinstance(widget, tk.Listbox) else "list", "reason": "not mapped"})
            continue
        x, y = widget.winfo_rootx(), widget.winfo_rooty()
        right, bottom = x + widget.winfo_width(), y + widget.winfo_height()
        parent = widget.master
        while parent is not None:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            if x < px - 2 or y < py - 2 or right > px + parent.winfo_width() + 2 or bottom > py + parent.winfo_height() + 2:
                issues.append({"type": type(widget).__name__, "text": widget.cget("text") if not isinstance(widget, tk.Listbox) else "list", "reason": "outside parent"})
                break
            if parent is window:
                break
            parent = parent.master
        if isinstance(widget, tk.Listbox) and widget.size():
            row_bounds = widget.bbox(widget.nearest(0))
            if row_bounds:
                target = window.winfo_containing(
                    widget.winfo_rootx() + 8,
                    widget.winfo_rooty() + row_bounds[1] + row_bounds[3] // 2,
                )
                if target is not widget:
                    issues.append({"type": "Listbox", "reason": "covered by another widget", "cover": str(target)})
                else:
                    selected_index = widget.nearest(0)
                    click_time += 1000
                    target.event_generate("<ButtonPress-1>", x=8, y=row_bounds[1] + row_bounds[3] // 2, time=click_time)
                    target.event_generate("<ButtonRelease-1>", x=8, y=row_bounds[1] + row_bounds[3] // 2, time=click_time + 50)
                    if selected_index not in widget.curselection():
                        issues.append({"type": "Listbox", "reason": "visible row cannot be selected"})
                    widget.selection_clear(0, tk.END)
            bottom_index = widget.nearest(widget.winfo_height() - 1)
            bounds = widget.bbox(bottom_index)
            clipped = bool(bounds and bounds[1] + bounds[3] > widget.winfo_height())
            lists.append({"height": widget.winfo_height(), "rows": bottom_index - widget.nearest(0) + 1, "partial_row": clipped})
            if clipped:
                issues.append({"type": "Listbox", "reason": "partial row"})
        label = getattr(widget, "_label", None)
        if isinstance(widget, ctk.CTkLabel) and label and widget.cget("text"):
            if label.winfo_reqwidth() > widget.winfo_width() + 2:
                issues.append({"type": "Label", "text": widget.cget("text"), "reason": "text too wide"})
    return {"issues": issues, "lists": lists}


scaling = float(sys.argv[1]) if len(sys.argv) > 1 else 1.0
ctk.set_widget_scaling(scaling)
ctk.set_window_scaling(scaling)
app = HealthCheckGuiApp()
app.attributes("-alpha", 0)
# Stress the rule panel without saving any changes to the user's configuration.
app.app_settings["special_formula_keywords"] = {
    str(find_db_item_robust(label)["internal_code"]): f"測試{index % 8}"
    for index, label in enumerate(all_item_names[:36])
    if find_db_item_robust(label) and find_db_item_robust(label).get("internal_code")
}
app.open_settings_window("db")
window = app.settings_window
window.attributes("-alpha", 1)
window.minsize(1, 1)
reports = []
try:
    for width, height in ((1520, 820), (1366, 700), (1150, 620)):
        window.geometry(f"{width}x{height}+0+0")
        for page in ("db", "clinic", "routing", "announce", "special"):
            print(f"Checking {width}x{height}: {page}", flush=True)
            app.switch_settings_tab("clinic" if page == "routing" else page)
            if page == "routing":
                for widget in children(window):
                    if isinstance(widget, ctk.CTkButton) and widget.cget("text") == "🔄 院所專屬分派":
                        widget.invoke()
                        break
            app.update()
            app.update_idletasks()
            if page == "db":
                assert app.listbox_db.size() > 0, "Database page did not load any rows"
                initial_count = app.listbox_db.size()
                app.db_search_var.set("__layout_no_matching_item__")
                assert app.listbox_db.size() == 0, "Database search did not filter"
                app.db_search_var.set("")
                app.update()
                assert app.listbox_db.size() == initial_count, "Database search did not restore rows"
            report = inspect(window)
            reports.append({"size": [width, height], "page": page, **report})
            if page == "special":
                for widget in list(children(window)):
                    if isinstance(widget, ctk.CTkButton) and str(widget.cget("text")).startswith("▼  K 欄"):
                        widget.invoke()
                app.update()
                reports.append({"size": [width, height], "page": "special_collapsed", **inspect(window)})
            if page == "routing":
                for widget in children(window):
                    if isinstance(widget, tk.Listbox) and widget.winfo_ismapped():
                        widget.selection_set(0)
                for widget in children(window):
                    if isinstance(widget, ctk.CTkButton) and widget.cget("text") == "➕ 加入杏聯":
                        widget.invoke()
                        break
                app.update()
                reports.append({"size": [width, height], "page": "routing_selected", **inspect(window)})
    failures = [report for report in reports if report["issues"]]
    print(json.dumps({"scaling": scaling, "checked": len(reports), "failures": failures}, ensure_ascii=False, indent=2, default=str))
finally:
    app.destroy()
if any(report["issues"] for report in reports):
    sys.exit(1)
