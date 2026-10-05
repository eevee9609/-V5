"""浮動熊貓與 Excel 加選小助手視窗。"""

import json
import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk
import xlwings as xw

from app.config import (
    APP_ROOT,
    HEALTH_CHECK_ITEMS_FILE,
    ICON_PATH,
    PANDA_IMAGE_PATH,
    UI_FONT,
)
from app.data import get_sorted_items_list


CURRENT_DIR = str(APP_ROOT)
json_file = str(HEALTH_CHECK_ITEMS_FILE)
ICON_PATH = str(ICON_PATH)
PANDA_IMAGE_PATH = str(PANDA_IMAGE_PATH)

class FloatingPanda:
    def __init__(self, main_assistant):
        self.main_root = main_assistant 
        self.top = tk.Toplevel(main_assistant)
        self.top.overrideredirect(True)
        self.top.attributes("-topmost", True)
        
        self.chroma_key = "#012345" 
        self.top.wm_attributes("-transparentcolor", self.chroma_key)
        self.top.configure(bg=self.chroma_key)

        img_path = PANDA_IMAGE_PATH
        try:
            original_img = tk.PhotoImage(file=img_path)
            self.panda_img = original_img.subsample(6, 6) 
            self.label = tk.Label(self.top, image=self.panda_img, bg=self.chroma_key)
        except Exception:
            self.label = tk.Label(self.top, text="🐼\n小幫手", font=(UI_FONT, 12), fg="white", bg="black", bd=2, relief="ridge")
        
        self.label.pack()
        self.label.bind("<Button-1>", self.start_move)
        self.label.bind("<B1-Motion>", self.do_move)
        self.label.bind("<Double-Button-1>", self.toggle_main_app)
        
        self.label.bind("<Button-3>", self.show_menu)
        self.menu = tk.Menu(self.top, tearoff=0, font=(UI_FONT, 10))
        self.menu.add_command(label="📝 打開小熊貓視窗", command=self.open_main_app)
        self.menu.add_separator()
        self.menu.add_command(label="💤 讓熊貓下班 (關閉外掛)", command=self.close_all)

        self.x, self.y = 0, 0
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        pos_x = screen_width - 150
        pos_y = screen_height - 150
        self.top.geometry(f"+{pos_x}+{pos_y}")

    def start_move(self, event):
        self.x, self.y = event.x, event.y

    def do_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        self.top.geometry(f"+{self.top.winfo_x() + deltax}+{self.top.winfo_y() + deltay}")

    def toggle_main_app(self, event=None):
        if self.main_root.state() == "withdrawn": self.main_root.deiconify()
        else: self.main_root.withdraw()
            
    def open_main_app(self):
        self.main_root.deiconify()

    def show_menu(self, event):
        self.menu.post(event.x_root, event.y_root)

    def close_all(self):
        self.main_root.destroy()

class SmartAssistantTool(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("⚡ 智能小熊貓加選助手")
        self.geometry("920x640")
        self.minsize(860, 590)
        self.is_pinned = True
        self.attributes('-topmost', self.is_pinned)

        # ⭐ 載入應用程式圖示
        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(ICON_PATH)
            except Exception:
                pass
        
        self.db_items = {}
        self.items_list = []
        self.person_row_mapping = {}
        self.history_stack = []
        self.selection_order = []

        # Excel 目標必須由使用者在介面中明確選擇，不再猜測作用中的活頁簿／工作表。
        self.workbook_options = {}
        self.sheet_options = {}
        self.selected_workbook_identity = None
        self.selected_sheet_name = None
        self.last_write_result = None
        self.last_write_records = []
        self.last_write_history_render_job = None
        self.template_validation = {"ok": False, "errors": ["尚未選擇 Excel 目標"], "warnings": []}

        self.item_vars = {}
        self.item_checkboxes = {}
        self.visible_checkboxes = [] 
        self.current_focus_idx = -1  

        self.current_alpha = parent.current_alpha
        self.attributes('-alpha', self.current_alpha)
        self.opacity_var = tk.DoubleVar(value=self.current_alpha)

        self.load_database()
        self.build_ui()
        self.after(100, self.auto_position)
        self.panda = FloatingPanda(self)
        self.lift()

    def auto_position(self):
        try:
            self.update_idletasks() 
            parent_x = self.master.winfo_x()
            parent_y = self.master.winfo_y()
            parent_width = self.master.winfo_width()
            
            my_width = self.winfo_width()
            my_height = self.winfo_height()
            
            if my_width <= 10: my_width = 920
            if my_height <= 10: my_height = 640
            
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            target_x = parent_x + parent_width + 10
            target_y = parent_y
            
            if target_x + my_width > screen_width:
                target_x = parent_x - my_width - 10
                if target_x < 0:
                    target_x = screen_width - my_width - 10

            if target_y + my_height > screen_height - 50: target_y = screen_height - my_height - 50
            if target_y < 0: target_y = 10
                
            self.geometry(f"+{int(target_x)}+{int(target_y)}")
        except Exception: pass

    def load_database(self):
        global json_file
        if not os.path.exists(json_file): return
            
        try:
            with open(json_file, 'r', encoding='utf-8') as f: raw_data = json.load(f)
                
            self.db_items.clear()
            for code, info in raw_data.items():
                name_zh = info.get("中文名稱", "").strip()
                name_en = info.get("英文簡寫", "").strip()
                org_raw = info.get("歸屬單位", "C2").strip().upper()
                org_display = org_raw.replace("C", "杏").replace("B", "博")
                internal_code = info.get("系統代碼", "").strip()
                category = info.get("檢驗類別", "").strip()
                
                if name_zh:
                    if internal_code: display_name = f"[{org_display}] {internal_code} - {name_zh} ({name_en})"
                    else: display_name = f"[{org_display}] {name_zh} ({name_en})"
                        
                    if category: display_name += f" | {category}"
                    self.db_items[display_name] = {'category': org_display, 'name_zh': name_zh, 'code': code, 'name_en': name_en, 'test_category': category}
            self.items_list = list(self.db_items.keys())
        except Exception as e: print("小熊貓讀取 JSON 失敗:", e)

    def format_opacity_label(self, alpha):
        return f"{int(round(alpha * 100))}%"

    def set_opacity(self, value=None):
        alpha = float(self.opacity_var.get() if value is None else value)
        alpha = max(0.35, min(1.0, alpha))
        self.current_alpha = alpha
        self.attributes('-alpha', alpha)
        if hasattr(self, 'opacity_label'): self.opacity_label.configure(text=self.format_opacity_label(alpha))
        if hasattr(self, 'panda') and getattr(self, 'panda', None) and self.panda.top.winfo_exists():
            try: self.panda.top.attributes('-alpha', alpha)
            except Exception: pass

    @staticmethod
    def _normalise_excel_text(value):
        """將 Excel 的空值與數字代碼轉成可穩定比較的文字。"""
        if value is None:
            return ""
        text = str(value).strip()
        if text.lower() == "none":
            return ""
        if text.endswith(".0"):
            text = text[:-2]
        return text

    @staticmethod
    def _as_row_values(value):
        """xlwings 在單一儲存格與單列範圍會傳回不同結構，統一成一維 list。"""
        if value is None:
            return []
        if isinstance(value, (list, tuple)):
            if len(value) == 1 and isinstance(value[0], (list, tuple)):
                return list(value[0])
            if value and all(isinstance(row, (list, tuple)) and len(row) == 1 for row in value):
                return [row[0] for row in value]
            return list(value)
        return [value]

    def _workbook_identity(self, app, wb):
        try:
            file_name = str(wb.fullname)
        except Exception:
            file_name = str(wb.name)
        try:
            app_id = str(app.pid)
        except Exception:
            app_id = "unknown"
        return f"{app_id}|{os.path.normcase(file_name)}"

    def _set_target_status(self, text, color="#AAAAAA"):
        if hasattr(self, "target_status_label"):
            self.target_status_label.configure(text=text, text_color=color)

    def _set_result_status(self, text, color="#AAAAAA"):
        if hasattr(self, "result_status_label"):
            self.result_status_label.configure(text=text, text_color=color)

    def _last_write_item_label(self, item_name):
        """最近寫入紀錄採英文簡碼優先，讓左側小面板保持好讀。"""
        data = self.db_items.get(item_name, {})
        return (
            str(data.get("name_en", "") or "").strip()
            or str(data.get("name_zh", "") or "").strip()
            or str(item_name or "").strip()
        )

    def _format_last_write_record(self, record):
        """以緊湊格式顯示一筆近期寫入，讓有限的高度容納更多紀錄。"""
        person_text = str(record.get("person", "") or "最近寫入完成")
        item_labels = record.get("items", [])
        items_text = "、".join(item_labels) if item_labels else "（本次沒有可顯示的項目）"
        return f"▸ {person_text}\n  {items_text}"

    def _set_last_write_history_text(self, text):
        if not hasattr(self, "last_write_history_box"):
            return
        self.last_write_history_box.configure(state="normal")
        self.last_write_history_box.delete("1.0", tk.END)
        self.last_write_history_box.insert("1.0", text)
        self.last_write_history_box.configure(state="disabled")

    def _last_write_display_line_count(self):
        """取得文字實際換行後的行數；無法取得時使用保守估算。"""
        try:
            text_widget = getattr(self.last_write_history_box, "_textbox", self.last_write_history_box)
            result = text_widget.count("1.0", "end-1c", "displaylines")
            return int(result[0] if isinstance(result, (tuple, list)) else result)
        except Exception:
            return len(self.last_write_history_box.get("1.0", "end-1c").splitlines())

    def _render_last_write_history(self):
        """填滿可用區域後才淘汰最早紀錄，永遠保留最新一筆。"""
        self.last_write_history_render_job = None
        if not hasattr(self, "last_write_history_box"):
            return

        if not self.last_write_records:
            self._set_last_write_history_text("尚未寫入任何項目")
            return

        while True:
            text = "\n\n".join(
                self._format_last_write_record(record)
                for record in self.last_write_records
            )
            self._set_last_write_history_text(text)
            try:
                self.update_idletasks()
                available_height = self.last_write_history_box.winfo_height()
            except Exception:
                return

            # 視窗尚未完成配置時，保留所有紀錄並在稍後依實際高度重算。
            if available_height < 40:
                self._schedule_last_write_history_render(delay=120)
                return

            font_line_height = 19
            available_lines = max(1, (available_height - 12) // font_line_height)
            if (
                self._last_write_display_line_count() <= available_lines
                or len(self.last_write_records) == 1
            ):
                return
            # 清單由新到舊排列；超出可視高度時移除最下方的最早紀錄。
            self.last_write_records.pop()

    def _schedule_last_write_history_render(self, event=None, delay=80):
        if self.last_write_history_render_job is not None:
            try:
                self.after_cancel(self.last_write_history_render_job)
            except Exception:
                pass
        self.last_write_history_render_job = self.after(delay, self._render_last_write_history)

    def clear_last_write_history(self):
        """供主程式初始化下一場時清除本輪小熊貓操作紀錄。"""
        self.last_write_records.clear()
        self._schedule_last_write_history_render(delay=0)

    def _update_last_write_display(self, person, item_names):
        """新增一筆紀錄；若空間不足，自動由最早的一筆開始淘汰。"""
        if not hasattr(self, "last_write_history_box"):
            return

        person = person if isinstance(person, dict) else {}
        seq = str(person.get("seq", "") or "").strip()
        name = str(person.get("name", "") or "").strip()
        item_labels = [
            self._last_write_item_label(item_name)
            for item_name in item_names
            if str(item_name or "").strip()
        ]
        person_text = f"{seq}｜{name}" if (seq or name) else "最近寫入完成"
        # 最新寫入固定排在最上方，方便連續作業時立即核對。
        self.last_write_records.insert(0, {"person": person_text, "items": item_labels})
        self._schedule_last_write_history_render(delay=0)

    def _find_default_workbook_display(self):
        """選出安全的預設活頁簿：優先含「團檢大表」，只有單一活頁簿才放寬。"""
        if not self.workbook_options:
            return None

        active_identity = None
        try:
            for app in xw.apps:
                try:
                    active_wb = app.books.active
                except Exception:
                    active_wb = None
                if active_wb is not None:
                    active_identity = self._workbook_identity(app, active_wb)
                    break
        except Exception:
            active_identity = None

        template_candidates = []
        for display, context in self.workbook_options.items():
            try:
                sheet_names = {str(sheet.name) for sheet in context["wb"].sheets}
            except Exception:
                sheet_names = set()
            if "團檢大表" in sheet_names:
                template_candidates.append((display, context))

        if template_candidates:
            active_match = next(
                (display for display, context in template_candidates if context["identity"] == active_identity),
                None,
            )
            return active_match or template_candidates[0][0]
        if len(self.workbook_options) == 1:
            return next(iter(self.workbook_options))
        return None

    def refresh_excel_context(self, preserve_selection=True):
        """重新列出所有已開啟的 Excel；不依檔名或作用中視窗猜測目標。"""
        previous_identity = self.selected_workbook_identity if preserve_selection else None
        previous_sheet = self.selected_sheet_name if preserve_selection else None
        self.workbook_options.clear()

        try:
            for app in xw.apps:
                for wb in app.books:
                    identity = self._workbook_identity(app, wb)
                    try:
                        app_id = str(app.pid)
                    except Exception:
                        app_id = "?"
                    # 下拉選單維持短檔名；只有同名活頁簿才補上 Excel PID 區分。
                    base_display = str(wb.name)
                    display = base_display
                    suffix = 2
                    while display in self.workbook_options:
                        display = f"{base_display}（Excel {app_id}）"
                        if display in self.workbook_options:
                            display = f"{base_display}（Excel {app_id}-{suffix}）"
                        suffix += 1
                    self.workbook_options[display] = {
                        "app": app,
                        "wb": wb,
                        "identity": identity,
                    }
        except Exception as exc:
            self.workbook_options.clear()
            self._set_target_status(f"✕ 無法讀取 Excel：{exc}", "#E57373")

        if hasattr(self, "workbook_cb"):
            choices = list(self.workbook_options)
            self.workbook_cb["values"] = choices

            restored_display = next(
                (display for display, context in self.workbook_options.items()
                 if context["identity"] == previous_identity),
                None,
            )
            if restored_display:
                self.workbook_var.set(restored_display)
                self.selected_workbook_identity = previous_identity
                self._populate_sheet_options(restored_display, previous_sheet)
            else:
                default_display = self._find_default_workbook_display() if previous_identity is None else None
                if default_display:
                    context = self.workbook_options[default_display]
                    self.workbook_var.set(default_display)
                    self.selected_workbook_identity = context["identity"]
                    self._populate_sheet_options(default_display, "團檢大表")
                else:
                    self.workbook_var.set("")
                    self.sheet_var.set("")
                    self.sheet_options.clear()
                    self.selected_workbook_identity = None
                    self.selected_sheet_name = None
                    if hasattr(self, "sheet_cb"):
                        self.sheet_cb["values"] = []
                    self._set_target_status("請先選擇 Excel 活頁簿與工作表", "#F0AD4E")
                self.refresh_person_selector(preserve_selection=False)

    def _populate_sheet_options(self, workbook_display, preferred_sheet=None):
        self.sheet_options.clear()
        context = self.workbook_options.get(workbook_display)
        if not context:
            return

        try:
            for sheet in context["wb"].sheets:
                self.sheet_options[str(sheet.name)] = str(sheet.name)
        except Exception as exc:
            self._set_target_status(f"✕ 無法讀取工作表：{exc}", "#E57373")

        if hasattr(self, "sheet_cb"):
            self.sheet_cb["values"] = list(self.sheet_options)

        if preferred_sheet and preferred_sheet in self.sheet_options:
            self.sheet_var.set(preferred_sheet)
            self.selected_sheet_name = preferred_sheet
            self._update_target_validation()
        else:
            self.sheet_var.set("")
            self.selected_sheet_name = None
            self._set_target_status("請選擇要寫入的工作表", "#F0AD4E")

    def _on_workbook_selected(self, event=None):
        del event
        workbook_display = self.workbook_var.get()
        context = self.workbook_options.get(workbook_display)
        if not context:
            return
        self.selected_workbook_identity = context["identity"]
        self.selected_sheet_name = None
        try:
            sheet_names = {str(sheet.name) for sheet in context["wb"].sheets}
        except Exception:
            sheet_names = set()
        preferred_sheet = "團檢大表" if "團檢大表" in sheet_names else None
        self._populate_sheet_options(workbook_display, preferred_sheet)
        self.refresh_person_selector(preserve_selection=False)

    def _on_sheet_selected(self, event=None):
        del event
        self.selected_sheet_name = self.sheet_var.get().strip() or None
        self._update_target_validation()
        self.refresh_person_selector(preserve_selection=False)

    def validate_target_sheet(self, sheet):
        """驗證小熊貓可安全使用的團檢模板基本結構。"""
        result = {
            "ok": False,
            "errors": [],
            "warnings": [],
            "source_col": None,
            "header_names": [],
            "header_codes": [],
            "max_row": 0,
        }
        try:
            max_row = int(sheet.used_range.last_cell.row)
            result["max_row"] = max_row
            if max_row < 4:
                result["errors"].append("工作表缺少第 1～4 列模板標頭")
                return result

            # 既有模板通常少於 ET 欄，但不要因項目變多而漏掉最右側項目欄。
            max_col = max(150, int(sheet.used_range.last_cell.column))
            header_names = self._as_row_values(sheet.range((2, 1), (2, max_col)).value)
            header_codes = self._as_row_values(sheet.range((3, 1), (3, max_col)).value)
            result["header_names"] = header_names
            result["header_codes"] = header_codes

            # 現有模板從 K 欄開始才是檢查項目；必須有可複製格式的項目欄，才允許新增。
            candidate_columns = [
                index + 1
                for index, value in enumerate(header_names)
                if index + 1 >= 11 and self._normalise_excel_text(value)
            ]
            if not candidate_columns:
                result["errors"].append("第 2 列 K 欄以後找不到可用的檢查項目標頭")
                return result

            result["source_col"] = max(candidate_columns)
            if not any(self._normalise_excel_text(value) for value in header_codes):
                result["warnings"].append("第 3 列未偵測到系統代碼，將以中文名稱比對既有欄位")
            if str(sheet.name) != "團檢大表":
                result["warnings"].append("目前不是「團檢大表」，請確認這是正確模板工作表")
            result["ok"] = True
        except Exception as exc:
            result["errors"].append(f"無法讀取模板標頭：{exc}")
        return result

    def _update_target_validation(self):
        target = self.get_selected_target(require_valid=False, notify=False)
        if not target:
            return False
        return bool(self.template_validation.get("ok"))

    def get_selected_target(self, require_valid=True, notify=False):
        """取得使用者明確選定的 app、workbook、sheet；絕不退回作用中工作表。"""
        workbook_display = self.workbook_var.get().strip() if hasattr(self, "workbook_var") else ""
        sheet_name = self.sheet_var.get().strip() if hasattr(self, "sheet_var") else ""
        context = self.workbook_options.get(workbook_display)

        if not context:
            message = "請先在小熊貓助手選擇 Excel 活頁簿"
            self.template_validation = {"ok": False, "errors": [message], "warnings": []}
            self._set_target_status(f"✕ {message}", "#E57373")
            if notify:
                messagebox.showwarning("小熊貓提示", message, parent=self)
            return None
        if not sheet_name:
            message = "請先選擇要寫入的工作表"
            self.template_validation = {"ok": False, "errors": [message], "warnings": []}
            self._set_target_status(f"✕ {message}", "#E57373")
            if notify:
                messagebox.showwarning("小熊貓提示", message, parent=self)
            return None

        try:
            app = context["app"]
            wb = context["wb"]
            # 讀取 name 可確認 COM 活頁簿仍存在；關閉後會拋出例外。
            _ = wb.name
            sheet = wb.sheets[sheet_name]
        except Exception:
            message = "選定的 Excel 或工作表已關閉，請重新掃描後再選擇"
            self.template_validation = {"ok": False, "errors": [message], "warnings": []}
            self._set_target_status(f"✕ {message}", "#E57373")
            if notify:
                messagebox.showwarning("小熊貓提示", message, parent=self)
            return None

        self.selected_workbook_identity = context["identity"]
        self.selected_sheet_name = str(sheet.name)
        validation = self.validate_target_sheet(sheet)
        self.template_validation = validation
        if validation["ok"]:
            status = "✓ 模板已通過基本驗證"
            if validation["warnings"]:
                status += f"（{validation['warnings'][0]}）"
            self._set_target_status(status, "#57C785")
        else:
            message = validation["errors"][0] if validation["errors"] else "模板驗證失敗"
            self._set_target_status(f"✕ {message}", "#E57373")
            if require_valid:
                if notify:
                    messagebox.showerror("小熊貓提示", f"不能寫入此工作表：\n{message}", parent=self)
                return None
        return app, wb, sheet

    def get_target_workbook(self):
        """保留舊介面相容性，但只回傳已選定的活頁簿。"""
        target = self.get_selected_target(require_valid=True, notify=False)
        if not target:
            return None, None
        app, wb, _sheet = target
        return app, wb

    def refresh_person_list(self):
        """從已驗證的目標工作表讀取受檢者，不再改用作用中工作表。"""
        self.person_row_mapping.clear()
        target = self.get_selected_target(require_valid=True, notify=False)
        if not target:
            return []

        try:
            _app, _wb, sheet = target
            max_row = max(5, int(sheet.used_range.last_cell.row))
            data_range = sheet.range((5, 4), (max_row, 8)).value
            if not data_range:
                return []
            if not isinstance(data_range, list):
                data_range = [[data_range]]
            elif data_range and not isinstance(data_range[0], (list, tuple)):
                data_range = [data_range]

            person_list = []
            for idx, row in enumerate(data_range):
                if not isinstance(row, (list, tuple)):
                    continue
                seq = row[0] if len(row) > 0 else None
                name = row[4] if len(row) > 4 else None
                seq_str = self._normalise_excel_text(seq)
                name_str = self._normalise_excel_text(name)
                if not seq_str or not name_str:
                    continue

                excel_row = 5 + idx
                display = f"{seq_str}｜{name_str}"
                if display in self.person_row_mapping:
                    display = f"{display}｜第 {excel_row} 列"
                person_list.append(display)
                self.person_row_mapping[display] = {
                    "row": excel_row,
                    "seq": seq_str,
                    "name": name_str,
                }
            return person_list
        except Exception as exc:
            self._set_target_status(f"✕ 無法讀取受檢者：{exc}", "#E57373")
            return []

    def refresh_person_selector(self, preserve_selection=True):
        if not hasattr(self, "person_cb"):
            return []
        current = self.person_cb.get().strip() if preserve_selection else ""
        persons = self.refresh_person_list()
        self.person_cb["values"] = persons
        if current and current in persons:
            self.person_cb.set(current)
        else:
            self.person_cb.set("— 請選擇受檢者 —")
        return persons

    def get_selected_items(self):
        """摘要、預覽與實際寫入共用同一個選取順序。"""
        selected = [
            item for item in self.selection_order
            if item in self.item_vars and self.item_vars[item].get()
        ]
        for item, var in self.item_vars.items():
            if var.get() and item not in selected:
                selected.append(item)
        return selected

    def _short_item_label(self, item):
        data = self.db_items.get(item, {})
        english = data.get("name_en") or "未設定英文簡碼"
        chinese = data.get("name_zh") or item
        code = data.get("code") or "無系統代碼"
        return f"{english} ｜ {chinese} ｜ {code}"

    def clear_selected_items(self):
        selected = self.get_selected_items()
        if not selected:
            return
        if not messagebox.askyesno(
            "清除勾選",
            f"確定要清除目前勾選的 {len(selected)} 個項目嗎？",
            parent=self,
        ):
            return
        for item in selected:
            self.item_vars[item].set(False)

    def _trace_search(self, *args):
        search_terms = self.search_var.get().strip().lower().split()
        sort_by = self.panda_sort_var.get()
        category_filter = self.category_filter_var.get() if hasattr(self, "category_filter_var") else "全部類別"
        self.visible_checkboxes.clear()
        self.current_focus_idx = -1

        for cb in self.item_checkboxes.values():
            cb.pack_forget()

        filtered_items = []
        for item in self.items_list:
            data = self.db_items.get(item, {})
            item_category = data.get("test_category") or "未分類"
            if category_filter != "全部類別" and item_category != category_filter:
                continue
            searchable = " ".join([
                item,
                str(data.get("code", "")),
                str(data.get("name_en", "")),
                str(data.get("name_zh", "")),
                str(data.get("category", "")),
                str(item_category),
            ]).lower()
            if all(term in searchable for term in search_terms):
                filtered_items.append(item)
        sorted_items = get_sorted_items_list(filtered_items, sort_by)

        for item in sorted_items:
            cb = self.item_checkboxes[item]
            cb.pack(fill="x", anchor="w", pady=3, padx=5)
            cb.configure(text_color="white") 
            self.visible_checkboxes.append(cb)
            
        try: self.scrollable_frame._parent_canvas.yview_moveto(0)
        except Exception: pass

    def _trace_summary(self, *args):
        try:
            selected = self.get_selected_items()
            if hasattr(self, "selection_count_label"):
                self.selection_count_label.configure(text=f"已勾選 {len(selected)} 項")
            self.summary_text.configure(state="normal")
            self.summary_text.delete("1.0", tk.END)
            if not selected:
                self.summary_text.insert(tk.END, "\n(尚未勾選任何項目)")
            else:
                for item in selected:
                    self.summary_text.insert(tk.END, f"☑ {self._short_item_label(item)}\n")
            self.summary_text.configure(state="disabled")
        except Exception:
            pass

    def build_write_plan(self, app, wb, sheet, person, selected_items):
        """在寫入前一次算出所有異動，讓確認內容與實際寫入保持一致。"""
        validation = self.validate_target_sheet(sheet)
        if not validation["ok"]:
            raise ValueError(validation["errors"][0])

        source_col = validation["source_col"]
        max_row = max(300, validation["max_row"])
        header_names = list(validation["header_names"])
        header_codes = list(validation["header_codes"])
        current_col = source_col
        plan_items = []

        def header_at(values, index):
            return self._normalise_excel_text(values[index - 1]) if index <= len(values) else ""

        for item_name in selected_items:
            item_data = self.db_items.get(item_name)
            if not item_data:
                raise ValueError(f"找不到項目資料：{item_name}")
            target_code = self._normalise_excel_text(item_data.get("code"))
            target_name = self._normalise_excel_text(item_data.get("name_zh"))
            found_col = 0

            for column in range(1, current_col + 1):
                header_code = header_at(header_codes, column)
                header_name = header_at(header_names, column)
                if target_code and header_code == target_code:
                    found_col = column
                    break
                if not target_code and target_name and header_name == target_name:
                    found_col = column
                    break

            is_new = found_col == 0
            if is_new:
                current_col += 1
                target_col = current_col
                while len(header_names) < target_col:
                    header_names.append(None)
                while len(header_codes) < target_col:
                    header_codes.append(None)
                header_names[target_col - 1] = item_data.get("name_zh", "")
                header_codes[target_col - 1] = item_data.get("code", "")
                previous_value = None
            else:
                target_col = found_col
                previous_value = sheet.range((person["row"], target_col)).value

            previous_text = self._normalise_excel_text(previous_value)
            plan_items.append({
                "item_name": item_name,
                "item_data": item_data,
                "column": target_col,
                "is_new": is_new,
                "previous_value": previous_value,
                "will_overwrite": bool(previous_text and previous_text != "@"),
            })

        return {
            "app": app,
            "wb": wb,
            "sheet": sheet,
            "workbook_identity": self._workbook_identity(app, wb),
            "workbook_name": str(wb.name),
            "sheet_name": str(sheet.name),
            "person": person,
            "source_col": source_col,
            "max_row": max_row,
            "items": plan_items,
        }

    def show_write_preview(self, plan):
        """保留既有呼叫相容性；寫入摘要改由左側紀錄與底部狀態即時呈現。"""
        del plan
        return True

    def _safe_new_column_for_undo(self, sheet, new_col, action):
        """僅在新增欄標頭與內容都未被其他人改動時才清除它。"""
        try:
            col = new_col["col"]
            expected = [self._normalise_excel_text(value) for value in new_col["headers"]]
            current = self._as_row_values(sheet.range((1, col), (4, col)).value)
            current = [self._normalise_excel_text(value) for value in current]
            if current != expected:
                return False

            body_values = sheet.range((5, col), (action["max_row"], col)).value
            if not isinstance(body_values, list):
                body_values = [body_values]
            for index, value in enumerate(body_values, start=5):
                text = self._normalise_excel_text(value)
                if index == action["row"]:
                    if text not in ("", "@"):
                        return False
                elif text:
                    return False
            return True
        except Exception:
            return False

    def execute_write_plan(self, plan):
        """套用已確認的寫入計畫，失敗時盡量還原已處理的格子。"""
        app, sheet = plan["app"], plan["sheet"]
        person_row = plan["person"]["row"]
        cell_changes = []
        new_columns = []
        try:
            for planned in plan["items"]:
                column = planned["column"]
                item_data = planned["item_data"]
                if planned["is_new"]:
                    sheet.range((1, plan["source_col"]), (plan["max_row"], plan["source_col"])).api.Copy()
                    sheet.range((1, column), (plan["max_row"], column)).api.PasteSpecial(Paste=-4122)
                    app.api.CutCopyMode = False
                    sheet.range((1, column), (4, column)).number_format = "@"
                    headers = [
                        item_data.get("category", ""),
                        item_data.get("name_zh", ""),
                        item_data.get("code", ""),
                        item_data.get("name_en", ""),
                    ]
                    for header_row, value in enumerate(headers, start=1):
                        sheet.range((header_row, column)).value = value
                    sheet.range((1, column)).column_width = sheet.range((1, plan["source_col"])).column_width
                    new_columns.append({"col": column, "headers": headers})

                target_cell = sheet.range((person_row, column))
                cell_changes.append({
                    "col": column,
                    "prev_val": target_cell.value,
                    "prev_color": target_cell.color,
                })
                target_cell.value = "@"
                target_cell.color = (165, 243, 241)

            action = {
                "workbook_identity": plan["workbook_identity"],
                "workbook_name": plan["workbook_name"],
                "sheet_name": plan["sheet_name"],
                "row": person_row,
                "person": dict(plan["person"]),
                "item_names": [planned["item_name"] for planned in plan["items"]],
                "new_cols": new_columns,
                "cell_changes": cell_changes,
                "max_row": plan["max_row"],
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            self.history_stack.append(action)
            self.last_write_result = action
            self.undo_btn.configure(state="normal")

            if not self.lock_var.get():
                for var in self.item_vars.values():
                    var.set(False)

            if cell_changes:
                first_col = cell_changes[0]["col"]
                try:
                    sheet.range((person_row, first_col)).api.Select()
                except Exception:
                    pass
                try:
                    app.api.ActiveWindow.ScrollRow = max(1, person_row - 3)
                    app.api.ActiveWindow.ScrollColumn = max(1, first_col - 3)
                except Exception:
                    pass
            return action
        except Exception:
            for change in reversed(cell_changes):
                try:
                    target_cell = sheet.range((person_row, change["col"]))
                    target_cell.value = change["prev_val"]
                    target_cell.color = change.get("prev_color")
                except Exception:
                    pass
            for new_col in reversed(new_columns):
                try:
                    sheet.range((1, new_col["col"]), (plan["max_row"], new_col["col"])).clear()
                except Exception:
                    pass
            raise
        finally:
            try:
                app.api.CutCopyMode = False
            except Exception:
                pass

    def write_to_excel(self):
        target = self.get_selected_target(require_valid=True, notify=True)
        if not target:
            return
        app, wb, sheet = target

        # 寫入前重新讀取名單，避免使用者在 Excel 排序／插列後寫到舊列號。
        self.refresh_person_selector(preserve_selection=True)
        selected_person = self.person_row_mapping.get(self.person_cb.get().strip())
        if not selected_person:
            messagebox.showwarning("小熊貓提示", "請先正確選擇受檢者。", parent=self)
            return

        selected_items = self.get_selected_items()
        if not selected_items:
            messagebox.showwarning("小熊貓提示", "請至少勾選一個檢查項目。", parent=self)
            return

        try:
            plan = self.build_write_plan(app, wb, sheet, selected_person, selected_items)
        except Exception as exc:
            messagebox.showerror("小熊貓提示", f"無法建立寫入摘要：{exc}", parent=self)
            return
        try:
            action = self.execute_write_plan(plan)
            new_count = len(action["new_cols"])
            self._set_result_status(
                f"✓ 已寫入 {len(action['cell_changes'])} 項；可使用「復原上一動」",
                "#57C785",
            )
            self._update_last_write_display(action["person"], action.get("item_names", []))
        except Exception as exc:
            self._set_result_status("✕ 寫入失敗，已嘗試還原本次變更", "#E57373")
            messagebox.showerror("小熊貓提示", f"寫入失敗：{exc}", parent=self)

    def undo_last_action(self):
        if not self.history_stack:
            return
        action = self.history_stack[-1]
        target = self.get_selected_target(require_valid=False, notify=True)
        if not target:
            return
        _app, wb, sheet = target
        if self._workbook_identity(_app, wb) != action["workbook_identity"] or str(sheet.name) != action["sheet_name"]:
            messagebox.showwarning(
                "小熊貓提示",
                f"請先選回剛才寫入的 Excel／工作表：\n{action['workbook_name']}／{action['sheet_name']}",
                parent=self,
            )
            return

        safe_new_cols = [
            new_col for new_col in action["new_cols"]
            if self._safe_new_column_for_undo(sheet, new_col, action)
        ]
        protected_new_count = len(action["new_cols"]) - len(safe_new_cols)
        try:
            for change in reversed(action["cell_changes"]):
                target_cell = sheet.range((action["row"], change["col"]))
                target_cell.value = change["prev_val"]
                target_cell.color = change.get("prev_color")
            for new_col in reversed(safe_new_cols):
                sheet.range((1, new_col["col"]), (action["max_row"], new_col["col"])).clear()

            self.history_stack.pop()
            if not self.history_stack:
                self.undo_btn.configure(state="disabled")
            if protected_new_count:
                message = f"已還原本次受檢者加選；{protected_new_count} 個新增欄位已有其他內容，為保護資料而保留。"
                self._set_result_status("⚠ 已還原格值；部分新增欄位因已有內容而保留", "#F0AD4E")
            else:
                message = "已安全復原上一筆小熊貓寫入。"
                self._set_result_status("↩ 已復原上一筆寫入", "#57C785")
            messagebox.showinfo("小熊貓復原完成", message, parent=self)
        except Exception as exc:
            messagebox.showerror("小熊貓提示", f"復原失敗：{exc}", parent=self)

    def toggle_pin(self):
        self.is_pinned = not self.is_pinned
        self.attributes('-topmost', self.is_pinned)
        if self.is_pinned: self.pin_btn.configure(text="📌 取消置頂", fg_color="#D9534F", hover_color="#C9302C")
        else: self.pin_btn.configure(text="📌 視窗置頂", fg_color="#1F77B4", hover_color="#155A8A")

    def highlight_cb(self):
        for i, cb in enumerate(self.visible_checkboxes):
            if i == self.current_focus_idx:
                cb.configure(text_color="#5BC0DE")
                try:
                    y_pos = cb.winfo_y()
                    self.scrollable_frame._parent_canvas.yview_moveto(max(0, (y_pos - 80) / self.scrollable_frame.winfo_height()))
                except Exception: pass
            else: cb.configure(text_color="white")

    def build_ui(self):
        opacity_bar = ctk.CTkFrame(self, fg_color="transparent")
        opacity_bar.pack(fill="x", padx=15, pady=(15, 6))
        ctk.CTkLabel(opacity_bar, text="透明度", font=(UI_FONT, 12, "bold")).pack(side="left")
        self.opacity_slider = ctk.CTkSlider(opacity_bar, from_=0.35, to=1.0, number_of_steps=13, variable=self.opacity_var, command=self.set_opacity, width=180)
        self.opacity_slider.pack(side="left", padx=(8, 8))
        self.opacity_label = ctk.CTkLabel(opacity_bar, text=self.format_opacity_label(self.current_alpha), font=(UI_FONT, 12), text_color="#AAAAAA")
        self.opacity_label.pack(side="left")
        self.set_opacity(self.current_alpha)

        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=5, padx=15)
        self.pin_btn = ctk.CTkButton(top_frame, text="📌 取消置頂", command=self.toggle_pin, fg_color="#D9534F", hover_color="#C9302C", font=(UI_FONT, 11, "bold"), height=28, width=100)
        self.pin_btn.pack(side=tk.LEFT)
        
        self.undo_btn = ctk.CTkButton(
            top_frame, text="↩️ 復原上一動", command=self.undo_last_action, 
            fg_color="#E67E22", hover_color="#D35400", 
            text_color="#FFFFFF", text_color_disabled="#7F7F7F", 
            font=(UI_FONT, 11, "bold"), state="disabled", 
            height=28, width=110
        )
        self.undo_btn.pack(side=tk.LEFT, padx=5)
        
        target_frame = ctk.CTkFrame(self, corner_radius=10, fg_color="#2B2B2B")
        target_frame.pack(side=tk.TOP, fill=tk.X, padx=15, pady=(2, 6))
        target_row = ctk.CTkFrame(target_frame, fg_color="transparent")
        target_row.pack(fill=tk.X, padx=10, pady=8)
        ctk.CTkLabel(target_row, text="Excel", font=(UI_FONT, 12, "bold"), text_color="#5BC0DE").pack(side=tk.LEFT)
        self.workbook_var = tk.StringVar()
        self.workbook_cb = ttk.Combobox(target_row, textvariable=self.workbook_var, state="readonly", width=24)
        self.workbook_cb.pack(side=tk.LEFT, padx=(6, 8))
        self.workbook_cb.bind("<<ComboboxSelected>>", self._on_workbook_selected)
        ctk.CTkLabel(target_row, text="工作表", font=(UI_FONT, 12, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.sheet_var = tk.StringVar()
        self.sheet_cb = ttk.Combobox(target_row, textvariable=self.sheet_var, state="readonly", width=17)
        self.sheet_cb.pack(side=tk.LEFT, padx=(0, 8))
        self.sheet_cb.bind("<<ComboboxSelected>>", self._on_sheet_selected)
        self.scan_excel_btn = ctk.CTkButton(
            target_row,
            text="🔄 掃描 Excel",
            command=self.refresh_excel_context,
            font=(UI_FONT, 11, "bold"),
            fg_color="#1F77B4",
            hover_color="#155A8A",
            height=28,
            width=110,
        )
        self.scan_excel_btn.pack(side=tk.RIGHT)
        self.target_status_label = ctk.CTkLabel(
            target_row,
            text="請先選擇 Excel 活頁簿與工作表",
            font=(UI_FONT, 11, "bold"),
            text_color="#F0AD4E",
            anchor="w",
        )
        self.target_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=5)

        left_outer = ctk.CTkFrame(main_frame, corner_radius=10, fg_color="#2B2B2B", width=190)
        left_outer.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))
        left_outer.pack_propagate(False)
        ctk.CTkLabel(left_outer, text="👤 選擇受檢者", font=(UI_FONT, 13, "bold"), text_color="#5BC0DE").pack(pady=(12, 5))
        
        style = ttk.Style()
        if "clam" in style.theme_names(): style.theme_use("clam")
        style.configure("TCombobox", fieldbackground="#2B2B2B", background="#333333", foreground="white", borderwidth=0, arrowcolor="white")
        
        self.option_add("*TCombobox*Listbox.background", "#2B2B2B")
        self.option_add("*TCombobox*Listbox.foreground", "white")
        self.option_add("*TCombobox*Listbox.selectBackground", "#1F6AA5")
        self.option_add("*TCombobox*Listbox.selectForeground", "white")
        self.option_add("*TCombobox*Listbox.font", (UI_FONT, 12))

        self.person_cb = ttk.Combobox(left_outer, font=(UI_FONT, 12), state="readonly", width=15)
        self.person_cb.pack(padx=10, pady=5)
        self.person_cb.set("— 請選擇受檢者 —")
        self.person_cb.bind("<<ComboboxSelected>>", lambda e: self.search_entry.focus_set())

        # 將原本受檢者選單下方的留白改為近期寫入紀錄，方便連續核對操作。
        last_write_frame = ctk.CTkFrame(left_outer, corner_radius=8, fg_color="#1E2C3A")
        last_write_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(14, 10))
        ctk.CTkLabel(
            last_write_frame,
            text="🧾 寫入紀錄",
            font=(UI_FONT, 12, "bold"),
            text_color="#5BC0DE",
        ).pack(anchor="w", padx=9, pady=(9, 3))
        self.last_write_history_box = ctk.CTkTextbox(
            last_write_frame,
            font=(UI_FONT, 11, "bold"),
            text_color="#D9E2EC",
            fg_color="#152231",
            corner_radius=6,
            wrap="word",
            activate_scrollbars=False,
        )
        self.last_write_history_box.pack(fill=tk.BOTH, expand=True, padx=9, pady=(0, 9))
        self.last_write_history_box.configure(state="disabled")
        self.last_write_history_box.bind("<Configure>", self._schedule_last_write_history_render)
        self._schedule_last_write_history_render(delay=0)

        mid_outer = ctk.CTkFrame(main_frame, corner_radius=10, fg_color="#2B2B2B")
        mid_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        ctk.CTkLabel(mid_outer, text="🔎 搜尋與勾選項目", font=(UI_FONT, 13, "bold"), text_color="#5BC0DE").pack(pady=(12, 5))

        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(mid_outer, textvariable=self.search_var, font=(UI_FONT, 12), placeholder_text="輸入縮寫或名稱搜尋...")
        self.search_entry.pack(fill=tk.X, padx=12, pady=5)
        self.search_var.trace_add("write", self._trace_search)

        category_frame = ctk.CTkFrame(mid_outer, fg_color="transparent")
        category_frame.pack(fill="x", padx=12, pady=(2, 2))
        ctk.CTkLabel(category_frame, text="分類:", font=(UI_FONT, 11, "bold")).pack(side="left")
        category_values = ["全部類別"] + sorted({
            data.get("test_category") or "未分類"
            for data in self.db_items.values()
        })
        self.category_filter_var = ctk.StringVar(value="全部類別")
        self.category_filter_combo = ctk.CTkOptionMenu(
            category_frame,
            values=category_values,
            variable=self.category_filter_var,
            command=self._trace_search,
            height=26,
            font=(UI_FONT, 11),
        )
        self.category_filter_combo.pack(side="left", padx=(5, 0), fill="x", expand=True)

        sort_panda_frame = ctk.CTkFrame(mid_outer, fg_color="transparent")
        sort_panda_frame.pack(fill="x", padx=12, pady=(2, 5))
        ctk.CTkLabel(sort_panda_frame, text="排序方式:", font=(UI_FONT, 11, "bold")).pack(side="left")
        self.panda_sort_var = ctk.StringVar(value="預設 (原始順序)")
        self.panda_sort_combo = ctk.CTkOptionMenu(
            sort_panda_frame, values=["預設 (原始順序)", "系統代碼", "中/英文名稱", "英文簡寫", "歸屬單位", "檢驗類別"],
            variable=self.panda_sort_var, command=self._trace_search, height=26, font=(UI_FONT, 11)
        )
        self.panda_sort_combo.pack(side="left", padx=(5, 0), fill="x", expand=True)

        self.scrollable_frame = ctk.CTkScrollableFrame(mid_outer, fg_color="#1E1E1E", corner_radius=8, height=180)
        self.scrollable_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        def _forward_mousewheel(event):
            try: self.scrollable_frame._parent_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception: pass
        self.scrollable_frame.bind("<MouseWheel>", _forward_mousewheel)

        right_outer = ctk.CTkFrame(main_frame, corner_radius=10, fg_color="#2B2B2B")
        right_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ctk.CTkLabel(right_outer, text="📋 已勾選項目", font=(UI_FONT, 13, "bold"), text_color="#1F6AA5").pack(pady=(12, 5))
        
        summary_toolbar = ctk.CTkFrame(right_outer, fg_color="transparent")
        summary_toolbar.pack(fill="x", padx=12, pady=(0, 6))
        self.selection_count_label = ctk.CTkLabel(
            summary_toolbar,
            text="已勾選 0 項",
            font=(UI_FONT, 11, "bold"),
            text_color="#AAAAAA",
        )
        self.selection_count_label.pack(side=tk.LEFT)
        self.clear_selection_btn = ctk.CTkButton(
            summary_toolbar,
            text="清除勾選",
            command=self.clear_selected_items,
            font=(UI_FONT, 10, "bold"),
            height=25,
            width=82,
            fg_color="#555555",
            hover_color="#444444",
        )
        self.clear_selection_btn.pack(side=tk.RIGHT)

        self.summary_text = ctk.CTkTextbox(right_outer, font=(UI_FONT, 12, "bold"), fg_color="#1E1E1E", text_color="#5BC0DE")
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
        self.summary_text.configure(state="disabled")

        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 15), padx=15)
        self.result_status_label = ctk.CTkLabel(
            bottom_frame,
            text="請先選擇 Excel 目標、受檢者與項目",
            font=(UI_FONT, 11, "bold"),
            text_color="#AAAAAA",
            anchor="w",
        )
        self.result_status_label.pack(fill=tk.X, pady=(0, 4))
        bottom_actions = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        bottom_actions.pack(fill=tk.X)
        
        self.lock_var = tk.BooleanVar()
        self.lock_cb = ctk.CTkCheckBox(bottom_actions, text="🔒 鎖定勾選 (連續加選)", variable=self.lock_var, font=(UI_FONT, 12, "bold"), text_color="#E91E63", fg_color="#E91E63", hover_color="#C2185B")
        self.lock_cb.pack(side=tk.RIGHT, padx=10) 

        self.btn = ctk.CTkButton(bottom_actions, text="⚡ 寫入 Excel", command=self.write_to_excel, bg_color="transparent", fg_color="#28A745", hover_color="#218838", font=(UI_FONT, 14, "bold"), height=38)
        self.btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10)) 

        def _on_down(e):
            if self.current_focus_idx < len(self.visible_checkboxes) - 1: self.current_focus_idx += 1
            self.highlight_cb(); return "break"
        def _on_up(e):
            if self.current_focus_idx > 0: self.current_focus_idx -= 1
            elif self.current_focus_idx == 0: self.current_focus_idx = -1; self.search_entry.focus_set()
            self.highlight_cb(); return "break"
        def _prevent_cb(e):
            if e.keysym == 'Down': return _on_down(e)
            elif e.keysym == 'Up': return _on_up(e)
            return "break"
        
        def _toggle(e):
            if e.widget == self.search_entry and e.keysym == 'space': return 
            if e.widget == self.person_cb: self.search_entry.focus_set(); return "break"
            if self.current_focus_idx == -1 and e.widget == self.search_entry and self.visible_checkboxes:
                self.visible_checkboxes[0].toggle()
                return "break"
            if 0 <= self.current_focus_idx < len(self.visible_checkboxes): 
                self.visible_checkboxes[self.current_focus_idx].toggle() 
            return "break"

        self.bind("<Down>", _on_down)
        self.bind("<Up>", _on_up)
        self.bind("<Return>", _toggle)
        self.bind("<space>", _toggle) 
        self.person_cb.bind("<Down>", _prevent_cb)
        self.person_cb.bind("<Up>", _prevent_cb)

        for item in self.items_list:
            var = tk.BooleanVar()
            def make_on_write(it, v):
                def on_write(*args):
                    try:
                        if v.get():
                            if it in self.selection_order: self.selection_order.remove(it)
                            self.selection_order.append(it)
                        else:
                            if it in self.selection_order: self.selection_order.remove(it)
                    except Exception: pass
                    try: self._trace_summary()
                    except Exception: pass
                return on_write

            var.trace_add("write", make_on_write(item, var))
            self.item_vars[item] = var
            cb = ctk.CTkCheckBox(self.scrollable_frame, text=item, variable=var, font=(UI_FONT, 12))
            cb.bind("<MouseWheel>", _forward_mousewheel) 
            self.item_checkboxes[item] = cb

        self._trace_search()
        self._trace_summary()
        self.after(500, self.refresh_excel_context)

# =========================================================================
# 2. 視覺化圓角介面 (GUI) 與主程式 (排表母艦)
# =========================================================================
