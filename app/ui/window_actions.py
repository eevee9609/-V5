"""主視窗的外觀、快捷鍵、院所側欄與設定視窗控制。"""

import json
import os
import sys
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from app.config import (
    APP_ROOT,
    HEALTH_CHECK_ITEMS_FILE,
    ICON_PATH,
    INSTITUTION_DB_FILE,
    UI_FONT,
)
from app.data import (
    all_item_names,
    find_db_item_robust,
    get_sorted_items_list,
    load_health_check_database,
    load_institution_database,
    save_institution_database,
)
from app.ui.assistant_tool import SmartAssistantTool
from app.ui.layout import size_settings_window
from app.ui.tab_names import TAB_PUBLIC, TAB_SELF_PAID, TAB_SPLIT


CURRENT_DIR = str(APP_ROOT)
json_file = str(HEALTH_CHECK_ITEMS_FILE)
INST_DB_FILE = str(INSTITUTION_DB_FILE)
ICON_PATH = str(ICON_PATH)
find_db_item_robust_global = find_db_item_robust


def load_json_database_directly(silent=True):
    return load_health_check_database(
        silent,
        on_warning=messagebox.showwarning,
        on_error=messagebox.showerror,
    )


class WindowActionsMixin:

    def format_opacity_label(self, alpha):
        return f"{int(round(alpha * 100))}%"

    def update_window_opacity(self, value=None):
        try:
            alpha = float(self.opacity_var.get() if value is None else value)
            alpha = max(0.35, min(1.0, alpha))
            self.current_alpha = alpha
            self.attributes("-alpha", alpha)
            if hasattr(self, "opacity_label") and self.opacity_label:
                self.opacity_label.configure(text=self.format_opacity_label(alpha))
            if hasattr(self, "assistant_tool") and self.assistant_tool and self.assistant_tool.winfo_exists():
                self.assistant_tool.set_opacity(alpha)
        except Exception: pass

    def reset_for_next_event(self, event=None):
        """清除本場排表資料，保留院所、快捷按鈕與系統設定。"""
        public_count = len(getattr(self, "public_items", []))
        self_paid_count = sum(len(items) for items in getattr(self, "self_paid_config", {}).values())
        roster_count = len(getattr(self, "parsed_roster_data", []))
        message = (
            "初始化下一場將清除目前的：\n"
            f"・公費排定項目：{public_count} 項\n"
            f"・自費排定項目：{self_paid_count} 項\n"
            f"・已匯入受檢者：{roster_count} 人\n\n"
            "啟用OB會取消勾選；院所資料、快捷按鈕與系統設定會保留。\n"
            "確定要繼續嗎？"
        )
        if not messagebox.askyesno("初始化下一場", message, default="no", parent=self):
            return "break" if event is not None else None

        self.public_items.clear()
        self.self_paid_config.clear()
        self.parsed_roster_data.clear()
        self.roster_checkbox_vars.clear()
        if hasattr(self, "enable_ob_var"):
            self.enable_ob_var.set(False)

        if hasattr(self, "listbox_preview_pub"):
            self.update_preview_pub()
        if hasattr(self, "listbox_preview_self"):
            self.update_preview_self()
        for listbox_name in ("listbox_pub", "listbox_self", "listbox_preview_pub", "listbox_preview_self"):
            listbox = getattr(self, listbox_name, None)
            if listbox is not None:
                try:
                    listbox.selection_clear(0, tk.END)
                except tk.TclError:
                    pass
        if hasattr(self, "entry_pkg"):
            self.entry_pkg.delete(0, tk.END)
            self.entry_pkg.insert(0, "A")
        if hasattr(self, "entry_l_pkg"):
            self.entry_l_pkg.delete(0, tk.END)
        if hasattr(self, "lbl_roster_count"):
            self.lbl_roster_count.configure(text="(目前尚未匯入名單)", text_color="#AAAAAA")
        if hasattr(self, "tabview"):
            self.tabview.set(TAB_PUBLIC)

        assistant = getattr(self, "assistant_tool", None)
        try:
            if assistant is not None and assistant.winfo_exists():
                for var in getattr(assistant, "item_vars", {}).values():
                    var.set(False)
                assistant.selection_order.clear()
                assistant.history_stack.clear()
                assistant.last_write_result = None
                if hasattr(assistant, "clear_last_write_history"):
                    assistant.clear_last_write_history()
                if hasattr(assistant, "undo_btn"):
                    assistant.undo_btn.configure(state="disabled")
                if hasattr(assistant, "refresh_person_selector"):
                    assistant.refresh_person_selector(preserve_selection=False)
                if hasattr(assistant, "_set_result_status"):
                    assistant._set_result_status("已初始化下一場，請重新選擇受檢者與項目", "#5BC0DE")
        except Exception:
            pass

        if hasattr(self, "inst_status_lbl"):
            self.inst_status_lbl.configure(text="✅ 已初始化下一場，可重新排定項目", text_color="#5BC0DE")
            self.after(3500, lambda: self.inst_status_lbl.configure(text=""))
        return "break" if event is not None else None

    def summon_panda(self):
        if hasattr(self, "assistant_tool") and self.assistant_tool is not None and self.assistant_tool.winfo_exists():
            self.assistant_tool.deiconify()
            self.assistant_tool.lift()
            self.assistant_tool.focus_force()
            if hasattr(self.assistant_tool, "panda") and self.assistant_tool.panda.top.state() == "withdrawn":
                self.assistant_tool.panda.top.deiconify()
            return
        self.assistant_tool = SmartAssistantTool(self)
        self.assistant_tool.lift()

    def load_institution_db(self):
        existing_data = load_institution_database()
        if existing_data:
            return existing_data
                
        default_db = {
            "A001": {
                "name": "杏聯測試診所",
                "remark": "【注意事項】\n1. 報告需加急處理 (三天內)。\n2. 統一不收現金，由院所月結。\n3. CEA 試管請務必使用紫頭管。",
                "shortcuts": [
                    {"label": "📋 載入常用基本包", "items": ["白血球(001)", "紅血球(002)", "心電圖"]}
                ]
            }
        }
        self.save_institution_db_to_disk(default_db)
        return default_db

    def save_institution_db_to_disk(self, db_data=None):
        if db_data is None: db_data = self.inst_db
        try:
            save_institution_database(db_data)
        except Exception as e: messagebox.showerror("存檔失敗", f"無法儲存院所資料：{e}")

    def build_right_panel(self):
        self.right_panel = ctk.CTkFrame(self, width=370, fg_color="#2b2b2b")
        
        header_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=(10, 0))
        
        self.right_title = ctk.CTkLabel(header_frame, text="🏥 院所資訊與整理", font=(UI_FONT, 16, "bold"), text_color="#61afef")
        self.right_title.pack(side="left", padx=5)
        
        close_btn = ctk.CTkButton(header_frame, text="✖ 關閉", width=50, fg_color="transparent", 
                                  hover_color="#c94f4f", command=self.close_right_panel)
        close_btn.pack(side="right")
        
        remark_header = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        remark_header.pack(fill="x", padx=15, pady=(15, 0))
        ctk.CTkLabel(remark_header, text="📋 專屬備註事項:", font=(UI_FONT, 14, "bold")).pack(side="left")
        
        self.save_remark_btn = ctk.CTkButton(remark_header, text="💾 儲存備註", width=80, height=24,
                                              fg_color="#2b7b5c", hover_color="#1e5c45",
                                              command=self.save_current_remark)
        self.save_remark_btn.pack(side="right")

        self.remark_textbox = ctk.CTkTextbox(self.right_panel, height=140, font=(UI_FONT, 13))
        self.remark_textbox.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkFrame(self.right_panel, height=2, fg_color="#444444").pack(fill="x", padx=15, pady=10)
        
        shortcut_header = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        shortcut_header.pack(fill="x", padx=15, pady=(0, 5))
        ctk.CTkLabel(shortcut_header, text="⚡ 常見快捷鍵 / 常用套組:", font=(UI_FONT, 14, "bold")).pack(side="left")
        
        edit_btn = ctk.CTkButton(shortcut_header, text="⚙️ 編輯/新增", width=80, height=22, fg_color="#444444", hover_color="#666666", command=lambda: self.open_settings_window("clinic"))
        edit_btn.pack(side="right")

        self.shortcut_container = ctk.CTkScrollableFrame(self.right_panel, height=240, fg_color="#222222")
        self.shortcut_container.pack(fill="both", expand=True, padx=15, pady=(5, 15))

    def auto_check_inst(self, *args):
        raw_val = self.inst_var.get().strip()
        inst_code = raw_val.split(" - ")[0].strip().upper()
        if inst_code in self.inst_db: self.search_and_open_panel(auto_triggered=True)

    def search_and_open_panel(self, auto_triggered=False):
        raw_val = self.inst_var.get().strip()
        inst_code = raw_val.split(" - ")[0].strip().upper()
        
        if not inst_code:
            if not auto_triggered: messagebox.showwarning("提示", "請先輸入或選擇代檢院所代號！")
            return
            
        self.current_inst_code = inst_code
        if hasattr(self, "refresh_routing_visuals"):
            self.refresh_routing_visuals()
        
        if not self.right_panel.winfo_ismapped():
            self.geometry(f"{self.expanded_width}x{self.window_height}")
            self.main_content_frame.pack_forget()
            self.right_panel.pack(side="right", fill="y", padx=(0, 10), pady=10)
            self.main_content_frame.pack(side="left", fill="both", expand=True)
            
        self.remark_textbox.delete("1.0", "end")
        
        if inst_code in self.inst_db:
            inst_data = self.inst_db[inst_code]
            raw_name = inst_data.get('name', '').strip()
            
            if raw_name:
                display_name = raw_name if raw_name.startswith("院所_") else f"院所_{raw_name}"
                title_text = f"🏥 {display_name} ({inst_code})"
            else: title_text = f"🏥 院所_{inst_code} ({inst_code})"
                
            self.right_title.configure(text=title_text)
            self.remark_textbox.insert("end", inst_data.get('remark', ''))
            self.render_shortcut_buttons(inst_data.get('shortcuts', []))
        else:
            self.right_title.configure(text=f"🏥 院所_{inst_code} ({inst_code})")
            self.remark_textbox.insert("end", "此院所尚未建檔。\n可在下方點擊『⚙️ 編輯/新增』設定名稱與備註！")
            self.render_shortcut_buttons([])

    def save_current_remark(self):
        inst_code = self.current_inst_code
        if not inst_code:
            messagebox.showwarning("提示", "請先輸入院所代號！")
            return
            
        new_remark = self.remark_textbox.get("1.0", "end").strip()
        if inst_code not in self.inst_db: self.inst_db[inst_code] = {"name": "", "remark": new_remark, "shortcuts": []}
        else: self.inst_db[inst_code]["remark"] = new_remark
            
        self.save_institution_db_to_disk()
        self.save_remark_btn.configure(text="✓ 已儲存", fg_color="#1b5e20")
        self.after(1800, lambda: self.save_remark_btn.configure(text="💾 儲存備註", fg_color="#2b7b5c"))

    def render_shortcut_buttons(self, shortcuts):
        for widget in self.shortcut_container.winfo_children(): widget.destroy()
            
        if not shortcuts:
            ctk.CTkLabel(self.shortcut_container, text="尚無設定快捷按鈕\n點選右上『⚙️ 編輯/新增』即可自由加入！", text_color="gray").pack(pady=20)
            return

        for idx, sc in enumerate(shortcuts):
            label = sc.get("label", "快捷選項")
            items = sc.get("items", [])
            
            btn = ctk.CTkButton(
                self.shortcut_container, text=label, height=40, font=(UI_FONT, 14, "bold"),
                fg_color="#3a5a78", hover_color="#2b4259",
                command=lambda name=label, itms=items: self.on_shortcut_clicked(name, itms)
            )
            btn.pack(fill="x", padx=10, pady=6)
            
            btn.bind("<Button-3>", lambda event, i=idx, sc_obj=sc: self.show_shortcut_menu(event, i, sc_obj))
            if sys.platform == "darwin": btn.bind("<Button-2>", lambda event, i=idx, sc_obj=sc: self.show_shortcut_menu(event, i, sc_obj))

    def show_shortcut_menu(self, event, idx, sc_obj):
        try:
            menu = tk.Menu(self, tearoff=0, font=(UI_FONT, 11))
            menu.add_command(label=f"📝 編輯套組 [{sc_obj.get('label')}]", command=lambda: self.open_shortcut_editor(idx, sc_obj))
            menu.add_separator()
            menu.add_command(label=f"🗑️ 刪除套組 [{sc_obj.get('label')}]", command=lambda: self.delete_shortcut_from_menu(idx, sc_obj.get('label')))
            menu.tk_popup(event.x_root, event.y_root)
        finally: menu.grab_release()

    def open_shortcut_editor(self, idx, sc_obj):
        self.target_edit_idx = idx 
        self.open_settings_window("clinic")

    def delete_shortcut_from_menu(self, idx, label):
        code = self.current_inst_code
        if not code or code not in self.inst_db: return
        
        if messagebox.askyesno("刪除確認", f"⚠️ 確定要刪除快捷套組「{label}」嗎？\n\n刪除後無法復原喔！"):
            if "shortcuts" in self.inst_db[code] and len(self.inst_db[code]["shortcuts"]) > idx:
                self.inst_db[code]["shortcuts"].pop(idx)
                self.save_institution_db_to_disk()
                self.search_and_open_panel(auto_triggered=True)
                messagebox.showinfo("刪除成功", f"套組「{label}」已成功移除！")
                
                if hasattr(self, "settings_window") and self.settings_window.winfo_exists():
                    self.switch_settings_tab("clinic")

    def on_shortcut_clicked(self, label, items):
        if not items:
            messagebox.showwarning("快捷按鈕", f"【{label}】尚未設定任何項目。")
            return

        added_count = 0
        current_tab = str(self.tabview.get() or "").strip()
        # 分頁文字統一由 tab_names 管理；名稱有前後空白或圖示異動時，仍以核心名稱辨識。
        is_self_paid_tab = current_tab == TAB_SELF_PAID or "自費套組設定" in current_tab
        is_public_or_split_tab = (
            current_tab in (TAB_PUBLIC, TAB_SPLIT)
            or "公費常規項目" in current_tab
            or "自動拆分工作檔案" in current_tab
        )

        if not is_self_paid_tab and not is_public_or_split_tab:
            messagebox.showwarning(
                "快捷按鈕",
                "請先切換至「公費常規項目」或「自費套組設定」後，再套用快捷按鈕。",
            )
            return
        
        for item in items:
            display_label = item
            for global_item in all_item_names:
                if item in global_item or global_item.startswith(item):
                    display_label = global_item; break
            
            # 判斷目前停留的分頁，智慧決定加入到哪裡
            if is_self_paid_tab:
                pkg = self.entry_pkg.get().strip().upper()
                if not pkg:
                    pkg = "A"  # 預設給一個套組名稱防呆
                    self.entry_pkg.delete(0, "end")
                    self.entry_pkg.insert(0, "A")
                
                if pkg not in self.self_paid_config:
                    self.self_paid_config[pkg] = []
                    
                if not any(x["item"] == display_label for x in self.self_paid_config[pkg]):
                    self.self_paid_config[pkg].append({
                        "item": display_label, 
                        "gender": "All",
                        "check_l_col": False, 
                        "l_pkg": ""
                    })
                    added_count += 1
            else:
                # 預設加入公費 (如果在公費分頁，或拆分分頁)
                if not any(x["item"] == display_label for x in self.public_items):
                    self.public_items.append({"item": display_label, "gender": "All"})
                    added_count += 1
                
        if is_self_paid_tab:
            # 即使項目已存在，也同步重新整理右側清單，避免畫面與資料不同步。
            self.update_preview_self()
            destination = f"自費套組 {pkg}"
        else:
            self.update_preview_pub()
            destination = "公費常規項目"
            if current_tab != TAB_PUBLIC:
                self.tabview.set(TAB_PUBLIC)

        if added_count:
            status_text = f"✅ 已將 {added_count} 項套用至 {destination}"
            status_color = "#28A745"
        else:
            status_text = f"ℹ️ 【{label}】的項目已在 {destination} 中"
            status_color = "#5BC0DE"
        self.inst_status_lbl.configure(text=status_text, text_color=status_color)
        self.after(4000, lambda: self.inst_status_lbl.configure(text=""))

    def close_right_panel(self):
        self.right_panel.pack_forget()
        self.geometry(f"{self.normal_width}x{self.window_height}")

    def open_settings_window(self, default_tab="db"):
        if hasattr(self, "settings_window") and self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.lift()
            self.switch_settings_tab(default_tab)
            return

        self.settings_window = ctk.CTkToplevel(self)
        self.settings_window.title("⚙️ 系統設定中心")
        size_settings_window(self.settings_window, self)
        self.settings_window.transient(self)
        self.settings_window.grab_set()

        # ⭐ 載入應用程式圖示
        if os.path.exists(ICON_PATH):
            try:
                self.settings_window.iconbitmap(ICON_PATH)
            except Exception:
                pass

        self.settings_window.grid_rowconfigure(0, weight=1)
        self.settings_window.grid_columnconfigure(1, weight=1)

        sidebar = ctk.CTkFrame(self.settings_window, width=220, corner_radius=0, fg_color="#2B2B2B")
        sidebar.grid(row=0, column=0, sticky="nswe")
        sidebar.grid_rowconfigure(6, weight=1)

        title_label = ctk.CTkLabel(sidebar, text="⚙️ 設定選單", font=(UI_FONT, 22, "bold"), text_color="white")
        title_label.grid(row=0, column=0, padx=20, pady=(30, 30))

        self.btn_tab_db = ctk.CTkButton(sidebar, text="🗄️ 資料庫維護", font=(UI_FONT, 16, "bold"), height=45, corner_radius=8, anchor="w", command=lambda: self.switch_settings_tab("db"))
        self.btn_tab_db.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")

        self.btn_tab_clinic = ctk.CTkButton(sidebar, text="🏥 院所管理中心", font=(UI_FONT, 16, "bold"), height=45, corner_radius=8, anchor="w", command=lambda: self.switch_settings_tab("clinic"))
        self.btn_tab_clinic.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="ew")

        self.btn_tab_announce = ctk.CTkButton(sidebar, text="📢 系統公告設定", font=(UI_FONT, 16, "bold"), height=45, corner_radius=8, anchor="w", command=lambda: self.switch_settings_tab("announce"))
        self.btn_tab_announce.grid(row=3, column=0, padx=15, pady=(0, 10), sticky="ew")

        self.btn_tab_special = ctk.CTkButton(sidebar, text="🧪 公式與進階綁定", font=(UI_FONT, 16, "bold"), height=45, corner_radius=8, anchor="w", command=lambda: self.switch_settings_tab("special"))
        self.btn_tab_special.grid(row=4, column=0, padx=15, pady=(0, 10), sticky="ew")

        self.settings_content_frame = ctk.CTkFrame(self.settings_window, fg_color="transparent")
        self.settings_content_frame.grid(row=0, column=1, sticky="nswe", padx=14, pady=12)

        self.current_settings_tab_frame = None
        self.switch_settings_tab(default_tab)

    def switch_settings_tab(self, tab_name):
        if hasattr(self, "current_settings_tab_frame") and self.current_settings_tab_frame:
            self.current_settings_tab_frame.destroy()

        default_color = "transparent"
        active_color = "#1F6AA5" 

        self.btn_tab_db.configure(fg_color=default_color if tab_name != "db" else active_color)
        self.btn_tab_clinic.configure(fg_color=default_color if tab_name != "clinic" else active_color)
        self.btn_tab_announce.configure(fg_color=default_color if tab_name != "announce" else active_color)
        if hasattr(self, 'btn_tab_special'):
            self.btn_tab_special.configure(fg_color=default_color if tab_name != "special" else active_color)

        self.current_settings_tab_frame = ctk.CTkFrame(self.settings_content_frame, fg_color="transparent")
        self.current_settings_tab_frame.pack(fill="both", expand=True)

        if tab_name == "db": self.build_db_maintenance_ui(self.current_settings_tab_frame)
        elif tab_name == "clinic": self.build_clinic_management_ui(self.current_settings_tab_frame)
        elif tab_name == "announce": self.build_announcement_ui(self.current_settings_tab_frame)
        elif tab_name == "special": self.build_special_formula_ui(self.current_settings_tab_frame)
