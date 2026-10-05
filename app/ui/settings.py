"""系統設定、院所管理、資料庫維護與公告設定頁面。"""

import json
import os
import re
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import openpyxl
import pandas as pd

from app.config import (
    APP_ROOT,
    HEALTH_CHECK_ITEMS_FILE,
    ICON_PATH,
    INSTITUTION_DB_FILE,
    SYSTEM_CONFIG_FILE,
    UI_FONT,
)
from app.data import (
    all_item_names,
    find_db_item_robust,
    get_sorted_items_list,
    load_health_check_database,
)
from app.data.json_storage import write_json_with_backup
from app.services.column_routing import analyze_routing_rules, make_precise_routing_rule
from app.ui.layout import fit_listbox_rows
from app.utils.sorting import safe_chinese_sort_key, safe_natural_sort_key


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


class SettingsMixin:

    def build_db_maintenance_ui(self, parent):
        left_frame = ctk.CTkFrame(parent, corner_radius=15)
        left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        right_frame = ctk.CTkFrame(parent, corner_radius=15)
        right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(left_frame, text=" 📚 目前大腦資料庫列表 ", font=self.title_font).pack(pady=(10, 5))

        self.db_search_var = tk.StringVar()
        self.entry_search_db = ctk.CTkEntry(left_frame, textvariable=self.db_search_var, font=self.default_font, placeholder_text="🔍 搜尋代碼或名稱...", height=35)
        self.entry_search_db.pack(fill="x", padx=15, pady=5)

        db_action_bar = ctk.CTkFrame(left_frame, fg_color="transparent")
        db_action_bar.pack(fill="x", padx=15, pady=(0, 5))
        
        btn_import_db = ctk.CTkButton(db_action_bar, text="📥 從 Excel/CSV 快速匯入", font=(UI_FONT, 12, "bold"), fg_color="#0275D8", hover_color="#025AA5", command=self.batch_import_db_from_excel)
        btn_import_db.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        btn_export_db = ctk.CTkButton(db_action_bar, text="📤 匯出備份", font=(UI_FONT, 12, "bold"), fg_color="#6C757D", hover_color="#5A6268", command=self.export_db_backup)
        btn_export_db.pack(side="right", expand=True, fill="x", padx=(5, 0))

        list_frame_db = ctk.CTkFrame(left_frame, fg_color="transparent")
        list_frame_db.pack(fill="both", expand=True, padx=15, pady=5)
        scroll_db = ctk.CTkScrollbar(list_frame_db)
        scroll_db.pack(side="right", fill="y")

        self.listbox_db = tk.Listbox(list_frame_db, font=(UI_FONT, 13), yscrollcommand=scroll_db.set,
                                     bg="#2b2b2b", fg="white", relief="flat", highlightthickness=0, selectbackground="#1f538d", selectmode="extended")
        self.listbox_db.pack(side="left", fill="both", expand=True)
        scroll_db.configure(command=self.listbox_db.yview)
        fit_listbox_rows(self.listbox_db, scroll_db)

        ctk.CTkLabel(right_frame, text=" 📝 新增 / 編輯單筆檢驗項目 ", font=self.title_font).pack(pady=(10, 5))

        form_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        form_frame.pack(fill="x", padx=20)

        ctk.CTkLabel(form_frame, text="1. 系統代碼 (唯一值，如 002):", font=self.default_font).pack(anchor="w", pady=(5, 0))
        self.entry_db_code = ctk.CTkEntry(form_frame, font=self.default_font, height=35)
        self.entry_db_code.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(form_frame, text="2. 中/英文名稱 (如 大腸癌胚抗原CEA 或 GA):", font=self.default_font).pack(anchor="w", pady=(5, 0))
        self.entry_db_zh = ctk.CTkEntry(form_frame, font=self.default_font, height=35)
        self.entry_db_zh.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(form_frame, text="3. 英文簡寫 (如 CEA(AU)):", font=self.default_font).pack(anchor="w", pady=(5, 0))
        self.entry_db_en = ctk.CTkEntry(form_frame, font=self.default_font, height=35)
        self.entry_db_en.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(form_frame, text="4. 檢驗類別 (如 血液, 生化 - 可空白):", font=self.default_font).pack(anchor="w", pady=(5, 0))
        self.entry_db_category = ctk.CTkEntry(form_frame, font=self.default_font, height=35)
        self.entry_db_category.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(form_frame, text="5. 歸屬單位:", font=self.default_font).pack(anchor="w", pady=(5, 0))
        self.db_org_var = ctk.StringVar(value="C2")
        org_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        org_frame.pack(fill="x", pady=(0, 5))
        ctk.CTkRadioButton(org_frame, text="C2 (杏聯)", variable=self.db_org_var, value="C2", font=self.default_font).pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(org_frame, text="B2 (博仁)", variable=self.db_org_var, value="B2", font=self.default_font).pack(side="left")

        btn_action_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        btn_action_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkButton(btn_action_frame, text="💾 儲存 / 更新至大腦", fg_color="#28A745", hover_color="#218838", font=(UI_FONT, 14, "bold"), height=45, command=self.save_to_json_db).pack(side="left", expand=True, fill="x", padx=(0, 5))
        ctk.CTkButton(btn_action_frame, text="🗑️ 批次刪除選取", fg_color="#D9534F", hover_color="#C9302C", font=(UI_FONT, 14, "bold"), height=45, command=self.delete_from_json_db).pack(side="right", fill="x", padx=(5, 0))

        self.listbox_db.bind("<<ListboxSelect>>", self.on_db_listbox_select)
        self.db_search_var.trace("w", self.refresh_db_listbox)
        self.refresh_db_listbox()

    def build_clinic_management_ui(self, parent_frame):
        win = parent_frame
        target_code = self.current_inst_code if self.current_inst_code else self.inst_var.get().split(" - ")[0].strip().upper()
        if not target_code: target_code = "H101"
            
        editing_idx = {"index": None}
        selected_temp_items = []

        selector_frame = ctk.CTkFrame(win, fg_color="#1a202c")
        selector_frame.pack(fill="x", padx=10, pady=(6, 5))
        
        ctk.CTkLabel(selector_frame, text="📋 快速選擇要管理的院所:", font=(UI_FONT, 14, "bold"), text_color="#61afef").pack(side="left", padx=12, pady=10)
        
        def build_combo_options():
            opts = []
            for c, d in self.inst_db.items():
                n = d.get('name', '').strip()
                opts.append(f"{c} - {n}" if n else f"{c}")
            opts.append("➕ 新增/新建院所代號")
            return opts

        inst_selector = ctk.CTkComboBox(selector_frame, values=build_combo_options(), width=260, font=(UI_FONT, 13))
        inst_selector.pack(side="left", padx=5)

        manager_body = ctk.CTkFrame(win, fg_color="transparent")
        manager_body.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        top_frame = ctk.CTkFrame(manager_body, fg_color="#2b2b2b")
        top_frame.pack(fill="x", pady=(0, 8))
        
        r1 = ctk.CTkFrame(top_frame, fg_color="transparent")
        r1.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(r1, text="🏥 院所代號:", font=(UI_FONT, 14, "bold")).pack(side="left", padx=(5, 2))
        code_entry = ctk.CTkEntry(r1, width=90, font=(UI_FONT, 14, "bold"))
        code_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(r1, text="院所名稱:", font=(UI_FONT, 14)).pack(side="left", padx=(15, 2))
        name_entry = ctk.CTkEntry(r1, width=160, placeholder_text="例：宏仁")
        name_entry.pack(side="left", padx=5)

        btn_save_info = ctk.CTkButton(r1, text="💾 儲存", width=80, fg_color="#2b7b5c", hover_color="#1e5c45")
        btn_save_info.pack(side="right", padx=5)
        
        btn_delete_inst = ctk.CTkButton(r1, text="🗑️ 刪除", width=80, fg_color="#D9534F", hover_color="#C9302C")
        btn_delete_inst.pack(side="right", padx=5)

        r2 = ctk.CTkFrame(top_frame, fg_color="transparent")
        r2.pack(fill="x", padx=10, pady=(0, 8))
        
        ctk.CTkLabel(r2, text="📋 專屬備註:", font=(UI_FONT, 13)).pack(anchor="w", padx=5, pady=(2, 2))
        win_remark_box = ctk.CTkTextbox(r2, height=70, font=(UI_FONT, 12))
        win_remark_box.pack(fill="x", padx=5, pady=2)

        # 使用自訂的瀏覽器式分頁列，按鈕從左側開始排列；避免內建 Tabview
        # 將選單置中後，在內容較寬時不易辨識目前所在頁面。
        management_tab_area = ctk.CTkFrame(manager_body, fg_color="transparent")
        management_tab_area.pack(fill="both", expand=True)
        browser_tab_bar = ctk.CTkFrame(management_tab_area, height=42, fg_color="#1F2937", corner_radius=8)
        browser_tab_bar.pack(fill="x", pady=(0, 6))
        browser_tab_bar.pack_propagate(False)

        management_tab_content = ctk.CTkFrame(management_tab_area, fg_color="transparent")
        management_tab_content.pack(fill="both", expand=True)
        shortcut_tab = ctk.CTkFrame(management_tab_content, fg_color="transparent")
        routing_tab = ctk.CTkFrame(management_tab_content, fg_color="transparent")
        manager_tab_frames = {"shortcuts": shortcut_tab, "routing": routing_tab}
        manager_tab_buttons = {}

        def show_manager_tab(tab_key):
            for frame in manager_tab_frames.values():
                frame.pack_forget()
            manager_tab_frames[tab_key].pack(fill="both", expand=True)
            for key, button in manager_tab_buttons.items():
                is_active = key == tab_key
                button.configure(
                    fg_color="#3182CE" if is_active else "#374151",
                    hover_color="#2B6CB0" if is_active else "#4B5563",
                )

        manager_tab_buttons["shortcuts"] = ctk.CTkButton(
            browser_tab_bar,
            text="⚡ 快捷按鈕管理",
            width=165,
            height=32,
            corner_radius=7,
            font=(UI_FONT, 13, "bold"),
            command=lambda: show_manager_tab("shortcuts"),
        )
        manager_tab_buttons["shortcuts"].pack(side="left", padx=(6, 3), pady=5)
        manager_tab_buttons["routing"] = ctk.CTkButton(
            browser_tab_bar,
            text="🔄 院所專屬分派",
            width=165,
            height=32,
            corner_radius=7,
            font=(UI_FONT, 13, "bold"),
            command=lambda: show_manager_tab("routing"),
        )
        manager_tab_buttons["routing"].pack(side="left", padx=3, pady=5)
        show_manager_tab("shortcuts")

        routing_layout = ctk.CTkFrame(routing_tab, fg_color="transparent")
        routing_layout.pack(fill="both", expand=True)
        routing_layout.grid_columnconfigure(0, weight=2, uniform="routing_columns")
        routing_layout.grid_columnconfigure(1, weight=3, uniform="routing_columns")
        routing_layout.grid_rowconfigure(0, weight=1)
        routing_layout.grid_propagate(False)

        r3 = ctk.CTkFrame(routing_layout, fg_color="#3d2a2a", corner_radius=8)
        r3.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ctk.CTkLabel(
            r3, text="🔄 院所專屬分派規則",
            font=(UI_FONT, 13, "bold"), text_color="#ff9999"
        ).pack(anchor="w", padx=10, pady=(10, 5))

        force_xing_frame = ctk.CTkFrame(r3, fg_color="transparent")
        force_xing_frame.pack(fill="x", padx=10, pady=2)
        ctk.CTkLabel(force_xing_frame, text="強制分派至【杏聯】:", font=(UI_FONT, 13)).pack(anchor="w", padx=5)
        force_xing_entry = ctk.CTkEntry(
            force_xing_frame, width=280,
            placeholder_text="由右側選取項目，或輸入系統:105"
        )
        force_xing_entry.pack(fill="x", padx=5)

        force_bo_frame = ctk.CTkFrame(r3, fg_color="transparent")
        force_bo_frame.pack(fill="x", padx=10, pady=(2, 10))
        ctk.CTkLabel(force_bo_frame, text="強制分派至【博仁】:", font=(UI_FONT, 13)).pack(anchor="w", padx=5)
        force_bo_entry = ctk.CTkEntry(
            force_bo_frame, width=280,
            placeholder_text="由右側選取項目，或輸入英文:GLU"
        )
        force_bo_entry.pack(fill="x", padx=5)

        rule_action_row = ctk.CTkFrame(r3, fg_color="transparent")
        rule_action_row.pack(fill="x", padx=10, pady=(0, 10))
        btn_preview_rules = ctk.CTkButton(
            rule_action_row,
            text="🔎 預覽規則命中",
            width=145,
            height=28,
            fg_color="#805ad5",
            hover_color="#6b46c1",
        )
        btn_preview_rules.pack(anchor="w", padx=5)
        routing_preview_status = ctk.CTkLabel(
            rule_action_row,
            text="尚未設定分派規則",
            font=(UI_FONT, 11),
            text_color="#A0AEC0",
            anchor="w",
        )
        routing_preview_status.pack(fill="x", padx=5, pady=(4, 0))

        def get_routing_catalog_items():
            """由目前資料庫建立去重後的分派規則預覽目錄。"""
            catalog = []
            seen = set()
            for display_label in all_item_names:
                db_item = find_db_item_robust(display_label)
                if not db_item:
                    continue
                key = (
                    db_item.get("internal_code", ""),
                    db_item.get("sys_code", ""),
                    db_item.get("full_name", ""),
                    db_item.get("org_display", ""),
                )
                if key not in seen:
                    catalog.append(db_item)
                    seen.add(key)
            return catalog

        def get_routing_analysis():
            return analyze_routing_rules(
                force_xing_entry.get(),
                force_bo_entry.get(),
                get_routing_catalog_items(),
            )

        def format_routing_item(item):
            code = item.get("internal_code") or item.get("sys_code") or "未設定代碼"
            english = item.get("sys_code", "")
            suffix = f" ({english})" if english and english != code else ""
            return f"[{item.get('org_display', '')}] {code} - {item.get('full_name', '')}{suffix}"

        def refresh_routing_preview_status(*_args):
            xing_rules = force_xing_entry.get().strip()
            boren_rules = force_bo_entry.get().strip()
            if not xing_rules and not boren_rules:
                routing_preview_status.configure(text="尚未設定分派規則", text_color="#A0AEC0")
                return
            analysis = get_routing_analysis()
            conflicts = analysis["conflicts"]
            unmatched = analysis["unmatched_rules"]
            ambiguous = analysis["ambiguous_rules"]
            if conflicts or unmatched:
                routing_preview_status.configure(
                    text=f"❌ 衝突 {len(conflicts)} 項／未命中 {len(unmatched)} 條，不能儲存",
                    text_color="#FC8181",
                )
            elif ambiguous:
                routing_preview_status.configure(
                    text=f"❌ 有 {len(ambiguous)} 條規則命中多個項目，請改用精準選取",
                    text_color="#FC8181",
                )
            else:
                routing_preview_status.configure(
                    text=f"✅ 共會覆寫 {len(analysis['item_destinations'])} 個項目",
                    text_color="#68D391",
                )

        def open_routing_preview():
            analysis = get_routing_analysis()
            preview_window = ctk.CTkToplevel(self.settings_window)
            preview_window.title("🔎 院所專屬分派規則預覽")
            preview_window.geometry("780x620")
            preview_window.transient(self.settings_window)
            preview_window.grab_set()

            ctk.CTkLabel(
                preview_window,
                text="規則只會在下方列出的資料庫項目中生效；請在儲存前確認。",
                font=(UI_FONT, 13),
                text_color="#A0AEC0",
            ).pack(anchor="w", padx=18, pady=(16, 8))

            summary = (
                f"覆寫項目：{len(analysis['item_destinations'])}　"
                f"未命中：{len(analysis['unmatched_rules'])}　"
                f"衝突：{len(analysis['conflicts'])}　"
                f"多項命中：{len(analysis['ambiguous_rules'])}"
            )
            summary_color = "#FC8181" if analysis["has_errors"] else "#F6AD55" if analysis["has_warnings"] else "#68D391"
            ctk.CTkLabel(preview_window, text=summary, font=(UI_FONT, 13, "bold"), text_color=summary_color).pack(anchor="w", padx=18, pady=(0, 8))

            scroll = ctk.CTkScrollableFrame(preview_window, fg_color="#1E1E1E")
            scroll.pack(fill="both", expand=True, padx=18, pady=(0, 12))

            if analysis["unmatched_rules"]:
                ctk.CTkLabel(scroll, text="❌ 完全未命中的規則", font=(UI_FONT, 14, "bold"), text_color="#FC8181").pack(anchor="w", padx=8, pady=(10, 3))
                for record in analysis["unmatched_rules"]:
                    target = "杏聯" if record["destination"] == "xing" else "博仁"
                    ctk.CTkLabel(scroll, text=f"・{target}：{record['rule']}", anchor="w", text_color="#FEB2B2").pack(fill="x", padx=16, pady=1)

            if analysis["conflicts"]:
                ctk.CTkLabel(scroll, text="❌ 同時命中兩側的衝突項目", font=(UI_FONT, 14, "bold"), text_color="#FC8181").pack(anchor="w", padx=8, pady=(10, 3))
                for record in analysis["conflicts"]:
                    ctk.CTkLabel(
                        scroll,
                        text=f"・{format_routing_item(record['item'])}　（杏聯：{record['xing_rule']}／博仁：{record['boren_rule']}）",
                        anchor="w",
                        justify="left",
                        wraplength=700,
                        text_color="#FEB2B2",
                    ).pack(fill="x", padx=16, pady=2)

            for destination, title, color in (("xing", "強制分派至杏聯", "#9AE6B4"), ("boren", "強制分派至博仁", "#90CDF4")):
                records = analysis["rule_matches"][destination]
                if not records:
                    continue
                ctk.CTkLabel(scroll, text=f"🔄 {title}", font=(UI_FONT, 14, "bold"), text_color=color).pack(anchor="w", padx=8, pady=(10, 3))
                for record in records:
                    matches = record["matches"]
                    if not matches:
                        continue
                    ctk.CTkLabel(
                        scroll,
                        text=f"規則「{record['rule']}」命中 {len(matches)} 項：",
                        anchor="w",
                        text_color="#E2E8F0",
                    ).pack(fill="x", padx=16, pady=(3, 1))
                    for item in matches:
                        ctk.CTkLabel(scroll, text=f"　• {format_routing_item(item)}", anchor="w", justify="left", wraplength=680, text_color="#CBD5E0").pack(fill="x", padx=20, pady=1)

            ctk.CTkButton(preview_window, text="關閉", command=preview_window.destroy, width=100).pack(pady=(0, 15))

        btn_preview_rules.configure(command=open_routing_preview)
        force_xing_entry.bind("<FocusOut>", refresh_routing_preview_status)
        force_bo_entry.bind("<FocusOut>", refresh_routing_preview_status)

        # 使用者可先以系統代碼、英文簡碼或名稱片段搜尋資料庫，再把選定項目
        # 寫成「系統:xxx」等精準規則，避免自由文字部分比對到不相干的項目。
        routing_picker = ctk.CTkFrame(routing_layout, fg_color="#29223A", corner_radius=8)
        routing_picker.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        # 清單只使用剩餘高度；下方的目標院所按鈕和狀態列必須固定可見。
        routing_picker.grid_columnconfigure(0, weight=1)
        routing_picker.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(
            routing_picker,
            text="🎯 搜尋並加入分派項目",
            font=(UI_FONT, 12, "bold"),
            text_color="#D6BCFA",
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 2))
        ctk.CTkLabel(
            routing_picker,
            text="輸入代碼或名稱，可多選後加入目標院所。",
            font=(UI_FONT, 11),
            text_color="#A0AEC0",
        ).grid(row=1, column=0, sticky="w", padx=10, pady=(0, 4))

        routing_search_var = tk.StringVar()
        routing_search_entry = ctk.CTkEntry(
            routing_picker,
            textvariable=routing_search_var,
            height=28,
            placeholder_text="例如：105、GLU、血糖",
        )
        routing_search_entry.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 4))

        routing_list_frame = ctk.CTkFrame(routing_picker, fg_color="transparent")
        routing_list_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 4))
        routing_scroll = ctk.CTkScrollbar(routing_list_frame)
        routing_scroll.pack(side="right", fill="y")
        routing_listbox = tk.Listbox(
            routing_list_frame,
            height=5,
            selectmode="extended",
            font=(UI_FONT, 11),
            yscrollcommand=routing_scroll.set,
            bg="#1A202C",
            fg="white",
            selectbackground="#6B46C1",
            relief="flat",
            highlightthickness=0,
            exportselection=False,
        )
        routing_listbox.pack(side="left", fill="both", expand=True)
        routing_scroll.configure(command=routing_listbox.yview)
        fit_listbox_rows(routing_listbox, routing_scroll)
        routing_picker_items = []
        routing_picker_status = ctk.CTkLabel(
            routing_picker,
            text="請先搜尋並選取資料庫項目",
            font=(UI_FONT, 11),
            text_color="#A0AEC0",
            anchor="w",
        )

        def routing_search_text(item):
            return " ".join(
                str(item.get(key, "") or "")
                for key in ("internal_code", "sys_code", "full_name", "org_display")
            ).casefold()

        def refresh_routing_picker(*_args):
            search_term = routing_search_var.get().strip().casefold()
            routing_listbox.delete(0, tk.END)
            routing_picker_items.clear()
            for item in get_routing_catalog_items():
                if search_term and search_term not in routing_search_text(item):
                    continue
                routing_picker_items.append(item)
                routing_listbox.insert(tk.END, format_routing_item(item))
            routing_picker_status.configure(
                text=f"找到 {len(routing_picker_items)} 項；可多選後加入目標院所",
                text_color="#A0AEC0",
            )

        def append_precise_rules(target_entry, target_name):
            selected_indices = routing_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("精準分派規則", "請先從資料庫清單選取至少一個項目。", parent=self.settings_window)
                return

            existing_rules = [
                value.strip()
                for value in re.split(r"[,，]", target_entry.get())
                if value.strip()
            ]
            existing_keys = {value.casefold() for value in existing_rules}
            added_rules = []
            for index in selected_indices:
                if not 0 <= index < len(routing_picker_items):
                    continue
                rule = make_precise_routing_rule(routing_picker_items[index])
                if rule.endswith(":") or rule.casefold() in existing_keys:
                    continue
                existing_rules.append(rule)
                existing_keys.add(rule.casefold())
                added_rules.append(rule)

            if not added_rules:
                routing_picker_status.configure(text="選取項目已存在於規則中", text_color="#F6AD55")
                return

            target_entry.delete(0, tk.END)
            target_entry.insert(0, ", ".join(existing_rules))
            routing_listbox.selection_clear(0, tk.END)
            refresh_routing_preview_status()
            routing_picker_status.configure(
                text=f"✅ 已加入 {len(added_rules)} 項至{target_name}（精準比對）",
                text_color="#68D391",
            )

        picker_buttons = ctk.CTkFrame(routing_picker, fg_color="transparent")
        picker_buttons.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 4))
        ctk.CTkButton(
            picker_buttons,
            text="➕ 加入杏聯",
            height=28,
            fg_color="#2F855A",
            hover_color="#276749",
            command=lambda: append_precise_rules(force_xing_entry, "杏聯"),
        ).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ctk.CTkButton(
            picker_buttons,
            text="➕ 加入博仁",
            height=28,
            fg_color="#2B6CB0",
            hover_color="#2C5282",
            command=lambda: append_precise_rules(force_bo_entry, "博仁"),
        ).pack(side="left", fill="x", expand=True, padx=(4, 0))
        routing_picker_status.grid(row=5, column=0, sticky="ew", padx=10, pady=(0, 7))
        routing_search_var.trace_add("write", refresh_routing_picker)
        refresh_routing_picker()

        def on_inst_select(choice):
            if choice == "➕ 新增/新建院所代號":
                code_entry.delete(0, "end")
                name_entry.delete(0, "end")
                win_remark_box.delete("1.0", "end")
                force_xing_entry.delete(0, "end")
                force_bo_entry.delete(0, "end")
                reset_form()
                refresh_manager_list()
                refresh_routing_preview_status()
            else:
                code = choice.split(" - ")[0].strip()
                code_entry.delete(0, "end")
                code_entry.insert(0, code)
                
                data = self.inst_db.get(code, {})
                name_entry.delete(0, "end")
                if data.get("name"): name_entry.insert(0, data.get("name"))
                    
                win_remark_box.delete("1.0", "end")
                if data.get("remark"): win_remark_box.insert("end", data.get("remark"))
                force_xing_entry.delete(0, "end")
                if data.get("force_to_xinglian"):
                    force_xing_entry.insert(0, data.get("force_to_xinglian"))
                force_bo_entry.delete(0, "end")
                if data.get("force_to_boren"):
                    force_bo_entry.insert(0, data.get("force_to_boren"))
                
                reset_form()
                refresh_manager_list()
                refresh_routing_preview_status()

        inst_selector.configure(command=on_inst_select)

        def save_inst_base_info():
            code = code_entry.get().strip().upper()
            name = name_entry.get().strip()
            remark = win_remark_box.get("1.0", "end").strip()
            force_to_xinglian = force_xing_entry.get().strip()
            force_to_boren = force_bo_entry.get().strip()
            
            if not code:
                messagebox.showwarning("提示", "請輸入院所代號！")
                return

            routing_analysis = get_routing_analysis()
            if routing_analysis["conflicts"]:
                messagebox.showerror(
                    "無法儲存",
                    "同一項目同時被分派至杏聯與博仁，請先修正衝突規則。\n\n"
                    "可按「🔎 預覽規則命中」查看衝突內容。",
                    parent=self.settings_window,
                )
                return
            if routing_analysis["unmatched_rules"]:
                messagebox.showerror(
                    "無法儲存",
                    "有分派規則完全未命中資料庫項目，請先修正規則。\n\n"
                    "可按「🔎 預覽規則命中」查看未命中內容。",
                    parent=self.settings_window,
                )
                return
            if routing_analysis["ambiguous_rules"]:
                messagebox.showerror(
                    "無法儲存",
                    f"有 {len(routing_analysis['ambiguous_rules'])} 條自由文字規則會命中多個項目。\n\n"
                    "請從右側「搜尋並加入分派項目」選取正確項目後，再儲存。",
                    parent=self.settings_window,
                )
                open_routing_preview()
                return
                
            if code not in self.inst_db:
                self.inst_db[code] = {
                    "name": name, "remark": remark, "shortcuts": [],
                    "force_to_xinglian": force_to_xinglian,
                    "force_to_boren": force_to_boren
                }
            else:
                self.inst_db[code]["name"] = name
                self.inst_db[code]["remark"] = remark
                self.inst_db[code]["force_to_xinglian"] = force_to_xinglian
                self.inst_db[code]["force_to_boren"] = force_to_boren
                
            self.save_institution_db_to_disk()
            inst_selector.configure(values=build_combo_options())
            inst_selector.set(f"{code} - {name}" if name else code)
            
            existing_insts = [f"{k} - {v.get('name', '')}".strip(" - ") for k, v in self.inst_db.items()] if self.inst_db else []
            self.inst_combo.configure(values=existing_insts if existing_insts else [""])

            self.search_and_open_panel(auto_triggered=True)
            btn_save_info.configure(text="✓ 已儲存", fg_color="#1b5e20")
            self.settings_window.after(1500, lambda: btn_save_info.configure(text="💾 儲存", fg_color="#2b7b5c"))

        def delete_inst_base_info():
            code = code_entry.get().strip().upper()
            if not code or code not in self.inst_db:
                messagebox.showwarning("提示", "找不到此院所或尚未建立！")
                return
                
            if messagebox.askyesno("確認刪除", f"確定要徹底刪除院所【{code}】的所有資料與快捷鍵嗎？\n（此動作無法復原）"):
                del self.inst_db[code]
                self.save_institution_db_to_disk()
                
                inst_selector.configure(values=build_combo_options())
                inst_selector.set("➕ 新增/新建院所代號")
                on_inst_select("➕ 新增/新建院所代號")
                
                existing_insts = [f"{k} - {v.get('name', '')}".strip(" - ") for k, v in self.inst_db.items()] if self.inst_db else []
                self.inst_combo.configure(values=existing_insts if existing_insts else [""])
                self.inst_combo.set("")
                
                if hasattr(self, 'current_inst_code') and self.current_inst_code == code:
                    self.close_right_panel()
                    
                messagebox.showinfo("刪除成功", f"院所【{code}】已成功刪除！")

        btn_delete_inst.configure(command=delete_inst_base_info)
        btn_save_info.configure(command=save_inst_base_info)

        split_container = ctk.CTkFrame(shortcut_tab, fg_color="transparent")
        split_container.pack(fill="both", expand=True)
        # 左欄：既有快捷鍵＋第 3 步；右欄：較窄但全高的第 1／2 步。
        # 讓資料庫搜尋清單取得完整高度，同時利用左下方原本閒置的空間。
        split_container.grid_columnconfigure(0, weight=6, uniform="shortcut_columns")
        split_container.grid_columnconfigure(1, weight=5, uniform="shortcut_columns")
        split_container.grid_rowconfigure(0, weight=1, uniform="shortcut_rows")
        split_container.grid_rowconfigure(1, weight=1, uniform="shortcut_rows")
        split_container.grid_propagate(False)

        list_frame = ctk.CTkFrame(split_container)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 6))
        list_frame.pack_propagate(False)
        
        ctk.CTkLabel(list_frame, text="📌 該院所現有的快捷按鈕:", font=(UI_FONT, 14, "bold")).pack(anchor="w", padx=10, pady=(8, 2))
        
        scroll_shortcuts = ctk.CTkScrollableFrame(list_frame, fg_color="#1e1e1e")
        scroll_shortcuts.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        add_box = ctk.CTkFrame(split_container, width=650, fg_color="#2d3748")
        add_box.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(8, 0))

        # 第 3 步延展填滿左下方空間，讓已加入項目清單有足夠的工作區。
        step3_box = ctk.CTkFrame(split_container, fg_color="#2d3748")
        step3_box.grid(row=1, column=0, sticky="nsew")
        step3_header = ctk.CTkFrame(step3_box, fg_color="transparent")

        # 右側編輯器使用固定列的 grid 配置，確保第 2 步搜尋清單不會被第 3 步擠掉。
        editor_layout = ctk.CTkFrame(add_box, fg_color="transparent")
        editor_layout.pack(fill="both", expand=True, padx=10, pady=8)
        editor_layout.grid_columnconfigure(0, weight=1)
        editor_layout.grid_rowconfigure(2, weight=1)
        editor_layout.grid_propagate(False)
        
        form_title = ctk.CTkLabel(editor_layout, text="➕ 新增快捷按鈕", font=(UI_FONT, 14, "bold"), text_color="#61afef")
        form_title.grid(row=0, column=0, sticky="w", pady=(0, 3))
        
        f1 = ctk.CTkFrame(editor_layout, fg_color="transparent")
        f1.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        ctk.CTkLabel(f1, text="1. 按鈕顯示名稱:").pack(side="left", padx=5)
        btn_name_entry = ctk.CTkEntry(f1, width=220, placeholder_text="例：➕ CEA 腫瘤套組")
        btn_name_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        submit_frame = ctk.CTkFrame(step3_header, fg_color="transparent")
        action_f = ctk.CTkFrame(step3_header, fg_color="transparent")
        # 捲動內容使用左下方剩餘空間，不讓項目數量改變工具列的高度。
        temp_items_viewport = ctk.CTkFrame(step3_box, height=55, fg_color="#1a202c", corner_radius=6)
        temp_items_viewport.pack_propagate(False)
        temp_items_scroll = ctk.CTkScrollableFrame(
            temp_items_viewport,
            height=35,
            fg_color="#1a202c",
            orientation="vertical",
        )
        lbl_temp_title = ctk.CTkLabel(step3_header, text="3. 已加入項目（✖ 移除）", font=(UI_FONT, 12))
        
        btn_submit = ctk.CTkButton(submit_frame, text="✓ 建立快捷按鈕", height=30, font=(UI_FONT, 12, "bold"), fg_color="#2b7b5c", hover_color="#1e5c45")
        btn_submit.pack(side="left", fill="x", expand=True)
        btn_cancel_edit = ctk.CTkButton(submit_frame, text="取消", width=55, height=30, fg_color="#555555")
        shortcut_feedback = ctk.CTkLabel(
            step3_header,
            text="",
            font=(UI_FONT, 11),
            text_color="#68D391",
            anchor="w",
        )

        def get_shortcut_item_display(item):
            """回傳快捷按鈕管理畫面專用的精簡英文簡碼。

            儲存時仍保留完整項目標籤，避免影響既有快捷按鈕、搜尋與匯出流程。
            """
            item_text = str(item).strip()
            db_item = find_db_item_robust(item_text)
            if not db_item:
                return item_text

            return (
                str(db_item.get("sys_code", "")).strip()
                or str(db_item.get("internal_code", "")).strip()
                or item_text
            )

        def render_temp_items():
            for w in temp_items_scroll.winfo_children(): w.destroy()
                
            if not selected_temp_items:
                ctk.CTkLabel(temp_items_scroll, text="（尚未加入任何項目，請從上方清單選取後加入）", text_color="#a0aec0", font=(UI_FONT, 12)).pack(pady=8)
                return
                
            for idx, item in enumerate(selected_temp_items):
                chip = ctk.CTkFrame(temp_items_scroll, fg_color="#3182ce", corner_radius=6)
                chip.pack(fill="x", padx=4, pady=2)
                
                def remove_single(i=idx):
                    selected_temp_items.pop(i)
                    render_temp_items()
                btn_del = ctk.CTkButton(chip, text="✖", width=20, height=20, fg_color="transparent", hover_color="#e53e3e", text_color="white", command=remove_single)
                btn_del.pack(side="right", padx=(0, 4), pady=2)
                ctk.CTkLabel(
                    chip,
                    text=get_shortcut_item_display(item),
                    anchor="w",
                    font=(UI_FONT, 12, "bold"),
                    text_color="white",
                ).pack(side="left", fill="x", expand=True, padx=(8, 4), pady=2)

        def clear_temp_items():
            selected_temp_items.clear()
            render_temp_items()

        ctk.CTkButton(action_f, text="＋ 加入清單", width=105, height=30, fg_color="#3182ce", hover_color="#2b6cb0", command=lambda: add_item_to_temp()).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ctk.CTkButton(action_f, text="清空", width=52, height=30, fg_color="#555555", command=clear_temp_items).pack(side="left")

        f2 = ctk.CTkFrame(editor_layout, fg_color="transparent")
        f2.grid(row=2, column=0, sticky="nsew", pady=(0, 4))
        
        ctk.CTkLabel(
            f2,
            text="2. 搜尋並選取項目（多選／雙擊加入）",
            font=("Microsoft JhengHei UI", 12),
        ).pack(anchor="w", padx=5, pady=(0, 2))
        
        search_bar = ctk.CTkFrame(f2, fg_color="transparent")
        search_bar.pack(fill="x", padx=5)
        sc_search_var = tk.StringVar()
        entry_sc_search = ctk.CTkEntry(search_bar, textvariable=sc_search_var, placeholder_text="🔍 輸入代碼/名稱快速搜尋...")
        entry_sc_search.pack(side="left", fill="x", expand=True, pady=2)
        
        list_box_f = ctk.CTkFrame(f2, fg_color="transparent")
        list_box_f.pack(fill="both", expand=True, padx=5, pady=2)
        
        scroll_sc = ctk.CTkScrollbar(list_box_f)
        scroll_sc.pack(side="right", fill="y")
        
        listbox_sc = tk.Listbox(list_box_f, font=(UI_FONT, 12), selectmode="extended",
                                yscrollcommand=scroll_sc.set, bg="#1a202c", fg="white",
                                selectbackground="#3182ce", relief="flat", highlightthickness=1, 
                                highlightcolor="#4a5568", exportselection=False)
        listbox_sc.pack(side="left", fill="both", expand=True)
        scroll_sc.configure(command=listbox_sc.yview)
        fit_listbox_rows(listbox_sc, scroll_sc)

        step3_header.pack(fill="x", padx=12, pady=(6, 1))
        step3_header.grid_columnconfigure(1, weight=1)
        lbl_temp_title.grid(row=0, column=0, sticky="w")
        shortcut_feedback.grid(row=0, column=1, sticky="ew", padx=(8, 4))
        action_f.grid(row=1, column=0, sticky="w", pady=(0, 4))
        submit_frame.grid(row=1, column=1, sticky="e", pady=(0, 4))

        temp_items_viewport.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        temp_items_scroll.pack(fill="both", expand=True, padx=1, pady=1)

        def _sc_mousewheel(event):
            listbox_sc.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"
        listbox_sc.bind("<MouseWheel>", _sc_mousewheel)

        def filter_sc_items(*args):
            search_term = sc_search_var.get().strip().lower()
            listbox_sc.delete(0, tk.END)
            for item in all_item_names:
                if search_term in item.lower(): listbox_sc.insert(tk.END, item)

        sc_search_var.trace_add("write", filter_sc_items)
        filter_sc_items()

        def add_item_to_temp():
            selections = listbox_sc.curselection()
            if not selections: return
            for idx in selections:
                it = listbox_sc.get(idx)
                if it and it not in selected_temp_items: selected_temp_items.append(it)
            render_temp_items()
            listbox_sc.selection_clear(0, tk.END)

        listbox_sc.bind("<Double-Button-1>", lambda e: add_item_to_temp())

        def reset_form():
            editing_idx["index"] = None
            form_title.configure(text="➕ 新增快捷按鈕", text_color="#61afef")
            btn_submit.configure(text="✓ 建立快捷按鈕", fg_color="#2b7b5c", hover_color="#1e5c45")
            btn_cancel_edit.pack_forget()
            btn_name_entry.delete(0, "end")
            clear_temp_items()
            shortcut_feedback.configure(text="")

        # 快捷按鈕預設收合，只展開使用者目前需要檢視的套組。
        expanded_shortcut_keys = set()

        def refresh_manager_list():
            for w in scroll_shortcuts.winfo_children():
                w.destroy()

            code = code_entry.get().strip().upper()
            shortcuts = self.inst_db.get(code, {}).get("shortcuts", [])

            if not shortcuts:
                ctk.CTkLabel(scroll_shortcuts, text="尚無設定快捷按鈕\n\n(在右側新增後，即可在此點擊右鍵管理)", text_color="gray").pack(pady=20)
                return

            def shortcut_state_key(index, shortcut_obj):
                items_text = "|".join(str(item) for item in shortcut_obj.get("items", []))
                return f"{index}:{shortcut_obj.get('label', '')}:{items_text}"

            def toggle_shortcut(state_key):
                if state_key in expanded_shortcut_keys:
                    expanded_shortcut_keys.remove(state_key)
                else:
                    expanded_shortcut_keys.add(state_key)
                refresh_manager_list()

            for idx, sc in enumerate(shortcuts):
                shortcut_item_codes = [get_shortcut_item_display(item) for item in sc.get("items", [])]
                state_key = shortcut_state_key(idx, sc)
                is_expanded = state_key in expanded_shortcut_keys

                def start_edit_sc(i=idx, shortcut_obj=sc):
                    editing_idx["index"] = i
                    form_title.configure(text=f"✏️ 編輯快捷按鈕 #{i+1}", text_color="#f6ad55")

                    btn_name_entry.delete(0, "end")
                    btn_name_entry.insert(0, shortcut_obj.get("label", ""))

                    selected_temp_items.clear()
                    selected_temp_items.extend(shortcut_obj.get("items", []))
                    render_temp_items()

                    btn_submit.configure(text="💾 儲存修改", fg_color="#dd6b20", hover_color="#c05621")
                    btn_cancel_edit.pack(side="right", padx=10)

                def delete_sc(i=idx):
                    sc_label = self.inst_db[code]["shortcuts"][i].get("label")
                    if messagebox.askyesno("刪除確認", f"確定要刪除快捷按鈕【{sc_label}】嗎？", parent=self.settings_window):
                        self.inst_db[code]["shortcuts"].pop(i)
                        self.save_institution_db_to_disk()
                        refresh_manager_list()
                        self.search_and_open_panel(auto_triggered=True)
                        if editing_idx["index"] == i:
                            reset_form()

                def show_context_menu(
                    event,
                    edit_callback=start_edit_sc,
                    delete_callback=delete_sc,
                ):
                    menu = tk.Menu(self.settings_window, tearoff=0, font=(UI_FONT, 11))
                    menu.add_command(label="📝 編輯 / 修改套組", command=edit_callback)
                    menu.add_separator()
                    menu.add_command(label="🗑️ 刪除此套組", command=delete_callback)
                    menu.tk_popup(event.x_root, event.y_root)

                header = ctk.CTkButton(
                    scroll_shortcuts,
                    text=("▼" if is_expanded else "▶")
                    + f"  【{sc.get('label', '')}】（{len(shortcut_item_codes)} 個項目）",
                    anchor="w",
                    height=32,
                    font=(UI_FONT, 12, "bold"),
                    fg_color="#2B3D52" if is_expanded else "#2B2B2B",
                    hover_color="#3D5872",
                    command=lambda key=state_key: toggle_shortcut(key),
                )
                header.pack(fill="x", padx=5, pady=(5, 1))
                header.bind("<Button-3>", show_context_menu)
                if sys.platform == "darwin":
                    header.bind("<Button-2>", show_context_menu)

                if is_expanded:
                    details = ctk.CTkFrame(scroll_shortcuts, fg_color="#252525", corner_radius=5)
                    details.pack(fill="x", padx=12, pady=(0, 3))
                    ctk.CTkLabel(
                        details,
                        text="項目：" + ", ".join(shortcut_item_codes),
                        anchor="w",
                        justify="left",
                        wraplength=450,
                        font=(UI_FONT, 12),
                    ).pack(fill="x", padx=10, pady=(7, 3))
                    ctk.CTkLabel(
                        details,
                        text="右鍵點擊上方標題可編輯或刪除",
                        anchor="w",
                        font=(UI_FONT, 10),
                        text_color="#A0AEC0",
                    ).pack(fill="x", padx=10, pady=(0, 7))

                if getattr(self, "target_edit_idx", None) == idx:
                    start_edit_sc(idx, sc)
                    self.target_edit_idx = None

        def save_or_create_shortcut():
            code = code_entry.get().strip().upper()
            name = name_entry.get().strip()
            btn_label = btn_name_entry.get().strip()
            
            if not code or not btn_label:
                messagebox.showwarning("提示", "請輸入院所代號與按鈕顯示名稱！")
                return
            if not selected_temp_items:
                messagebox.showwarning("提示", "請先加入至少一個快捷項目！", parent=self.settings_window)
                return
                
            if code not in self.inst_db:
                self.inst_db[code] = {
                    "name": name,
                    "remark": win_remark_box.get("1.0", "end").strip(),
                    "shortcuts": [],
                }
            else:
                self.inst_db[code]["name"] = name
                self.inst_db[code]["remark"] = win_remark_box.get("1.0", "end").strip()

            shortcuts = self.inst_db[code].get("shortcuts")
            if not isinstance(shortcuts, list):
                shortcuts = []
                self.inst_db[code]["shortcuts"] = shortcuts
                
            shortcut_data = {"label": btn_label, "items": list(selected_temp_items)}
            editing_index = editing_idx["index"]
            if editing_index is None:
                shortcuts.append(shortcut_data)
                action_text = f"已新增第 {len(shortcuts)} 個快捷按鈕：{btn_label}"
            else:
                if 0 <= editing_index < len(shortcuts):
                    shortcuts[editing_index] = shortcut_data
                    action_text = f"已更新快捷按鈕：{btn_label}"
                else:
                    # 設定頁重繪或刪除後索引可能失效；此時安全地視為新增，
                    # 避免第二個快捷按鈕覆蓋第一個。
                    shortcuts.append(shortcut_data)
                    action_text = f"已新增第 {len(shortcuts)} 個快捷按鈕：{btn_label}"
            
            self.save_institution_db_to_disk()
            
            inst_selector.configure(values=build_combo_options())
            inst_selector.set(f"{code} - {name}" if name else code)
            
            existing_insts = [f"{k} - {v.get('name', '')}".strip(" - ") for k, v in self.inst_db.items()] if self.inst_db else []
            self.inst_combo.configure(values=existing_insts if existing_insts else [""])

            reset_form()
            refresh_manager_list()
            selected_main_code = self.inst_var.get().split(" - ")[0].strip().upper()
            if selected_main_code == code:
                self.current_inst_code = code
                self.search_and_open_panel(auto_triggered=True)
            shortcut_feedback.configure(text=f"✅ {action_text}；可直接繼續新增下一個。")

        btn_submit.configure(command=save_or_create_shortcut)
        btn_cancel_edit.configure(command=reset_form)

        if target_code in self.inst_db:
            matched_opt = f"{target_code} - {self.inst_db[target_code].get('name', '')}".strip(" - ")
            inst_selector.set(matched_opt)
            on_inst_select(matched_opt)
        else:
            code_entry.insert(0, target_code)
            render_temp_items()
            refresh_manager_list()

    def build_announcement_ui(self, parent_frame):
        frame = ctk.CTkFrame(parent_frame, corner_radius=10) 
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(frame, text="📢 系統公告設定", font=(UI_FONT, 18, "bold")).pack(pady=(10, 10))

        self.var_show_announcement = ctk.BooleanVar(value=self.app_settings.get("show_announcement", False))
        chk = ctk.CTkCheckBox(
            frame, text="顯示系統公告 (於主畫面上方跑馬燈)", variable=self.var_show_announcement, 
            font=(UI_FONT, 14, "bold")
        )
        chk.pack(anchor="w", padx=20, pady=10)

        input_frame = ctk.CTkFrame(frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        ctk.CTkLabel(input_frame, text="輸入公告文字:", font=(UI_FONT, 14)).pack(side="left")
        
        self.announce_var = tk.StringVar(value=self.app_settings.get("announcement_text", ""))
        self.announce_var.trace_add("write", self._trace_announce_text)
        
        self.lbl_char_count = ctk.CTkLabel(input_frame, text="0/40", font=(UI_FONT, 12), text_color="#AAAAAA")
        self.lbl_char_count.pack(side="right")

        self.entry_announcement = ctk.CTkEntry(
            frame, textvariable=self.announce_var, font=(UI_FONT, 14), width=400
        )
        self.entry_announcement.pack(padx=20, pady=5)
        
        self._trace_announce_text() 

        btn_save = ctk.CTkButton(
            frame, text="💾 儲存設定", font=(UI_FONT, 16, "bold"), 
            fg_color="#28A745", hover_color="#218838",
            command=self.save_system_settings
        )
        btn_save.pack(pady=(20, 10))

    def build_special_formula_ui(self, parent_frame):
        main_container = ctk.CTkFrame(parent_frame, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # 以院所選單切換頁面；不再把 L 欄設定拆成右側另一個畫面。
        # 這樣每家院所的公式與單項 BC 行為會在同一個專屬頁面中維護。
        ctk.CTkLabel(main_container, text="🧪 公式與進階綁定", font=(UI_FONT, 18, "bold")).pack(
            anchor="w", padx=8, pady=(4, 2)
        )
        ctk.CTkLabel(
            main_container,
            text="多選資料庫項目，共用同一組 K 欄關鍵字；院所規則優先於通用規則。",
            font=(UI_FONT, 13), text_color="#AAAAAA", justify="left",
        ).pack(anchor="w", padx=8, pady=(0, 10))

        selector_frame = ctk.CTkFrame(main_container, corner_radius=8, fg_color="#252525")
        selector_frame.pack(fill="x", padx=8, pady=(0, 10))
        ctk.CTkLabel(
            selector_frame, text="📚 選擇院所頁面：", font=(UI_FONT, 14, "bold"), text_color="#61afef"
        ).pack(side="left", padx=(14, 8), pady=10)

        # 院所選擇與 L 欄開關共用同一條工具列，避免佔用規則操作空間。
        l_toggle_host = ctk.CTkFrame(selector_frame, width=245, height=28, fg_color="transparent")
        l_toggle_host.pack(side="right", padx=(8, 14), pady=6)
        l_toggle_host.pack_propagate(False)

        page_frame = ctk.CTkFrame(main_container, corner_radius=10, fg_color="#2B2B2B")
        page_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        selected_page = {"code": "__ALL__"}
        specific_rules_key = "special_formula_keywords_by_institution"
        selector_values = {}

        def normalise_institution_code(code):
            """院所設定的唯一索引：忽略舊資料的大小寫與前後空白。"""
            return str(code or "").strip().upper()

        def normalise_rules(value):
            if not isinstance(value, dict):
                return {}
            normalised = {}
            for item, keywords in value.items():
                item_text = str(item).strip()
                if isinstance(keywords, (list, tuple)):
                    keyword_text = ",".join(str(word).strip() for word in keywords if str(word).strip())
                else:
                    keyword_text = str(keywords).strip()
                if item_text and keyword_text:
                    normalised[item_text] = keyword_text
            return normalised

        def get_all_specific_rules():
            """讀取院所規則並標準化索引，兼容舊版設定檔。"""
            stored = self.app_settings.get(specific_rules_key, {})
            if not isinstance(stored, dict):
                return {}
            normalised = {}
            for stored_code, stored_rules in stored.items():
                page_code = normalise_institution_code(stored_code)
                if page_code and page_code != "__ALL__":
                    normalised[page_code] = normalise_rules(stored_rules)
            return normalised

        def get_page_rules(code):
            page_code = normalise_institution_code(code)
            if page_code == "__ALL__":
                return normalise_rules(self.app_settings.get("special_formula_keywords", {}))
            return dict(get_all_specific_rules().get(page_code, {}))

        def save_page_rules(code, rules):
            page_code = normalise_institution_code(code)
            cleaned_rules = normalise_rules(rules)
            if page_code == "__ALL__":
                self.app_settings["special_formula_keywords"] = cleaned_rules
                return
            all_specific = get_all_specific_rules()
            all_specific[page_code] = cleaned_rules
            self.app_settings[specific_rules_key] = all_specific

        def institution_pages():
            pages = [("__ALL__", "🌐 通用規則（全部院所）")]
            known_codes = set()
            for stored_code, info in self.inst_db.items():
                code = normalise_institution_code(stored_code)
                if not code:
                    continue
                known_codes.add(code)
                name = str(info.get("name", "") or "").strip()
                pages.append((code, f"🏥 {code} - {name}" if name else f"🏥 {code}"))

            for code in get_all_specific_rules():
                if code not in known_codes:
                    pages.append((code, f"⚠️ 未建檔院所 {code}"))
            return pages

        def get_formula_catalog_items():
            """建立可供多選的資料庫項目清單，並排除重複資料。"""
            catalog = []
            seen = set()
            for display_label in all_item_names:
                db_item = find_db_item_robust(display_label)
                if not db_item:
                    continue
                identity = (
                    str(db_item.get("internal_code", "") or "").strip(),
                    str(db_item.get("sys_code", "") or "").strip(),
                    str(db_item.get("full_name", "") or "").strip(),
                    str(db_item.get("org_display", "") or "").strip(),
                )
                if identity in seen:
                    continue
                seen.add(identity)
                catalog.append(db_item)
            return catalog

        def formula_rule_key(item):
            """優先使用資料庫代碼，讓公式產生器可做精準比對。"""
            return str(
                item.get("internal_code") or item.get("sys_code") or item.get("full_name") or ""
            ).strip()

        def format_formula_item(item):
            internal = str(item.get("internal_code", "") or "").strip()
            english = str(item.get("sys_code", "") or "").strip()
            full_name = str(item.get("full_name", "") or "").strip()
            code_text = internal or english or "未設定代碼"
            english_text = f" | {english}" if english and english != code_text else ""
            org = str(item.get("org_display", "") or "").strip()
            org_text = f"[{org}] " if org else ""
            return f"{org_text}{code_text}{english_text} - {full_name}"

        def label_for_code(code):
            page_code = normalise_institution_code(code)
            return dict(institution_pages()).get(page_code, page_code)

        def rebuild_selector():
            pages = institution_pages()
            selector_values.clear()
            options = []
            for code, label in pages:
                selector_values[label] = code
                options.append(label)
            page_selector.configure(values=options)
            current_label = label_for_code(selected_page["code"])
            if current_label not in selector_values:
                current_label = options[0] if options else ""
            if current_label:
                page_selector.set(current_label)

        def render_page(code):
            code = normalise_institution_code(code)
            selected_page["code"] = code
            for child in page_frame.winfo_children():
                child.destroy()
            for child in l_toggle_host.winfo_children():
                child.destroy()

            page_label = label_for_code(code)
            rebuild_selector()
            ctk.CTkLabel(page_frame, text=page_label, font=(UI_FONT, 16, "bold"), text_color="#61afef").pack(
                anchor="w", padx=18, pady=(14, 3)
            )
            if code == "__ALL__":
                hint = "套用至所有院所；L 欄功能請至個別院所設定。"
            else:
                hint = "只會在目前院所代號產生模板時套用；同一項目會覆蓋通用規則。"
            ctk.CTkLabel(page_frame, text=hint, font=(UI_FONT, 12), text_color="#A0AEC0").pack(
                anchor="w", padx=18, pady=(0, 10)
            )

            rules = get_page_rules(code)

            # 原本右側獨立的 L 欄設定，改為放在目前院所的選擇工具列中。
            if code != "__ALL__":
                l_list = {
                    normalise_institution_code(item)
                    for item in (self.app_settings.get("l_col_institutions", []) or [])
                }
                l_enabled = ctk.BooleanVar(value=(code in l_list))

                def toggle_l_col_inst():
                    updated_list = {
                        normalise_institution_code(item)
                        for item in (self.app_settings.get("l_col_institutions", []) or [])
                    }
                    if l_enabled.get():
                        updated_list.add(code)
                    else:
                        updated_list.discard(code)
                    self.app_settings["l_col_institutions"] = sorted(item for item in updated_list if item)
                    self.save_system_settings()

                ctk.CTkCheckBox(
                    l_toggle_host,
                    text="L 欄（單項 BC）自動搜尋",
                    font=(UI_FONT, 12, "bold"),
                    variable=l_enabled,
                    command=toggle_l_col_inst,
                ).pack()
            else:
                ctk.CTkLabel(
                    l_toggle_host,
                    text="L 欄設定請選擇個別院所",
                    font=(UI_FONT, 11), text_color="#A0AEC0",
                ).pack()

            # 左邊負責搜尋與多選，右邊固定顯示已對應項目；切換頁面後不用再往下拉。
            work_split = ctk.CTkFrame(page_frame, fg_color="transparent")
            work_split.pack(fill="both", expand=True, padx=18, pady=(0, 12))
            # 使用固定的等寬格線，不讓文字長度或院所切換改變兩欄比例。
            work_split.grid_rowconfigure(0, weight=1)
            work_split.grid_columnconfigure(0, weight=1, uniform="special_formula_columns")
            work_split.grid_columnconfigure(1, weight=1, uniform="special_formula_columns")
            work_split.grid_propagate(False)

            picker_frame = ctk.CTkFrame(work_split, corner_radius=8, fg_color="#29223A")
            picker_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
            picker_frame.grid_columnconfigure(0, weight=1)
            picker_frame.grid_rowconfigure(4, weight=1)

            rules_panel = ctk.CTkFrame(work_split, corner_radius=8, fg_color="#202A36")
            rules_panel.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
            rules_panel.grid_columnconfigure(0, weight=1)
            rules_panel.grid_rowconfigure(1, weight=1)
            ctk.CTkLabel(
                picker_frame,
                text="🎯 選取項目並設定 K 欄關鍵字",
                font=(UI_FONT, 13, "bold"), text_color="#D6BCFA",
            ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 2))
            ctk.CTkLabel(
                picker_frame,
                text="例如：選取 CRP、ESR，共用關鍵字「炎」。",
                font=(UI_FONT, 11), text_color="#A0AEC0",
            ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 6))

            keyword_row = ctk.CTkFrame(picker_frame, fg_color="transparent")
            keyword_row.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 5))
            ctk.CTkLabel(keyword_row, text="K 欄關鍵字：", font=(UI_FONT, 13)).pack(side="left", padx=(0, 6))
            entry_kw = ctk.CTkEntry(
                keyword_row, font=(UI_FONT, 13), placeholder_text="例如：炎,傷（逗號分隔）"
            )
            entry_kw.pack(side="left", fill="x", expand=True)

            search_var = tk.StringVar()
            search_entry = ctk.CTkEntry(
                picker_frame,
                textvariable=search_var,
                height=29,
                placeholder_text="🔍 搜尋系統代碼、英文簡碼或項目名稱；點一下選取，再點取消",
            )
            search_entry.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 4))

            candidate_frame = ctk.CTkFrame(picker_frame, fg_color="#1A202C")
            candidate_frame.grid(row=4, column=0, sticky="nsew", padx=12, pady=(0, 5))
            candidate_scroll = ctk.CTkScrollbar(candidate_frame)
            candidate_scroll.pack(side="right", fill="y")
            candidate_listbox = tk.Listbox(
                candidate_frame,
                height=6,
                # 由下方的點擊事件處理每一項的切換，不需要 Ctrl / Shift。
                selectmode="multiple",
                font=(UI_FONT, 11),
                yscrollcommand=candidate_scroll.set,
                bg="#1A202C",
                fg="white",
                selectbackground="#6B46C1",
                relief="flat",
                highlightthickness=0,
                exportselection=False,
            )
            candidate_listbox.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            candidate_scroll.configure(command=candidate_listbox.yview)
            fit_listbox_rows(candidate_listbox, candidate_scroll, padx=5, pady=5)
            candidate_items = []

            def toggle_candidate_selection(event):
                """單擊切換單筆選取，方便逐項慢慢挑選。"""
                index = candidate_listbox.nearest(event.y)
                bounds = candidate_listbox.bbox(index)
                if not bounds or not (bounds[1] <= event.y < bounds[1] + bounds[3]):
                    return "break"

                if index in candidate_listbox.curselection():
                    candidate_listbox.selection_clear(index)
                else:
                    candidate_listbox.selection_set(index)
                    candidate_listbox.activate(index)
                selected_count = len(candidate_listbox.curselection())
                picker_status.configure(
                    text=f"已選取 {selected_count} 項；可繼續點選項目或再點一次取消",
                    text_color="#90CDF4",
                )
                return "break"

            candidate_listbox.bind("<Button-1>", toggle_candidate_selection)

            picker_actions = ctk.CTkFrame(picker_frame, fg_color="transparent")
            picker_actions.grid(row=5, column=0, sticky="ew", padx=12, pady=(0, 6))
            btn_add_selected = ctk.CTkButton(
                picker_actions,
                text="➕ 綁定至 K 欄關鍵字",
                height=29,
                fg_color="#2F855A",
                hover_color="#276749",
            )
            btn_add_selected.pack(fill="x")
            picker_status = ctk.CTkLabel(
                picker_actions,
                text="請搜尋並選取資料庫項目",
                font=(UI_FONT, 11), text_color="#A0AEC0", anchor="w", wraplength=300,
            )
            picker_status.pack(fill="x", pady=(2, 0))

            list_header = ctk.CTkFrame(rules_panel, fg_color="transparent")
            list_header.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 3))
            ctk.CTkLabel(
                list_header, text="📌 已對應項目", font=(UI_FONT, 13, "bold"), text_color="#90CDF4"
            ).pack(side="left", padx=(0, 8))
            list_count = ctk.CTkLabel(list_header, text="", font=(UI_FONT, 12, "bold"), text_color="#A0AEC0")
            list_count.pack(side="left")
            # 不使用 Treeview：它收合節點後仍保留原本的白色視窗高度。
            # 改以原生深色群組面板，讓收合後的內容區確實縮小。
            rules_viewport = ctk.CTkFrame(rules_panel, fg_color="transparent", height=1)
            rules_viewport.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 6))
            rules_viewport.pack_propagate(False)
            list_frame = ctk.CTkScrollableFrame(
                rules_viewport,
                fg_color="#1E1E1E",
                corner_radius=6,
                height=220,
            )
            list_frame.pack(fill="x")
            rule_selection_vars = {}
            collapsed_keyword_groups = set()
            catalog_by_key = {formula_rule_key(item): item for item in get_formula_catalog_items() if formula_rule_key(item)}

            def selected_rule_items():
                return [item for item, var in rule_selection_vars.items() if var.get()]

            visible_rule_height = {"height": 44}

            def fit_rule_height(event=None):
                available = rules_viewport._reverse_widget_scaling(rules_viewport.winfo_height())
                height = max(1, min(visible_rule_height["height"], available - 12))
                if list_frame.cget("height") != height:
                    list_frame.configure(height=height)

            rules_viewport.bind("<Configure>", fit_rule_height)

            def update_list_height(grouped_rules):
                """群組收合後同步縮短右側內容區，避免形成清單內的大塊空白。"""
                visible_rows = len(grouped_rules)
                for keywords, items in grouped_rules.items():
                    if keywords not in collapsed_keyword_groups:
                        visible_rows += len(items)
                visible_rule_height["height"] = max(44, min(390, visible_rows * 29 + 8))
                fit_rule_height()

            def toggle_keyword_group(keywords):
                if keywords in collapsed_keyword_groups:
                    collapsed_keyword_groups.remove(keywords)
                else:
                    collapsed_keyword_groups.add(keywords)
                refresh_list()

            def refresh_list():
                for child in list_frame.winfo_children():
                    child.destroy()
                rule_selection_vars.clear()

                grouped_rules = {}
                for item, keywords in rules.items():
                    keyword_text = str(keywords).strip()
                    grouped_rules.setdefault(keyword_text, []).append(item)

                for keywords, items in grouped_rules.items():
                    is_collapsed = keywords in collapsed_keyword_groups
                    keyword_group_button = ctk.CTkButton(
                        list_frame,
                        text=("▶" if is_collapsed else "▼") + f"  K 欄關鍵字：{keywords}（{len(items)} 項）",
                        anchor="w",
                        height=29,
                        font=(UI_FONT, 12, "bold"),
                        fg_color="#2E3B4E",
                        hover_color="#3D4F67",
                        text_color="#F6E05E",
                        command=lambda group_key=keywords: toggle_keyword_group(group_key),
                    )
                    keyword_group_button.pack(fill="x", padx=4, pady=(4, 1))
                    keyword_group_button.bind(
                        "<Button-3>",
                        lambda event, group_key=keywords: show_keyword_group_context_menu(event, group_key),
                    )

                    if is_collapsed:
                        continue

                    group_body = ctk.CTkFrame(list_frame, fg_color="#252525", corner_radius=4)
                    group_body.pack(fill="x", padx=10, pady=(0, 2))
                    for item in items:
                        db_item = catalog_by_key.get(item)
                        item_text = format_formula_item(db_item) if db_item else f"原始對應：{item}"
                        item_var = ctk.BooleanVar(value=False)
                        rule_selection_vars[item] = item_var
                        rule_checkbox = ctk.CTkCheckBox(
                            group_body,
                            text=item_text,
                            variable=item_var,
                            height=25,
                            font=(UI_FONT, 11),
                            command=lambda selected_item=item: load_selected_rule(selected_item),
                        )
                        rule_checkbox.pack(anchor="w", fill="x", padx=8, pady=1)

                list_count.configure(
                    text=f"{len(rules)} 項／{len(grouped_rules)} 組（點擊收合）"
                )
                update_list_height(grouped_rules)

            def load_selected_rule(item=None):
                selected_items = selected_rule_items()
                if item is None and not selected_items:
                    return
                item = item or selected_items[0]
                entry_kw.delete(0, tk.END)
                entry_kw.insert(0, rules[item])
                picker_status.configure(
                    text=f"已載入【{item}】；可修改 K 欄關鍵字後按「更新選取規則」",
                    text_color="#90CDF4",
                )

            def save_rules():
                save_page_rules(code, rules)
                self.save_system_settings()
                # 寫入後回讀目前頁面的規則，讓記憶體、設定檔和右側清單
                # 永遠使用同一份已標準化的資料。
                rules.clear()
                rules.update(get_page_rules(code))

            def refresh_rule_panel():
                """立即重繪右側清單，不必切換院所頁面才看得到異動。"""
                refresh_list()
                # CTkScrollableFrame 的捲動範圍在 idle 才計算；主動刷新
                # 可避免新增／刪除後畫面暫時仍保留舊內容。
                list_frame.update_idletasks()
                rules_panel.update_idletasks()
                page_frame.update_idletasks()

            def refresh_candidates(*_args):
                search_term = search_var.get().strip().casefold()
                candidate_listbox.delete(0, tk.END)
                candidate_items.clear()
                for db_item in get_formula_catalog_items():
                    searchable = " ".join(
                        str(db_item.get(key, "") or "")
                        for key in ("internal_code", "sys_code", "full_name", "org_display")
                    ).casefold()
                    if search_term and search_term not in searchable:
                        continue
                    if not formula_rule_key(db_item):
                        continue
                    candidate_items.append(db_item)
                    candidate_listbox.insert(tk.END, format_formula_item(db_item))
                picker_status.configure(
                    text=f"找到 {len(candidate_items)} 項；點一下選取，再點一次即可取消",
                    text_color="#A0AEC0",
                )

            def add_selected_items():
                keywords = entry_kw.get().strip()
                selected_indices = candidate_listbox.curselection()
                if not keywords:
                    messagebox.showwarning("公式與進階綁定", "請先輸入 K 欄關鍵字。", parent=self.settings_window)
                    return
                if not selected_indices:
                    messagebox.showwarning("公式與進階綁定", "請先從資料庫清單選取至少一個項目。", parent=self.settings_window)
                    return
                added_count = 0
                for index in selected_indices:
                    if not 0 <= index < len(candidate_items):
                        continue
                    item_key = formula_rule_key(candidate_items[index])
                    if not item_key:
                        continue
                    rules[item_key] = keywords
                    added_count += 1
                if not added_count:
                    return
                save_rules()
                refresh_rule_panel()
                candidate_listbox.selection_clear(0, tk.END)
                picker_status.configure(
                    text=f"✅ 已將 {added_count} 個項目綁定至 K 欄關鍵字「{keywords}」",
                    text_color="#68D391",
                )

            def update_selected_rule():
                selected_items = selected_rule_items()
                keywords = entry_kw.get().strip()
                if not selected_items:
                    messagebox.showwarning("公式與進階綁定", "請先從已設定規則中選取一個項目。", parent=self.settings_window)
                    return
                if not keywords:
                    messagebox.showwarning("公式與進階綁定", "K 欄關鍵字不可為空。", parent=self.settings_window)
                    return
                updated_count = 0
                for item in selected_items:
                    rules[item] = keywords
                    updated_count += 1
                save_rules()
                refresh_rule_panel()
                picker_status.configure(
                    text=f"✅ 已更新 {updated_count} 個項目的 K 欄關鍵字",
                    text_color="#68D391",
                )

            def delete_rule():
                items_to_delete = selected_rule_items()
                if not items_to_delete:
                    return
                if not messagebox.askyesno(
                    "確認刪除",
                    f"確定刪除「{page_label}」的 {len(items_to_delete)} 個對應規則嗎？",
                    parent=self.settings_window,
                ):
                    return
                for item in items_to_delete:
                    rules.pop(item, None)
                save_rules()
                refresh_rule_panel()
                entry_kw.delete(0, tk.END)
                picker_status.configure(text=f"已刪除 {len(items_to_delete)} 個對應規則", text_color="#A0AEC0")

            def get_keyword_group_items(keywords):
                return [item for item, value in rules.items() if str(value).strip() == keywords]

            def edit_keyword_group_from_menu(keywords):
                """展開並選取同一關鍵字的整組項目，交由既有更新按鈕寫回。"""
                group_items = get_keyword_group_items(keywords)
                if not group_items:
                    return
                collapsed_keyword_groups.discard(keywords)
                refresh_list()
                for item, item_var in rule_selection_vars.items():
                    item_var.set(item in group_items)
                entry_kw.delete(0, tk.END)
                entry_kw.insert(0, keywords)
                picker_status.configure(
                    text=f"正在編輯 K 欄關鍵字「{keywords}」的 {len(group_items)} 個項目；修改後按「更新選取規則」",
                    text_color="#90CDF4",
                )
                entry_kw.focus_set()

            def delete_keyword_group_from_menu(keywords):
                group_items = get_keyword_group_items(keywords)
                if not group_items:
                    return
                if not messagebox.askyesno(
                    "確認刪除",
                    f"確定刪除 K 欄關鍵字「{keywords}」及其 {len(group_items)} 個對應項目嗎？",
                    parent=self.settings_window,
                ):
                    return
                for item in group_items:
                    rules.pop(item, None)
                collapsed_keyword_groups.discard(keywords)
                save_rules()
                refresh_rule_panel()
                entry_kw.delete(0, tk.END)
                picker_status.configure(
                    text=f"已刪除 K 欄關鍵字「{keywords}」的 {len(group_items)} 個對應項目",
                    text_color="#A0AEC0",
                )

            def show_keyword_group_context_menu(event, keywords):
                """在 K 欄關鍵字群組標題提供整組編輯／刪除。"""
                group_count = len(get_keyword_group_items(keywords))
                context_menu = tk.Menu(self.settings_window, tearoff=False)
                context_menu.add_command(
                    label=f"✏️ 編輯這組關鍵字（{group_count} 項）",
                    command=lambda: edit_keyword_group_from_menu(keywords),
                )
                context_menu.add_separator()
                context_menu.add_command(
                    label=f"🗑️ 刪除這組關鍵字（{group_count} 項）",
                    command=lambda: delete_keyword_group_from_menu(keywords),
                )
                try:
                    context_menu.tk_popup(event.x_root, event.y_root)
                finally:
                    context_menu.grab_release()
                return "break"

            btn_add_selected.configure(command=add_selected_items)
            search_var.trace_add("write", refresh_candidates)
            refresh_candidates()
            refresh_list()

            rule_actions = ctk.CTkFrame(rules_panel, fg_color="transparent")
            rule_actions.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))
            ctk.CTkButton(
                rule_actions, text="💾 更新選取規則", font=(UI_FONT, 12, "bold"),
                fg_color="#2B6CB0", hover_color="#2C5282", command=update_selected_rule,
            ).pack(side="left", padx=(0, 6))
            ctk.CTkButton(
                rule_actions, text="🗑️ 刪除選取規則", font=(UI_FONT, 12, "bold"),
                fg_color="#D9534F", hover_color="#C9302C", command=delete_rule,
            ).pack(side="left")

        def on_page_selected(label):
            code = selector_values.get(label)
            if code is not None:
                render_page(code)

        page_selector = ctk.CTkComboBox(
            selector_frame,
            width=390,
            font=(UI_FONT, 13),
            dropdown_font=(UI_FONT, 13),
            command=on_page_selected,
        )
        page_selector.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=8)
        rebuild_selector()
        render_page("__ALL__")

    def _trace_announce_text(self, *args):
        max_len = 40
        text = self.announce_var.get()
        if len(text) > max_len:
            text = text[:max_len]
            self.announce_var.set(text)
            try: self.entry_announcement.icursor(tk.END)
            except: pass
            
        color = "#FF6B6B" if len(text) >= max_len else "#AAAAAA"
        self.lbl_char_count.configure(text=f"{len(text)}/{max_len}", text_color=color)

    def save_system_settings(self):
        # 整合了公告與通用的 JSON 存檔邏輯
        if hasattr(self, 'var_show_announcement'):
            self.app_settings["show_announcement"] = self.var_show_announcement.get()
        if hasattr(self, 'entry_announcement'):
            self.app_settings["announcement_text"] = self.entry_announcement.get().strip()
        
        try:
            write_json_with_backup(self.app_settings, SYSTEM_CONFIG_FILE)
        except Exception as e: print("設定檔存檔失敗", e)

        self.update_announcement_ui()

    def update_announcement_ui(self):
        if self.marquee_job:
            self.after_cancel(self.marquee_job)
            self.marquee_job = None

        if self.app_settings.get("show_announcement", False):
            raw_text = self.app_settings.get("announcement_text", "今天也是順利排表的一天！")
            self.base_marquee_text = f"{raw_text}   ⭐   "
            self.marquee_offset = 0
            self.lbl_announcement.configure(
                fg_color="#FFF3CD", text_color="#856404", font=(UI_FONT, 14, "bold")
            )
            self.animate_marquee()
        else:
            self.lbl_announcement.configure(text="", fg_color="transparent")

    def animate_marquee(self):
        if not self.app_settings.get("show_announcement", False): return
        
        text = self.base_marquee_text
        self.marquee_offset = (self.marquee_offset + 1) % len(text)
        
        display_str = text[self.marquee_offset:] + text[:self.marquee_offset]
        
        visible_len = 16
        if len(display_str) < visible_len: display_str = display_str.ljust(visible_len, " ")
        else: display_str = display_str[:visible_len]
            
        self.lbl_announcement.configure(text=f" 📢 {display_str} ")
        self.marquee_job = self.after(350, self.animate_marquee)

    def batch_import_db_from_excel(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel 或 CSV 檔案", "*.xlsx *.xls *.xlsm *.csv")], 
            title="選擇包含總表的檔案"
        )
        if not filepath: return
        
        try:
            df = None
            if filepath.lower().endswith('.csv'):
                try: df = pd.read_csv(filepath, header=None, encoding='utf-8-sig')
                except Exception: df = pd.read_csv(filepath, header=None, encoding='big5')
            else:
                for engine in ['openpyxl', 'xlrd', None]:
                    try:
                        df = pd.read_excel(filepath, sheet_name="1", header=None, engine=engine)
                        break
                    except Exception: continue
                if df is None: df = pd.read_excel(filepath, header=None)

            if df is None or df.empty:
                messagebox.showerror("錯誤", "無法讀取該檔案，檔案可能為空或格式不符！")
                return

            if os.path.exists(json_file):
                with open(json_file, 'r', encoding='utf-8') as f:
                    self.current_raw_json = json.load(f)
            else: self.current_raw_json = {}
                
            add_count = 0
            update_count = 0
            skip_count = 0

            header_row_idx = None
            col_map = {"zh": 0, "en": 1, "code": 2, "org": 3, "category": 4} 
            
            for idx, row in df.head(10).iterrows():
                row_str = [str(x).lower().strip() for x in row.values]
                if any("中文" in x or "名稱" in x or "項目" in x or "中/英文" in x for x in row_str):
                    header_row_idx = idx
                    for c_idx, val in enumerate(row_str):
                        if "系統" in val or "代碼" in val: col_map["code"] = c_idx
                        elif "英文" in val or "簡寫" in val: col_map["en"] = c_idx
                        elif "單位" in val or "歸屬" in val: col_map["org"] = c_idx
                        elif "類別" in val: col_map["category"] = c_idx
                        elif "中文" in val or "名稱" in val or "中/英文" in val: col_map["zh"] = c_idx
                    break
            
            if header_row_idx is not None:
                df = df.iloc[header_row_idx+1:].reset_index(drop=True)
                
            for idx, row in df.iterrows():
                try:
                    full_name = str(row[col_map["zh"]]).strip() if col_map["zh"] < len(row) and pd.notna(row[col_map["zh"]]) else ""
                    if not full_name or full_name.lower() in ["nan", "none"]: continue
                    
                    en_name = str(row[col_map["en"]]).strip() if col_map["en"] < len(row) and pd.notna(row[col_map["en"]]) else ""
                    if en_name.lower() in ["nan", "none"]: en_name = ""

                    internal_code = str(row[col_map["code"]]).strip() if col_map["code"] < len(row) and pd.notna(row[col_map["code"]]) else ""
                    if internal_code.lower() in ["nan", "none"]: internal_code = ""

                    if internal_code.endswith('.0'): internal_code = internal_code[:-2]
                    if internal_code.isdigit(): internal_code = internal_code.zfill(3)
                    if not internal_code: 
                        internal_code = en_name or full_name
                    
                    org_raw = str(row[col_map["org"]]).strip().upper() if col_map["org"] < len(row) and pd.notna(row[col_map["org"]]) else "C2"
                    if org_raw.lower() in ["nan", "none"]: org_raw = "C2"

                    category_val = str(row[col_map["category"]]).strip() if col_map["category"] < len(row) and pd.notna(row[col_map["category"]]) else ""
                    if category_val.lower() in ["nan", "none"]: category_val = ""

                    new_item = {
                        "中文名稱": full_name, "英文簡寫": en_name, "系統代碼": internal_code, "歸屬單位": org_raw, "檢驗類別": category_val
                    }

                    if internal_code in self.current_raw_json:
                        if self.current_raw_json[internal_code] == new_item: skip_count += 1
                        else:
                            self.current_raw_json[internal_code] = new_item
                            update_count += 1
                    else:
                        self.current_raw_json[internal_code] = new_item
                        add_count += 1
                except IndexError: continue

            write_json_with_backup(self.current_raw_json, HEALTH_CHECK_ITEMS_FILE)

            success, msg = load_json_database_directly(silent=True)
            if hasattr(self, 'lbl_db_status'): self.lbl_db_status.configure(text=msg, text_color="#28A745")
                
            self.refresh_db_listbox()
            self.refresh_listbox(self.listbox_pub, all_item_names, self.pub_sort_var.get())
            self.refresh_listbox(self.listbox_self, all_item_names, self.self_sort_var.get())
            
            if hasattr(self, 'assistant_tool') and getattr(self, 'assistant_tool', None) and self.assistant_tool.winfo_exists():
                self.assistant_tool.load_database()
                self.assistant_tool._trace_search()

            result_msg = (
                f"🎉 智慧匯入與修復完成！\n\n"
                f"（本次匯入已為您自動修正所有中英文代碼對調之欄位問題）\n\n"
                f"🟢 新增：{add_count} 筆\n🟡 更新與修正：{update_count} 筆\n⚪ 略過 (完全重複)：{skip_count} 筆\n\n"
                f"所有系統已連動更新完成。"
            )
            messagebox.showinfo("匯入成功", result_msg)
            
        except Exception as e: messagebox.showerror("匯入失敗", f"處理 Excel 檔案時發生錯誤：\n{e}")

    def export_db_backup(self):
        if not hasattr(self, 'current_raw_json') or not self.current_raw_json:
            messagebox.showwarning("警告", "目前大腦資料庫是空的，無法匯出！")
            return
            
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel 檔案", "*.xlsx"), ("Excel 2003 檔案", "*.xls")], 
            initialfile="團檢項目(杏).xlsx", 
            title="匯出大腦資料庫備份"
        )
        if filepath:
            try:
                data_list = []
                for sys_code, info in self.current_raw_json.items():
                    data_list.append({
                        "中/英文名稱": info.get("中文名稱", ""), "英文簡寫": info.get("英文簡寫", ""), 
                        "系統代碼": info.get("系統代碼", ""), "歸屬單位": info.get("歸屬單位", ""), "檢驗類別": info.get("檢驗類別", "") 
                    })
                df = pd.DataFrame(data_list)
                df.to_excel(filepath, index=False, sheet_name="1") 
                messagebox.showinfo("成功", f"✅ 大腦資料庫備份已成功匯出至：\n{os.path.basename(filepath)}")
            except Exception as e: messagebox.showerror("匯出失敗", f"存檔時發生錯誤：\n{e}")

    def refresh_db_listbox(self, *args):
        self.listbox_db.delete(0, tk.END)
        search_term = self.db_search_var.get().lower()

        if not os.path.exists(json_file): return
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                self.current_raw_json = json.load(f)

            sorted_items = sorted(
                self.current_raw_json.items(),
                key=lambda x: (safe_natural_sort_key(x[1].get('系統代碼', '')), safe_chinese_sort_key(x[1].get('中文名稱', '')))
            )

            for code, info in sorted_items:
                category = info.get('檢驗類別', '')
                display_str = f"[{info.get('歸屬單位')}] {info.get('系統代碼')} - {info.get('中文名稱')} ({info.get('英文簡寫')})"
                if category: display_str += f" | {category}" 
                    
                if search_term in display_str.lower(): self.listbox_db.insert(tk.END, display_str)
        except Exception: pass

    def on_db_listbox_select(self, event):
        selection = self.listbox_db.curselection()
        if not selection: return
        val = self.listbox_db.get(selection[0])
        
        clean_label = val.split(" | ")[0].strip()
        if "] " in clean_label:
            parts = clean_label.split("] ", 1)
            if " - " in parts[1]:
                code = parts[1].split(" - ", 1)[0].strip()
                if hasattr(self, 'current_raw_json') and code in self.current_raw_json:
                    info = self.current_raw_json[code]
                    self.entry_db_code.delete(0, tk.END); self.entry_db_code.insert(0, info.get("系統代碼", ""))
                    self.entry_db_zh.delete(0, tk.END); self.entry_db_zh.insert(0, info.get("中文名稱", ""))
                    self.entry_db_en.delete(0, tk.END); self.entry_db_en.insert(0, info.get("英文簡寫", ""))
                    self.entry_db_category.delete(0, tk.END); self.entry_db_category.insert(0, info.get("檢驗類別", "")) 
                    self.db_org_var.set(info.get("歸屬單位", "C2"))

    def save_to_json_db(self):
        code = self.entry_db_code.get().strip()
        zh = self.entry_db_zh.get().strip()
        en = self.entry_db_en.get().strip()
        category = self.entry_db_category.get().strip()
        org = self.db_org_var.get()

        if not code or not zh:
            messagebox.showwarning("欄位不完整", "系統代碼與名稱為必填！")
            return

        if not hasattr(self, 'current_raw_json'): self.current_raw_json = {}

        self.current_raw_json[code] = {
            "中文名稱": zh, "英文簡寫": en, "系統代碼": code, "歸屬單位": org, "檢驗類別": category
        }

        write_json_with_backup(self.current_raw_json, HEALTH_CHECK_ITEMS_FILE)

        success, msg = load_json_database_directly(silent=True)
        if hasattr(self, 'lbl_db_status'): self.lbl_db_status.configure(text=msg, text_color="#28A745")
        
        self.refresh_listbox(self.listbox_pub, all_item_names, self.pub_sort_var.get())
        self.refresh_listbox(self.listbox_self, all_item_names, self.self_sort_var.get())
        if hasattr(self, 'assistant_tool') and getattr(self, 'assistant_tool', None) and self.assistant_tool.winfo_exists():
            self.assistant_tool.load_database()
            self.assistant_tool._trace_search()

        self.refresh_db_listbox()
        messagebox.showinfo("同步成功", f"🎉 項目 {zh} ({code}) 已成功儲存！\n選單與小熊貓大腦已同步更新完成！", parent=self.settings_window)
        self.entry_db_code.delete(0, tk.END); self.entry_db_zh.delete(0, tk.END); self.entry_db_en.delete(0, tk.END); self.entry_db_category.delete(0, tk.END)

    def delete_from_json_db(self):
        selections = self.listbox_db.curselection()
        codes_to_delete = []
        if not selections:
            code = self.entry_db_code.get().strip()
            if not code or code not in self.current_raw_json:
                messagebox.showwarning("錯誤", "請先從列表中選取要刪除的項目！\n（提示：可按住 Ctrl 或 Shift 進行多選）")
                return
            codes_to_delete.append(code)
        else:
            for idx in selections:
                val = self.listbox_db.get(idx)
                clean_label = val.split(" | ")[0].strip()
                if "] " in clean_label:
                    parts = clean_label.split("] ", 1)
                    if " - " in parts[1]:
                        code = parts[1].split(" - ", 1)[0].strip()
                        codes_to_delete.append(code)
                    
        if not codes_to_delete: return
            
        msg_text = f"確定要永遠刪除選取的 {len(codes_to_delete)} 個項目嗎？" if len(codes_to_delete) > 1 else f"確定要永遠刪除代碼 [{codes_to_delete[0]}] 嗎？"
        
        if messagebox.askyesno("確認刪除", msg_text):
            for code in codes_to_delete:
                if code in self.current_raw_json: del self.current_raw_json[code]
                    
            write_json_with_backup(self.current_raw_json, HEALTH_CHECK_ITEMS_FILE)
                
            success, msg = load_json_database_directly(silent=True)
            if hasattr(self, 'lbl_db_status'): self.lbl_db_status.configure(text=msg, text_color="#28A745")
            
            self.refresh_listbox(self.listbox_pub, all_item_names, self.pub_sort_var.get())
            self.refresh_listbox(self.listbox_self, all_item_names, self.self_sort_var.get())
            if hasattr(self, 'assistant_tool') and getattr(self, 'assistant_tool', None) and self.assistant_tool.winfo_exists():
                self.assistant_tool.load_database()
                self.assistant_tool._trace_search()

            self.refresh_db_listbox()
            self.entry_db_code.delete(0, tk.END); self.entry_db_zh.delete(0, tk.END); self.entry_db_en.delete(0, tk.END); self.entry_db_category.delete(0, tk.END)
            messagebox.showinfo("已刪除", f"共 {len(codes_to_delete)} 個項目已成功刪除並同步完成！", parent=self.settings_window)
