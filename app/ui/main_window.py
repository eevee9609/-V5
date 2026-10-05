"""主視窗組裝：以多個職責明確的 mixin 組成完整應用程式。"""

import json
import os
import tkinter as tk

import customtkinter as ctk

from app.config import (
    APP_ROOT,
    HEALTH_CHECK_ITEMS_FILE,
    ICON_PATH,
    INSTITUTION_DB_FILE,
    SYSTEM_CONFIG_FILE,
    UI_FONT,
)
from app.data import load_health_check_database
from app.services.excel_generator import ExcelGenerationMixin
from app.ui.roster import RosterMixin
from app.ui.settings import SettingsMixin
from app.ui.tab_names import TAB_PUBLIC, TAB_SELF_PAID, TAB_SPLIT
from app.ui.template_tabs import TemplateTabsMixin
from app.ui.window_actions import WindowActionsMixin


CURRENT_DIR = str(APP_ROOT)
json_file = str(HEALTH_CHECK_ITEMS_FILE)
INST_DB_FILE = str(INSTITUTION_DB_FILE)
ICON_PATH = str(ICON_PATH)


def load_json_database_directly(silent=True):
    from tkinter import messagebox

    return load_health_check_database(
        silent,
        on_warning=messagebox.showwarning,
        on_error=messagebox.showerror,
    )


class HealthCheckGuiApp(
    WindowActionsMixin,
    SettingsMixin,
    RosterMixin,
    TemplateTabsMixin,
    ExcelGenerationMixin,
    ctk.CTk,
):

    def __init__(self):
        super().__init__()
        self.title("醫檢所專用 — 團檢模板自動生成器Ver 5.0")
        self.normal_width = 1150
        self.expanded_width = 1520
        # 依螢幕可用高度設定初始視窗，避免底部的熊貓與設定按鈕被螢幕裁切。
        screen_height = self.winfo_screenheight()
        self.window_height = min(820, max(620, screen_height - 80))
        self.geometry(f"{self.normal_width}x{self.window_height}")

        # ⭐ 載入應用程式圖示
        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(ICON_PATH)
            except Exception:
                pass
        
        self.default_font = (UI_FONT, 13)
        self.title_font = (UI_FONT, 14, "bold")
        self.current_alpha = 0.95
        self.attributes("-alpha", self.current_alpha)
        self.opacity_var = tk.DoubleVar(value=self.current_alpha)
        
        self.public_items = []       
        self.self_paid_config = {}   
        
        self.parsed_roster_data = []
        self.roster_checkbox_vars = []

        self.marquee_job = None
        self.base_marquee_text = ""
        self.marquee_offset = 0

        self.app_settings = {
            "show_announcement": False,
            "announcement_text": "今天也是順利排表的一天！",
            "l_col_institutions": [],
            "special_formula_keywords_by_institution": {},
        }
        config_path = str(SYSTEM_CONFIG_FILE)
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    self.app_settings.update(json.load(f))
            except: pass
        
        success, msg = load_json_database_directly(silent=True)
        if not success: msg = "⚠️ " + msg
            
        self.inst_db = self.load_institution_db()
        self.current_inst_code = ""

        self.main_content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_content_frame.pack(side="left", fill="both", expand=True)

        top_frame = ctk.CTkFrame(self.main_content_frame, fg_color="transparent")
        top_frame.pack(fill="x", padx=20, pady=(15, 5))
        
        self.lbl_db_status = ctk.CTkLabel(top_frame, text=msg, font=(UI_FONT, 14, "bold"), text_color="#28A745" if success else "#FFC107")
        self.lbl_db_status.pack(side="left")

        opacity_frame = ctk.CTkFrame(self.main_content_frame, fg_color="transparent")
        opacity_frame.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(opacity_frame, text="透明度", font=self.default_font).pack(side="left")
        self.opacity_slider = ctk.CTkSlider(opacity_frame, from_=0.35, to=1.0, number_of_steps=13, variable=self.opacity_var, command=self.update_window_opacity, width=180)
        self.opacity_slider.pack(side="left", padx=(10, 8))
        self.opacity_label = ctk.CTkLabel(opacity_frame, text=self.format_opacity_label(self.current_alpha), font=self.default_font, text_color="#AAAAAA")
        self.opacity_label.pack(side="left")
        
        self.inst_frame = ctk.CTkFrame(self.main_content_frame, fg_color="#3a1c1c", border_width=1, border_color="#aa3333")
        self.inst_frame.pack(fill="x", padx=20, pady=(0, 5))
        
        ctk.CTkLabel(self.inst_frame, text="代檢院所:", font=(UI_FONT, 15, "bold"), text_color="#ff9999").pack(side="left", padx=(15, 10), pady=12)
        
        self.inst_var = tk.StringVar()
        self.inst_var.trace_add("write", self.auto_check_inst)
        
        existing_insts = [f"{k} - {v.get('name', '')}".strip(" - ") for k, v in self.inst_db.items()] if hasattr(self, 'inst_db') and self.inst_db else []
        
        self.inst_combo = ctk.CTkComboBox(
            self.inst_frame, width=250, variable=self.inst_var, values=existing_insts if existing_insts else [""], 
            font=(UI_FONT, 14), dropdown_font=(UI_FONT, 14)
        )
        self.inst_combo.pack(side="left", padx=5)

        self.enable_ob_var = tk.BooleanVar(master=self, value=False)
        self.checkbox_enable_ob = ctk.CTkCheckBox(
            self.inst_frame, text="啟用OB", variable=self.enable_ob_var,
            font=self.default_font, width=95, checkbox_width=20, checkbox_height=20,
        )
        self.checkbox_enable_ob.pack(side="left", padx=(15, 5))
        
        if not existing_insts: self.inst_combo.set("")
        
        self.inst_combo.bind("<Return>", lambda e: self.search_and_open_panel())
        self.inst_combo.bind("<Double-1>", lambda e: self.open_settings_window("clinic"))
        
        self.inst_status_lbl = ctk.CTkLabel(self.inst_frame, text="", font=(UI_FONT, 13, "bold"))
        self.inst_status_lbl.pack(side="right", padx=15)

        self.update_window_opacity(self.current_alpha)
        
        roster_action_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        roster_action_frame.pack(side="right", padx=(10, 0))
        self.btn_open_roster = ctk.CTkButton(
            roster_action_frame, text="👥 獨立匯入受檢者名單", font=(UI_FONT, 14, "bold"),
            fg_color="#0275D8", hover_color="#025AA5", height=40, command=self.open_roster_window
        )
        self.btn_open_roster.pack(fill="x")
        session_action_frame = ctk.CTkFrame(roster_action_frame, fg_color="transparent")
        session_action_frame.pack(fill="x", pady=(5, 0))
        # 系統公告固定顯示於「初始化」左側，讓主畫面作業時也能看見提醒。
        self.lbl_announcement = ctk.CTkLabel(
            session_action_frame,
            text="",
            width=210,
            height=32,
            anchor="w",
            font=(UI_FONT, 14, "bold"),
            corner_radius=5,
        )
        self.lbl_announcement.pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.btn_reset_session = ctk.CTkButton(
            session_action_frame,
            text="初始化",
            font=(UI_FONT, 13, "bold"),
            fg_color="#C0392B",
            hover_color="#992D22",
            width=78,
            height=32,
            command=self.reset_for_next_event,
        )
        self.btn_reset_session.pack(side="right")
        self.update_announcement_ui()

        self.lbl_roster_count = ctk.CTkLabel(top_frame, text="(目前尚未匯入名單)", font=(UI_FONT, 13, "bold"), text_color="#AAAAAA")
        self.lbl_roster_count.pack(side="right", padx=10)

        self.tabview = ctk.CTkTabview(self.main_content_frame, corner_radius=15)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=5)
        self.tabview._segmented_button.configure(font=(UI_FONT, 15, "bold"))
        
        self.tab_public = self.tabview.add(TAB_PUBLIC)
        self.tab_self_paid = self.tabview.add(TAB_SELF_PAID)
        self.tab_split = self.tabview.add(TAB_SPLIT)
        
        self.build_public_tab(self.tab_public)
        self.build_self_paid_tab(self.tab_self_paid)
        self.build_split_tab(self.tab_split)

        bottom_frame = ctk.CTkFrame(self.main_content_frame, fg_color="transparent")
        bottom_frame.pack(fill="x", padx=20, pady=10)
        
        btn_save = ctk.CTkButton(bottom_frame, text="💾 儲存預設組合", font=self.default_font, fg_color="#6C757D", hover_color="#5A6268", command=self.save_config)
        btn_save.pack(side="left", padx=5)
        
        btn_load = ctk.CTkButton(bottom_frame, text="📂 載入預設組合", font=self.default_font, fg_color="#6C757D", hover_color="#5A6268", command=self.load_config)
        btn_load.pack(side="left", padx=5)

        reserve_frame = ctk.CTkFrame(bottom_frame, fg_color="#2b2b2b", corner_radius=8)
        reserve_frame.pack(side="left", padx=15)
        ctk.CTkLabel(reserve_frame, text="最少產出總行數:", font=self.default_font).pack(side="left", padx=(10, 5), pady=5)
        self.entry_reserve = ctk.CTkEntry(reserve_frame, width=50, height=28, font=self.default_font)
        self.entry_reserve.pack(side="left", padx=(0, 10), pady=5)
        self.entry_reserve.insert(0, "100") 
        
        btn_generate = ctk.CTkButton(
            bottom_frame, text="🚀 導出標準團檢模板 (Excel)", font=(UI_FONT, 16, "bold"),
            fg_color="#28A745", hover_color="#218838", height=45, corner_radius=10, command=self.generate_excel
        )
        btn_generate.pack(side="right", fill="x", expand=True, padx=(20, 5))

        footer_frame = ctk.CTkFrame(self.main_content_frame, fg_color="transparent")
        footer_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 10))
        
        self.btn_summon_panda = ctk.CTkButton(
            footer_frame, text="🐼 召喚智能小熊貓", font=(UI_FONT, 13, "bold"),
            fg_color="#8E44AD", hover_color="#732D91", height=30, command=self.summon_panda
        )
        self.btn_summon_panda.pack(side="left")

        self.lbl_loaded_preset = ctk.CTkLabel(
            footer_frame,
            text="",
            anchor="w",
            font=(UI_FONT, 12),
            text_color="#5BC0DE",
        )
        self.lbl_loaded_preset.pack(side="left", fill="x", expand=True, padx=(12, 0))
        
        btn_settings = ctk.CTkButton(
            footer_frame, text="⚙️ 系統設定", width=80, font=(UI_FONT, 12, "bold"),
            fg_color="transparent", hover_color="#444444", command=lambda: self.open_settings_window("db")
        )
        btn_settings.pack(side="right", padx=10)

        lbl_signature = ctk.CTkLabel(footer_frame, text="Ver 5.0  |    By 布布 ◍´ᯅ `◍", font=(UI_FONT, 12, "italic"), text_color="#777777")
        lbl_signature.pack(side="right")

        self.build_right_panel()
        self.assistant_tool = None
        self.bind("<F3>", self.reset_for_next_event)
