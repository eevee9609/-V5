"""受檢者名單匯入、勾選與預覽視窗。"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

from app.config import ICON_PATH, UI_FONT
from app.utils.dates import convert_to_taiwan_date, normalize_date_text


ICON_PATH = str(ICON_PATH)


class RosterMixin:

    def open_roster_window(self):
        roster_win = ctk.CTkToplevel(self)
        roster_win.title("👥 受檢者名單匯入與預覽")
        roster_win.geometry("900x650")
        roster_win.transient(self) 
        roster_win.grab_set()       
        self._begin_roster_draft(roster_win)
        roster_win.protocol("WM_DELETE_WINDOW", lambda: self.cancel_roster(roster_win))

        # ⭐ 載入應用程式圖示
        if os.path.exists(ICON_PATH):
            try:
                roster_win.iconbitmap(ICON_PATH)
            except Exception:
                pass
        
        action_bar = ctk.CTkFrame(roster_win, fg_color="transparent")
        action_bar.pack(fill="x", padx=20, pady=(20, 10))
        
        btn_import = ctk.CTkButton(
            action_bar, text="📂 瀏覽 Excel 檔案", font=self.title_font, 
            fg_color="#0275D8", hover_color="#025AA5", height=40, 
            command=lambda: self.browse_roster_file(roster_win)
        )
        btn_import.pack(side="left", padx=(0, 15))
        
        roster_win.lbl_status = ctk.CTkLabel(
            action_bar,
            text="💡 支援多選檔案；匯入內容會先暫存在此視窗，確認後才套用。",
            font=self.default_font,
            text_color="#AAAAAA",
        )
        roster_win.lbl_status.pack(side="left", padx=5)
        
        btn_clear_all = ctk.CTkButton(
            action_bar, text="🗑️ 清空所有名單", width=100, font=self.default_font,
            fg_color="#D9534F", hover_color="#C9302C",
            command=lambda: self.clear_all_roster(roster_win),
        )
        btn_clear_all.pack(side="right", padx=10)
        btn_none = ctk.CTkButton(
            action_bar, text="⬛ 全不選", width=80, font=self.default_font,
            fg_color="#6C757D", hover_color="#5A6268",
            command=lambda: self.select_none_roster(roster_win),
        )
        btn_none.pack(side="right", padx=5)
        btn_all = ctk.CTkButton(
            action_bar, text="☑️ 全選", width=80, font=self.default_font,
            fg_color="#6C757D", hover_color="#5A6268",
            command=lambda: self.select_all_roster(roster_win),
        )
        btn_all.pack(side="right", padx=5)
        
        roster_win.roster_scroll_frame = ctk.CTkScrollableFrame(
            roster_win, corner_radius=15, label_text=" 📋 匯入名單預覽區 "
        )
        roster_win.roster_scroll_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        bottom_bar = ctk.CTkFrame(roster_win, fg_color="transparent")
        bottom_bar.pack(fill="x", padx=20, pady=(0, 20))
        
        btn_confirm = ctk.CTkButton(
            bottom_bar, text="✅ 確認並關閉", font=self.title_font, 
            fg_color="#28A745", hover_color="#218838", height=40, 
            command=lambda: self.confirm_roster(roster_win)
        )
        btn_confirm.pack(side="right", padx=(10, 0))
        
        btn_cancel = ctk.CTkButton(
            bottom_bar, text="❌ 取消", font=self.title_font, 
            fg_color="#D9534F", hover_color="#C9302C", height=40, 
            command=lambda: self.cancel_roster(roster_win),
        )
        btn_cancel.pack(side="right")
        
        self.render_roster_preview(roster_win)

    def _begin_roster_draft(self, win):
        """將正式名單複製成視窗專用草稿，避免取消時誤改主畫面資料。"""
        win.roster_draft_rows = []
        saved_rows = getattr(self, "parsed_roster_data", [])
        saved_vars = getattr(self, "roster_checkbox_vars", [])

        for index, patient in enumerate(saved_rows):
            selected = True
            if index < len(saved_vars):
                try:
                    selected = bool(saved_vars[index].get())
                except tk.TclError:
                    selected = True
            win.roster_draft_rows.append(
                self._make_roster_draft_row(win, patient, selected=selected)
            )

    @staticmethod
    def _cell_text(value):
        """把 pandas / Excel 儲存格安全轉為可顯示的文字。"""
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass

        text = str(value).strip()
        return "" if text.lower() in {"nan", "nat"} else text

    @classmethod
    def _normalise_gender(cls, value):
        """男／男性／M／1 統一為 1；女／女性／F／2 統一為 2。"""
        raw_text = cls._cell_text(value)
        lookup_text = raw_text.casefold().replace(" ", "").replace("　", "")
        if lookup_text in {"1", "1.0", "男", "男性", "男生", "m", "male"}:
            return "1"
        if lookup_text in {"2", "2.0", "女", "女性", "女生", "f", "female"}:
            return "2"
        return raw_text

    def _make_roster_draft_row(self, win, patient, selected=True, source_file="", source_row=None, warnings=None):
        patient_data = {
            "name": self._cell_text(patient.get("name", "")),
            "gender": self._normalise_gender(patient.get("gender", "")),
            "birth": self._cell_text(patient.get("birth", "")),
            "id": self._cell_text(patient.get("id", "")),
        }
        return {
            "patient": patient_data,
            "selected_var": tk.BooleanVar(master=win, value=selected),
            "source_file": source_file,
            "source_row": source_row,
            "warnings": list(warnings or []),
        }

    def _set_roster_status(self, win, text, text_color="#AAAAAA"):
        status_label = getattr(win, "lbl_status", None)
        if status_label is not None:
            status_label.configure(text=text, text_color=text_color)

    def clear_all_roster(self, win):
        draft_rows = getattr(win, "roster_draft_rows", [])
        if not draft_rows:
            self._set_roster_status(win, "💡 目前草稿名單已是空的。")
            return

        confirmed = messagebox.askyesno(
            "確認清空",
            "確定要清空此視窗中的所有名單嗎？\n\n"
            "這項變更會先停留在草稿；按下「確認並關閉」後才會套用到主畫面。",
            parent=win,
        )
        if not confirmed:
            return

        draft_rows.clear()
        self.render_roster_preview(win)
        self._set_roster_status(win, "🗑️ 草稿名單已清空；尚未影響主畫面名單。")

    def confirm_roster(self, win):
        draft_rows = getattr(win, "roster_draft_rows", [])
        committed_data = [dict(row["patient"]) for row in draft_rows]
        committed_vars = [
            ctk.BooleanVar(value=bool(row["selected_var"].get()))
            for row in draft_rows
        ]

        # 用 slice 寫回可保留其他元件若有持有這兩個 list 的參考。
        self.parsed_roster_data[:] = committed_data
        self.roster_checkbox_vars[:] = committed_vars

        active_count = sum(1 for var in self.roster_checkbox_vars if var.get())
        if active_count > 0:
            self.lbl_roster_count.configure(text=f"(目前匯入 {active_count} 人)", text_color="#5BC0DE")
        else:
            self.lbl_roster_count.configure(text="(目前尚未匯入名單)", text_color="#AAAAAA")
        self._close_roster_window(win)

    def cancel_roster(self, win):
        """取消或按右上角 X 時，直接捨棄草稿，不碰正式名單。"""
        self._close_roster_window(win)

    @staticmethod
    def _close_roster_window(win):
        try:
            win.grab_release()
        except tk.TclError:
            pass
        win.destroy()

    def render_roster_preview(self, win):
        if not hasattr(win, "roster_scroll_frame"):
            return
        for child in win.roster_scroll_frame.winfo_children():
            child.destroy()
            
        for draft_row in getattr(win, "roster_draft_rows", []):
            patient = draft_row["patient"]
            id_tag = f"  |  [證號] {patient['id']}" if patient["id"] else ""
            display_text = (
                f"👤 姓名: {patient['name']:<8}  |  性別: {patient['gender']:<3}  |  "
                f"🎂 生日: {patient['birth']:<12}{id_tag}"
            )
            warnings = draft_row.get("warnings", [])
            if warnings:
                display_text += f"  ⚠️ {'、'.join(warnings)}"
            checkbox_options = {
                "text": display_text,
                "variable": draft_row["selected_var"],
                "font": (UI_FONT, 14),
            }
            if warnings:
                checkbox_options["text_color"] = "#D97706"
            chk = ctk.CTkCheckBox(win.roster_scroll_frame, **checkbox_options)
            chk.pack(anchor="w", padx=20, pady=5)

    def browse_roster_file(self, win):
        filepaths = filedialog.askopenfilenames(
            filetypes=[("Excel 或 CSV 檔案", "*.xlsx *.xls *.xlsm *.csv")], title="選擇受檢者名單 (可多選)"
        )
        if not filepaths:
            return
        
        try:
            pending_rows = []
            invalid_row_count = 0
            skipped_file_count = 0
            for filepath in filepaths:
                try:
                    raw_df = self._read_roster_source(filepath)
                except Exception as e:
                    messagebox.showerror("錯誤", f"無法讀取檔案 {os.path.basename(filepath)}: {e}")
                    skipped_file_count += 1
                    continue

                if raw_df.empty or raw_df.shape[1] == 0:
                    messagebox.showwarning("提示", f"檔案 {os.path.basename(filepath)} 沒有可讀取的資料，已跳過。", parent=win)
                    skipped_file_count += 1
                    continue

                mapping = self._detect_trusted_roster_header(raw_df)
                if mapping is None:
                    mapping = self._open_roster_column_mapping_dialog(win, raw_df, filepath)
                if mapping is None:
                    skipped_file_count += 1
                    continue

                candidates, invalid_count = self._parse_roster_candidates(raw_df, mapping, filepath)
                pending_rows.extend(candidates)
                invalid_row_count += invalid_count

            unique_rows, duplicate_count = self._exclude_duplicate_roster_rows(
                getattr(win, "roster_draft_rows", []), pending_rows
            )

            if not unique_rows:
                self._set_roster_status(
                    win,
                    "⚠️ 沒有可加入的有效資料；請確認欄位對應與來源檔案內容。",
                    "#D97706",
                )
                return

            # 已在預覽視窗的草稿中操作，按匯入後直接加入，避免每次都要
            # 再確認一次。略過資料的資訊改顯示在下方狀態列即可。
            import_notes = []
            if duplicate_count:
                import_notes.append(f"略過 {duplicate_count} 筆重複資料")
            if invalid_row_count:
                import_notes.append(f"略過 {invalid_row_count} 筆缺少姓名的資料列")
            if skipped_file_count:
                import_notes.append(f"略過 {skipped_file_count} 個檔案")

            for candidate in unique_rows:
                win.roster_draft_rows.append(
                    self._make_roster_draft_row(
                        win,
                        candidate["patient"],
                        selected=not candidate["warnings"],
                        source_file=candidate["source_file"],
                        source_row=candidate["source_row"],
                        warnings=candidate["warnings"],
                    )
                )

            self.render_roster_preview(win)
            total_count = len(win.roster_draft_rows)
            notes_text = f"（{'；'.join(import_notes)}）" if import_notes else ""
            self._set_roster_status(
                win,
                f"✅ 已加入草稿：本次 {len(unique_rows)} 人，草稿總計 {total_count} 人。{notes_text}請確認勾選後按右下角套用。",
                "#28A745",
            )
            
        except Exception as e:
            messagebox.showerror("讀取失敗", "讀取受檢者名單時發生錯誤：\n" + str(e), parent=win)

    def _read_roster_source(self, filepath):
        if filepath.lower().endswith(".csv"):
            try:
                return pd.read_csv(filepath, header=None, dtype=object, keep_default_na=False, encoding="utf-8-sig")
            except UnicodeDecodeError:
                return pd.read_csv(filepath, header=None, dtype=object, keep_default_na=False, encoding="big5")
        return pd.read_excel(filepath, header=None, dtype=object, keep_default_na=False)

    def _header_field(self, value):
        text = self._cell_text(value).replace(" ", "").replace("　", "").replace("_", "").lower()
        if not text:
            return None

        if text in {"姓名", "受檢者姓名", "受檢人姓名", "name", "patientname"}:
            return "name"
        if "姓名" in text and "聯絡" not in text and "連絡" not in text:
            return "name"

        if text in {"性別", "性别", "sex", "gender"} or "性別" in text or "性别" in text:
            return "gender"

        if text in {"生日", "出生日期", "出生年月日", "birthdate", "birthday", "dob"}:
            return "birth"
        if "生日" in text or "出生" in text:
            return "birth"

        if text in {"身分證號", "身分證字號", "身份證號", "身份證字號", "證號", "id", "idnumber"}:
            return "id"
        if "身分" in text or "身份" in text or "證號" in text:
            return "id"
        return None

    def _detect_trusted_roster_header(self, raw_df):
        """僅在姓名加上至少一個佐證欄位同列時，才信任為正式表頭。"""
        max_scan_rows = min(len(raw_df.index), 30)
        for row_index in range(max_scan_rows):
            mapping = {}
            for column_index, value in enumerate(raw_df.iloc[row_index].tolist()):
                field = self._header_field(value)
                if field and field not in mapping:
                    mapping[field] = column_index

            has_name = "name" in mapping
            has_supporting_field = any(field in mapping for field in ("gender", "birth", "id"))
            if has_name and has_supporting_field:
                mapping["data_start_row"] = row_index + 1
                return mapping
        return None

    @staticmethod
    def _excel_column_name(column_index):
        name = ""
        value = column_index + 1
        while value:
            value, remainder = divmod(value - 1, 26)
            name = chr(65 + remainder) + name
        return name

    def _first_nonempty_row(self, raw_df):
        for row_index in range(len(raw_df.index)):
            if any(self._cell_text(value) for value in raw_df.iloc[row_index].tolist()):
                return row_index
        return 0

    def _column_choice_map(self, raw_df):
        choices = {}
        for column_index in range(raw_df.shape[1]):
            examples = []
            for value in raw_df.iloc[:8, column_index].tolist():
                text = self._cell_text(value).replace("\n", " ")
                if text:
                    examples.append(text)
                if len(examples) == 2:
                    break
            sample = "、".join(examples)[:36] if examples else "（空白）"
            label = f"{self._excel_column_name(column_index)} 欄｜範例：{sample}"
            choices[label] = column_index
        return choices

    def _open_roster_column_mapping_dialog(self, win, raw_df, filepath):
        """無可信表頭時要求操作者明確指定欄位，絕不再猜測欄位順序。"""
        mapping_win = ctk.CTkToplevel(win)
        mapping_win.title(f"確認欄位對應：{os.path.basename(filepath)}")
        mapping_win.geometry("780x610")
        mapping_win.transient(win)
        mapping_win.grab_set()

        ctk.CTkLabel(
            mapping_win,
            text="找不到可信的「姓名＋其他欄位」表頭，請確認資料欄位後再匯入。",
            font=self.title_font,
            text_color="#D97706",
        ).pack(anchor="w", padx=20, pady=(20, 8))
        ctk.CTkLabel(
            mapping_win,
            text="姓名為必填；其餘欄位可選擇「不匯入」。資料起始列請填 Excel 中第一筆受檢者資料所在列。",
            font=self.default_font,
            text_color="#AAAAAA",
        ).pack(anchor="w", padx=20, pady=(0, 10))

        preview_lines = []
        for row_index in range(min(len(raw_df.index), 8)):
            values = []
            for column_index, value in enumerate(raw_df.iloc[row_index].tolist()):
                text = self._cell_text(value).replace("\n", " ")[:24]
                if text:
                    values.append(f"{self._excel_column_name(column_index)}={text}")
            preview_lines.append(f"第 {row_index + 1} 列：" + ("  |  ".join(values) or "（空白）"))

        preview = ctk.CTkTextbox(mapping_win, height=165, font=(UI_FONT, 13))
        preview.pack(fill="x", padx=20, pady=(0, 12))
        preview.insert("1.0", "\n".join(preview_lines))
        preview.configure(state="disabled")

        options = self._column_choice_map(raw_df)
        option_labels = list(options.keys())
        no_import_label = "不匯入"
        required_placeholder = "請選擇姓名欄位"

        form = ctk.CTkFrame(mapping_win, fg_color="transparent")
        form.pack(fill="x", padx=20, pady=(0, 10))
        form.grid_columnconfigure(1, weight=1)

        field_specs = [
            ("name", "姓名（必填）", required_placeholder),
            ("gender", "性別", no_import_label),
            ("birth", "出生日期／生日", no_import_label),
            ("id", "身分證號", no_import_label),
        ]
        field_vars = {}
        for row_index, (field, label, default) in enumerate(field_specs):
            ctk.CTkLabel(form, text=label, font=self.default_font).grid(
                row=row_index, column=0, sticky="w", padx=(0, 12), pady=5
            )
            field_var = tk.StringVar(master=mapping_win, value=default)
            field_vars[field] = field_var
            values = [required_placeholder] + option_labels if field == "name" else [no_import_label] + option_labels
            ctk.CTkOptionMenu(form, variable=field_var, values=values, font=self.default_font).grid(
                row=row_index, column=1, sticky="ew", pady=5
            )

        ctk.CTkLabel(form, text="資料起始列", font=self.default_font).grid(
            row=len(field_specs), column=0, sticky="w", padx=(0, 12), pady=(8, 5)
        )
        start_row_var = tk.StringVar(master=mapping_win, value=str(self._first_nonempty_row(raw_df) + 1))
        ctk.CTkEntry(form, textvariable=start_row_var, width=130, font=self.default_font).grid(
            row=len(field_specs), column=1, sticky="w", pady=(8, 5)
        )

        result = {"mapping": None}

        def close_dialog():
            try:
                mapping_win.grab_release()
            except tk.TclError:
                pass
            mapping_win.destroy()

        def confirm_mapping():
            chosen_mapping = {}
            for field, _, _ in field_specs:
                choice = field_vars[field].get()
                if field == "name" and choice == required_placeholder:
                    messagebox.showwarning("需要姓名欄位", "請選擇受檢者的姓名欄位。", parent=mapping_win)
                    return
                chosen_mapping[field] = None if choice == no_import_label else options.get(choice)

            selected_columns = [column for column in chosen_mapping.values() if column is not None]
            if len(selected_columns) != len(set(selected_columns)):
                messagebox.showwarning("欄位重複", "同一欄位不能同時指定為兩種資料。", parent=mapping_win)
                return

            try:
                start_row = int(start_row_var.get().strip()) - 1
            except ValueError:
                messagebox.showwarning("資料起始列錯誤", "請輸入有效的列號，例如 2。", parent=mapping_win)
                return
            if start_row < 0 or start_row >= len(raw_df.index):
                messagebox.showwarning("資料起始列錯誤", "資料起始列必須在來源檔案的範圍內。", parent=mapping_win)
                return

            chosen_mapping["data_start_row"] = start_row
            result["mapping"] = chosen_mapping
            close_dialog()

        button_bar = ctk.CTkFrame(mapping_win, fg_color="transparent")
        button_bar.pack(fill="x", padx=20, pady=(10, 20))
        ctk.CTkButton(
            button_bar, text="略過此檔", font=self.default_font,
            fg_color="#6C757D", hover_color="#5A6268", command=close_dialog,
        ).pack(side="left")
        ctk.CTkButton(
            button_bar, text="確認欄位並繼續", font=self.title_font,
            fg_color="#0275D8", hover_color="#025AA5", command=confirm_mapping,
        ).pack(side="right")

        mapping_win.protocol("WM_DELETE_WINDOW", close_dialog)
        mapping_win.wait_window()
        if win.winfo_exists():
            win.grab_set()
        return result["mapping"]

    def _parse_roster_candidates(self, raw_df, mapping, filepath):
        candidates = []
        invalid_row_count = 0
        start_row = mapping["data_start_row"]

        def get_raw_value(row, field):
            column_index = mapping.get(field)
            if column_index is None:
                return None
            return row.iloc[column_index]

        def get_value(row, field):
            return self._cell_text(get_raw_value(row, field))

        for row_index in range(start_row, len(raw_df.index)):
            row = raw_df.iloc[row_index]
            name = get_value(row, "name")
            row_has_content = any(self._cell_text(value) for value in row.tolist())
            if not name:
                if row_has_content:
                    invalid_row_count += 1
                continue

            # 避免檔案中重複出現的表頭列被當成人員資料。
            if self._header_field(name) == "name":
                continue

            gender = self._normalise_gender(get_raw_value(row, "gender"))
            raw_birth = get_value(row, "birth")
            birth = raw_birth
            warnings = []
            if raw_birth:
                try:
                    normalized_birth = self._cell_text(normalize_date_text(get_raw_value(row, "birth")))
                    if normalized_birth:
                        birth = normalized_birth
                    else:
                        warnings.append("生日格式需確認")
                except Exception:
                    warnings.append("生日格式需確認")

            if gender and gender not in {"1", "2"}:
                warnings.append("性別格式需確認")

            candidates.append({
                "patient": {"name": name, "gender": gender, "birth": birth, "id": get_value(row, "id")},
                "source_file": os.path.basename(filepath),
                "source_row": row_index + 1,
                "warnings": warnings,
            })
        return candidates, invalid_row_count

    @staticmethod
    def _identity_text(value):
        text = str(value or "").strip().upper().replace(" ", "")
        if text.endswith(".0") and text[:-2].isdigit():
            text = text[:-2]
        return text

    def _patient_identity_values(self, patient):
        id_key = self._identity_text(patient.get("id", ""))
        name = self._identity_text(patient.get("name", ""))
        birth = self._identity_text(patient.get("birth", ""))
        name_birth_key = f"{name}|{birth}" if name and birth else ""
        return id_key, name_birth_key

    def _exclude_duplicate_roster_rows(self, draft_rows, pending_rows):
        """證號優先比對；未提供證號時才以姓名＋生日避免重複匯入。"""
        seen_ids = set()
        seen_name_birth = set()
        for draft_row in draft_rows:
            id_key, name_birth_key = self._patient_identity_values(draft_row["patient"])
            if id_key:
                seen_ids.add(id_key)
            if name_birth_key:
                seen_name_birth.add(name_birth_key)

        unique_rows = []
        duplicate_count = 0
        for candidate in pending_rows:
            id_key, name_birth_key = self._patient_identity_values(candidate["patient"])
            is_duplicate = bool(id_key and id_key in seen_ids)
            if not id_key and name_birth_key and name_birth_key in seen_name_birth:
                is_duplicate = True
            if is_duplicate:
                duplicate_count += 1
                continue

            unique_rows.append(candidate)
            if id_key:
                seen_ids.add(id_key)
            if name_birth_key:
                seen_name_birth.add(name_birth_key)
        return unique_rows, duplicate_count

    def select_all_roster(self, win):
        for draft_row in getattr(win, "roster_draft_rows", []):
            draft_row["selected_var"].set(True)

    def select_none_roster(self, win):
        for draft_row in getattr(win, "roster_draft_rows", []):
            draft_row["selected_var"].set(False)
