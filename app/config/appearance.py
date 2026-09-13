"""CustomTkinter 外觀與 Tcl/Tk 相容性設定。"""

from __future__ import annotations

import os
import sys


UI_FONT = "MiSans"


def configure_tk_environment() -> None:
    """在建立 Tk 視窗前，嘗試補上缺失的 Tcl/Tk library 路徑。"""
    base_prefix = getattr(sys, "base_prefix", sys.prefix)
    for candidate in (
        os.path.join(base_prefix, "tcl", "tcl8.6"),
        os.path.join(base_prefix, "tcl8.6"),
        os.path.join(base_prefix, "lib", "tcl8.6"),
        os.path.join(base_prefix, "tcl8.6.14", "library"),
    ):
        if os.path.isdir(candidate):
            os.environ.setdefault("TCL_LIBRARY", candidate)
            break

    for candidate in (
        os.path.join(base_prefix, "tcl", "tk8.6"),
        os.path.join(base_prefix, "tk8.6"),
        os.path.join(base_prefix, "lib", "tk8.6"),
    ):
        if os.path.isdir(candidate):
            os.environ.setdefault("TK_LIBRARY", candidate)
            break


def configure_ui_theme() -> None:
    """只在應用程式啟動時設定一次全域 CustomTkinter 外觀。"""
    import customtkinter as ctk

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
