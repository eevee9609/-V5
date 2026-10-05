"""設定頁共用的清單配置。"""

import tkinter as tk
from tkinter import font as tkfont


def size_settings_window(window, owner):
    """依所在螢幕工作區及 DPI 限制初始尺寸，保留標題列與工作列空間。"""
    left, top = 0, 0
    right, bottom = window.winfo_screenwidth(), window.winfo_screenheight()
    try:
        import ctypes
        from ctypes import wintypes

        class MonitorInfo(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("rcMonitor", wintypes.RECT),
                ("rcWork", wintypes.RECT),
                ("dwFlags", wintypes.DWORD),
            ]

        user32 = ctypes.windll.user32
        user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
        user32.MonitorFromWindow.restype = wintypes.HANDLE
        user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
        monitor = user32.MonitorFromWindow(owner.winfo_id(), 2)
        info = MonitorInfo()
        info.cbSize = ctypes.sizeof(info)
        if user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            left, top, right, bottom = (
                info.rcWork.left, info.rcWork.top, info.rcWork.right, info.rcWork.bottom
            )
    except (AttributeError, OSError):
        pass

    scale = window._get_window_scaling()
    width = min(owner.expanded_width, int((right - left - 24) / scale))
    height = min(owner.window_height, int((bottom - top - 56) / scale))
    x = left + max(0, int((right - left - width * scale) / 2))
    y = top + 28
    window.geometry(f"{width}x{height}{x:+d}{y:+d}")
    window.minsize(min(owner.normal_width, width), min(620, height))


def fit_listbox_rows(listbox, scrollbar, *, padx=0, pady=0):
    """以實際字型行高配置清單，剩餘不足一列的空間留在底部。"""
    parent = listbox.master
    listbox.pack_forget()
    scrollbar.pack_forget()
    parent.grid_columnconfigure(0, weight=1)
    parent.grid_rowconfigure(0, weight=1)
    viewport = tk.Frame(parent, bg=listbox.cget("bg"), bd=0, highlightthickness=0)
    viewport.grid(row=0, column=0, sticky="nsew", padx=padx, pady=pady)
    scrollbar.grid(row=0, column=1, sticky="ns", pady=pady)

    def resize_rows(event):
        bounds = listbox.bbox(listbox.nearest(0)) if listbox.size() else None
        # 中英文字型替代後的實際行高可能比 Font.metrics 大，以 Tk 畫出的行為準。
        line_height = bounds[3] + 1 if bounds else (
            tkfont.Font(font=listbox.cget("font")).metrics("linespace") + 1
        )
        border = 2 * (int(listbox.cget("borderwidth")) + int(listbox.cget("highlightthickness")))
        rows = max(0, (event.height - border) // max(1, line_height))
        if rows:
            listbox.place(in_=viewport, x=0, y=0, relwidth=1, height=rows * line_height + border)
            # place(in_=...) 只改配置基準，不會改變父子關係或堆疊順序。
            # viewport 比清單晚建立，必須將清單置於其上，否則背景會遮住文字及點擊。
            listbox.lift(viewport)
        else:
            listbox.place_forget()

    viewport.bind("<Configure>", resize_rows, add="+")
    return viewport
