"""應用程式啟動與全域初始化。"""

from app.config import configure_tk_environment, configure_ui_theme


# 必須先完成 Tcl/Tk 環境設定，之後才匯入建立視窗的模組。
configure_tk_environment()
configure_ui_theme()

from app.ui.main_window import HealthCheckGuiApp


def run() -> None:
    app = HealthCheckGuiApp()
    app.mainloop()
