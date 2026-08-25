"""Helper Flet yang dapat dipakai ulang: snackbar, log aktivitas, komponen.

Fungsi pabrik (make_*) mengembalikan closure siap-pakai sehingga kode GUI
pemanggil cukup melakukan `log = make_activity_log(log_list, page)` tanpa
mengubah ribuan titik pemanggilan yang sudah ada.
"""
import datetime

import flet as ft

LOG_STYLES = {
    "INFO":  {"icon": ft.Icons.INFO_OUTLINE,        "color": ft.Colors.BLUE_300},
    "OK":    {"icon": ft.Icons.CHECK_CIRCLE,         "color": ft.Colors.GREEN_400},
    "WARN":  {"icon": ft.Icons.WARNING_AMBER,        "color": ft.Colors.AMBER_400},
    "ERROR": {"icon": ft.Icons.ERROR_OUTLINE,        "color": ft.Colors.RED_400},
    "STEP":  {"icon": ft.Icons.FLAG_CIRCLE_OUTLINED, "color": ft.Colors.CYAN_300},
}


def stat_box(label: str, value: str, good: bool = True) -> ft.Container:
    """Kartu statistik kecil (label + angka) dengan warna baik/peringatan."""
    return ft.Container(
        content=ft.Column(
            [
                ft.Text(label, size=11, color=ft.Colors.GREY_700),
                ft.Text(value, size=20, weight=ft.FontWeight.BOLD,
                        color=ft.Colors.GREEN_700 if good else ft.Colors.AMBER_800),
            ],
            spacing=2,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        expand=True, padding=12, border_radius=8,
        bgcolor=ft.Colors.GREEN_50 if good else ft.Colors.AMBER_50,
        border=ft.Border.all(1, ft.Colors.GREEN_200 if good else ft.Colors.AMBER_200),
    )


def make_snackbar(page: ft.Page):
    """Pabrik penampilkan SnackBar; mengembalikan fungsi show(message, color)."""

    def show_snackbar(message: str, color: str = ft.Colors.BLUE_700):
        page.show_dialog(
            ft.SnackBar(
                content=ft.Text(message, color=ft.Colors.WHITE, weight=ft.FontWeight.W_500),
                bgcolor=color,
                duration=3500,
            )
        )

    return show_snackbar


def make_activity_log(log_list: ft.ListView, page: ft.Page):
    """Pabrik logger panel aktivitas; mengembalikan fungsi log(msg, level)."""

    def log(message: str, level: str = "INFO"):
        style = LOG_STYLES.get(level, LOG_STYLES["INFO"])
        now = datetime.datetime.now().strftime("%H:%M:%S")
        log_list.controls.append(
            ft.Row(
                [
                    ft.Text(now, size=10, color=ft.Colors.BLUE_GREY_300,
                            font_family="Consolas"),
                    ft.Icon(style["icon"], size=13, color=style["color"]),
                    ft.Text(
                        message, size=12, expand=True, selectable=True,
                        color=ft.Colors.AMBER_100 if level == "WARN" else ft.Colors.GREY_100,
                    ),
                ],
                spacing=6,
                vertical_alignment=ft.CrossAxisAlignment.START,
            )
        )
        page.update()

    return log


async def close_window(page: ft.Page):
    """Tutup jendela aplikasi secara anggun; fallback paksa destroy().

    Flet >= 0.80 memakai API window yang awaitable (Window.close/destroy
    didefinisikan sebagai `async def`), sehingga fungsi ini harus
    di-await oleh handler event.
    """
    try:
        await page.window.close()
    except Exception:
        try:
            await page.window.destroy()
        except Exception:
            pass
