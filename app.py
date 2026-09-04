"""
Generator Dokumen Administrasi SE2026
Workflow wizard dengan step tracker (ala balenaEtcher) + panel log.

Alur:
  1. Pilih dokumen yang ingin dibuat
  2. Upload template dokumen (atau pakai template bawaan)
  3. Upload data (download template -> upload -> verifikasi)
  4. Generate (log pengisian data per baris)
  5. Simpan hasil (ZIP PDF tiap petugas / PDF gabungan urut no_urut / DOCX / PDF tunggal)
"""
import asyncio
import datetime
import os
import shutil
import tempfile
import time

import flet as ft

from src import lampiran_spk
from src import bapp_pml
from src import bapp_ppl
from src import bapp_pml_t2
from src import bapp_ppl_t2
from src import spp
from src import spp_t2
from src import bast
from src import bukti_terima
from src.validator import analyze_nulls
from src.workflow_routing import validate_document_input
from src.template_generator import generate_template
from src.workflow import (
    GROUP_ORDER,
    documents_by_group,
    get_document_by_id,
)
from src.document_generator import (
    HAS_DOCX,
    PLACEHOLDER_RE,
    build_values,
    copy_template,
    extract_template_placeholders,
    fill_row,
    slugify,
    validate_template_placeholders,
)
from src.release import (
    APP_FULL_NAME,
    APP_TITLE,
    APPLICATION_IDENTIFIER,
    DISPLAY_VERSION,
    PACKAGE_VERSION,
    PUBLISHER,
    RELEASE_CHANNEL,
    RELEASES_URL,
)
from src.updates import check_for_update
from utils import (
    DebounceGate,
    FEEDBACK_URL,
    MERGE_AVAILABLE,
    PDF_AVAILABLE,
    close_window,
    conversion_estimate_messages,
    conversion_workload,
    convert_docx_files_to_pdf,
    duration_info_box,
    dismiss_feedback_prompt,
    ensure_extension,
    estimate_conversion_seconds,
    file_order_key,
    format_duration,
    format_timer_clock,
    make_activity_log,
    make_snackbar,
    merge_pdfs,
    open_external_url,
    open_in_explorer,
    record_launch_and_should_prompt,
    save_dialog_options,
    stat_box,
    zip_files,
)


STEP_DEFS = [
    {"label": "Pilih Dokumen", "icon": ft.Icons.DESCRIPTION_OUTLINED},
    {"label": "Template Dokumen", "icon": ft.Icons.FILE_COPY_OUTLINED},
    {"label": "Upload Data", "icon": ft.Icons.UPLOAD_FILE_OUTLINED},
    {"label": "Generate", "icon": ft.Icons.AUTO_FIX_HIGH_OUTLINED},
    {"label": "Simpan & Selesai", "icon": ft.Icons.TASK_ALT_ROUNDED},
]

DOC_ICONS = {
    "Lampiran SPK": ft.Icons.DESCRIPTION_ROUNDED,
    "BAPP Termin 1": ft.Icons.FACT_CHECK_OUTLINED,
    "SPP Termin 1": ft.Icons.ASSIGNMENT_RETURNED_ROUNDED,
    "BAPP Termin 2": ft.Icons.FACT_CHECK_OUTLINED,
    "SPP Termin 2": ft.Icons.ASSIGNMENT_RETURNED_ROUNDED,
    "BAST": ft.Icons.HANDSHAKE_OUTLINED,
    "Bukti Terima": ft.Icons.RECEIPT_LONG_ROUNDED,
}


def main(page: ft.Page):
    page.title = APP_FULL_NAME
    page.theme_mode = ft.ThemeMode.LIGHT
    page.theme = ft.Theme(
        # Scrollbar tebal & selalu jelas agar area yang bisa digulir terlihat
        scrollbar_theme=ft.ScrollbarTheme(
            thickness=8,
            thumb_color=ft.Colors.BLUE_GREY_300,
            track_color=ft.Colors.GREY_200,
            track_visibility=True,
            main_axis_margin=4,
            cross_axis_margin=0,
            radius=4,
        ),
    )
    page.padding = 0
    page.bgcolor = ft.Colors.GREY_100
    page.window.maximized = True          # jendela maksimized sejak awal
    page.window.min_width = 1100
    page.window.min_height = 700

    # ------------------------------------------------------------------ #
    # State aplikasi                                                      #
    # ------------------------------------------------------------------ #
    state = {
        "step": 0,            # langkah aktif (0..4)
        "max_step": 0,        # langkah tertinggi yang boleh dikunjungi
        "doc_id": None,       # pilihan step 1
        "file_path": None,    # excel data (step 3)
        "dfs": {},
        "data_ok": False,
        "errors": [],
        "template_path": None,       # template docx (step 2)
        "template_source": None,     # "unggahan" (template selalu diunggah user)
        "generated_files": [],       # hasil step 4
        "generation_done": False,
        "gen_duration": None,        # durasi pembuatan dokumen (detik)
        "conv_duration": None,       # durasi konversi PDF/penyimpanan (detik)
        "saved_path": None,          # hasil step 5
        "saved": False,              # True setelah output berhasil disimpan
        "evidence_image_layout": "grid",  # grid | dedicated_pages
        "evidence_image_orientation": "automatic",  # automatic | portrait | landscape
        "busy": False,
    }

    def current_doc():
        return get_document_by_id(state["doc_id"])

    # ------------------------------------------------------------------ #
    # Panel log (kanan)                                                   #
    # ------------------------------------------------------------------ #
    log_list = ft.ListView(expand=True, spacing=2, padding=10, auto_scroll=True)
    show_snackbar = make_snackbar(page)
    log = make_activity_log(log_list, page)

    def clear_logs(e=None):
        log_list.controls.clear()
        log("Log dibersihkan.", "INFO")

    def close_dialog(dialog: ft.AlertDialog):
        dialog.open = False
        page.update()

    def show_about(e=None):
        about_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Tentang SIOMAY"),
            content=ft.Column(
                [
                    ft.Text(APP_FULL_NAME, weight=ft.FontWeight.BOLD,
                            color=ft.Colors.BLUE_900),
                    ft.Text("Sistem otomasi dokumen administrasi SE2026.",
                            size=13, color=ft.Colors.GREY_700),
                    ft.Divider(),
                    ft.Text(f"Versi: {DISPLAY_VERSION} ({PACKAGE_VERSION})", size=13),
                    ft.Text(f"Kanal rilis: {RELEASE_CHANNEL}", size=13),
                    ft.Text(f"Publisher: {PUBLISHER}", size=13),
                    ft.Text(f"ID aplikasi: {APPLICATION_IDENTIFIER}", size=11,
                            color=ft.Colors.GREY_600, selectable=True),
                ],
                tight=True,
                spacing=8,
            ),
            actions=[
                ft.TextButton(
                    "Buka GitHub Releases",
                    on_click=lambda _: (close_dialog(about_dialog), open_external_url(RELEASES_URL, page)),
                ),
                ft.TextButton("Tutup", on_click=lambda _: close_dialog(about_dialog)),
            ],
        )
        page.show_dialog(about_dialog)

    def show_feedback(e=None, resume_update_check: bool = False):
        """Invite the user to send feedback through the official BPS form."""
        def finish(send: bool):
            dismiss_feedback_prompt()
            close_dialog(feedback_dialog)
            if send:
                open_external_url(FEEDBACK_URL, page)
            if resume_update_check:
                page.run_task(on_check_updates)

        feedback_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Bantu tingkatkan SIOMAY"),
            content=ft.Text(
                "Bagaimana pengalaman Anda menggunakan SIOMAY? Masukan Anda membantu "
                "kami meningkatkan aplikasi.",
                size=13,
            ),
            actions=[
                ft.TextButton("Nanti", on_click=lambda _: finish(False)),
                ft.FilledButton("Kirim feedback", on_click=lambda _: finish(True)),
            ],
        )
        page.show_dialog(feedback_dialog)

    async def on_check_updates(e=None):
        """Check GitHub metadata only; installers are never run by this app."""
        show_snackbar("Memeriksa pembaruan…", ft.Colors.BLUE_700)
        try:
            update = await asyncio.to_thread(check_for_update)
        except Exception as ex:
            log(f"Pemeriksaan pembaruan gagal: {ex}", "WARN")
            show_snackbar("Pembaruan tidak dapat diperiksa. Coba lagi nanti.",
                          ft.Colors.AMBER_800)
            return

        if update is None:
            log(f"SIOMAY {DISPLAY_VERSION} sudah versi terbaru di kanal {RELEASE_CHANNEL}.",
                "OK")
            show_snackbar("Anda sudah menggunakan versi terbaru.", ft.Colors.GREEN_700)
            return

        update_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Pembaruan tersedia"),
            content=ft.Column(
                [
                    ft.Text(f"{update.display_version} tersedia untuk kanal {RELEASE_CHANNEL}.",
                            weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                    ft.Text(f"Versi terpasang: {DISPLAY_VERSION}", size=13),
                    ft.Divider(),
                    ft.Text("Yang baru", weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Markdown(update.changelog, selectable=True)
                                if update.changelog
                                else ft.Text(
                                    "Catatan perubahan tidak dapat dimuat. Anda tetap dapat "
                                    "melihatnya di halaman rilis resmi.",
                                    size=13,
                                    italic=True,
                                    color=ft.Colors.GREY_700,
                                )
                            ],
                            scroll=ft.ScrollMode.AUTO,
                        ),
                        width=600,
                        height=260,
                        padding=12,
                        border=ft.Border.all(1, ft.Colors.GREY_300),
                        border_radius=8,
                    ),
                    ft.Text(
                        "Untuk keamanan, unduh pembaruan hanya dari halaman "
                        "GitHub Releases resmi.",
                        size=13, color=ft.Colors.GREY_700,
                    ),
                ],
                tight=True,
                spacing=8,
            ),
            actions=[
                ft.TextButton(
                    "Buka halaman rilis",
                    on_click=lambda _: (close_dialog(update_dialog), open_external_url(update.release_notes_url, page)),
                ),
                ft.TextButton("Nanti", on_click=lambda _: close_dialog(update_dialog)),
            ],
        )
        log(f"Pembaruan tersedia: {update.display_version} ({update.package_version}).", "INFO")
        page.show_dialog(update_dialog)

    # ------------------------------------------------------------------ #
    # Step tracker (atas, ala balenaEtcher)                               #
    # ------------------------------------------------------------------ #
    CIRCLE = 50
    LABEL_WIDTH = CIRCLE + 40

    def build_circle(index: int):
        st = state["step"]
        if index < st:
            bg, border = ft.Colors.GREEN_600, None
            icon, icolor = ft.Icons.CHECK_ROUNDED, ft.Colors.WHITE
        elif index == st:
            bg = ft.Colors.BLUE_700
            border = ft.BorderSide(3, ft.Colors.BLUE_200)
            icon, icolor = STEP_DEFS[index]["icon"], ft.Colors.WHITE
        else:
            bg = ft.Colors.WHITE
            border = ft.BorderSide(2, ft.Colors.GREY_300)
            icon, icolor = STEP_DEFS[index]["icon"], ft.Colors.GREY_400
        clickable = index <= state["max_step"]
        return ft.Container(
            content=ft.Icon(icon, size=22, color=icolor),
            width=CIRCLE, height=CIRCLE, bgcolor=bg, border=border,
            border_radius=CIRCLE // 2,
            alignment=ft.Alignment.CENTER,
            tooltip=f"Langkah {index + 1}: {STEP_DEFS[index]['label']}",
            on_click=(lambda e, i=index: goto(i)) if clickable else None,
        )

    def build_connector(done: bool):
        return ft.Container(
            expand=True, height=4, border_radius=2,
            bgcolor=ft.Colors.GREEN_400 if done else ft.Colors.GREY_200,
        )

    def build_label(index: int):
        st = state["step"]
        if index < st:
            color, weight = ft.Colors.GREEN_700, ft.FontWeight.W_600
        elif index == st:
            color, weight = ft.Colors.BLUE_900, ft.FontWeight.BOLD
        else:
            color, weight = ft.Colors.GREY_500, ft.FontWeight.W_400
        clickable = index <= state["max_step"]
        return ft.Container(
            # Gunakan sel selebar ikon agar posisi tengah setiap label
            # persis sama dengan pusat lingkaran. Teks dibuat lebih lebar
            # dan ditumpangkan simetris supaya label panjang tetap terbaca.
            content=ft.Stack(
                [
                    ft.Container(
                        content=ft.Text(
                            STEP_DEFS[index]["label"], size=12, color=color,
                            weight=weight, text_align=ft.TextAlign.CENTER,
                        ),
                        width=LABEL_WIDTH,
                        alignment=ft.Alignment.CENTER,
                        left=-(LABEL_WIDTH - CIRCLE) // 2,
                    ),
                ],
                clip_behavior=ft.ClipBehavior.NONE,
            ),
            width=CIRCLE,
            height=36,
            on_click=(lambda e, i=index: goto(i)) if clickable else None,
            tooltip=None if clickable else "Selesaikan langkah sebelumnya dahulu",
        )

    tracker_box = ft.Column(spacing=6)

    def render_tracker():
        circles, labels = [], []
        for i in range(len(STEP_DEFS)):
            circles.append(build_circle(i))
            labels.append(build_label(i))
            if i < len(STEP_DEFS) - 1:
                circles.append(build_connector(i < state["step"]))
                labels.append(ft.Container(expand=True))
        tracker_box.controls = [
            ft.Row(circles, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Row(labels, vertical_alignment=ft.CrossAxisAlignment.START),
        ]
        page.update()

    # ------------------------------------------------------------------ #
    # File pickers (Flet >= 0.80: awaitable API via page.services)        #
    # ------------------------------------------------------------------ #
    data_picker = ft.FilePicker()
    tpl_picker = ft.FilePicker()
    save_tpl_picker = ft.FilePicker()
    save_out_picker = ft.FilePicker()
    for p in (data_picker, tpl_picker, save_tpl_picker, save_out_picker):
        page.services.append(p)

    # ------------------------------------------------------------------ #
    # LANGKAH 1 - Pilih dokumen                                           #
    # ------------------------------------------------------------------ #
    doc_cards = {}
    step1_summary = ft.Text(
        "Belum ada dokumen dipilih.", italic=True,
        color=ft.Colors.GREY_600, size=13,
    )

    def restyle_doc_cards():
        for doc_id, card in doc_cards.items():
            selected = doc_id == state["doc_id"]
            card.bgcolor = ft.Colors.BLUE_50 if selected else ft.Colors.WHITE
            card.border = ft.Border.all(
                2 if selected else 1,
                ft.Colors.BLUE_700 if selected else ft.Colors.BLUE_GREY_100,
            )
            card.content.controls[0].icon = (
                ft.Icons.CHECK_CIRCLE_ROUNDED if selected
                else DOC_ICONS.get(card.data["group"], ft.Icons.INSERT_DRIVE_FILE_OUTLINED)
            )
            card.content.controls[0].icon_color = (
                ft.Colors.BLUE_700 if selected else ft.Colors.BLUE_GREY_300
            )

    def select_doc(doc_id: str):
        state["doc_id"] = doc_id
        # Input validity belongs to the previously selected workflow schema.
        # Never carry it into another document type without revalidation.
        state["file_path"] = None
        state["dfs"] = {}
        state["errors"] = []
        state["data_ok"] = False
        state["generated_files"] = []
        state["generation_done"] = False
        state["gen_duration"] = None
        state["conv_duration"] = None
        gen_duration_box.visible = False
        save_duration_box.visible = False
        data_file_chip.value = "Belum ada file dipilih."
        data_file_chip.italic = True
        data_file_chip.color = ft.Colors.GREY_600
        verify_area.controls = []
        verify_area.visible = False
        doc = current_doc()
        builtin = doc.builtin_template_path
        restyle_doc_cards()
        # Dokumen tanpa template: auto-set sentinel agar gate Step 2 terpenuhi
        if doc.no_template:
            state["template_path"] = "__blank__"
            step1_summary.value = (
                f"Dipilih: {doc.label} — dokumen dibuat dari halaman kosong "
                "(tidak perlu template Word)."
            )
            step1_summary.color = ft.Colors.BLUE_900
        else:
            state["template_path"] = None
            step1_summary.value = (
                f"Dipilih: {doc.label} — template bawaan "
                + (f"tersedia ({doc.template_filename})." if builtin
                   else f"BELUM tersedia ({doc.template_filename}); Anda dapat mengunggah sendiri di langkah 2.")
            )
            if doc.input_template_path:
                step1_summary.value += (
                    " Data memakai format input khusus — template Excel-nya dapat "
                    "diunduh di langkah 3."
                )
            step1_summary.color = ft.Colors.BLUE_900 if builtin else ft.Colors.AMBER_800
        log(f"Dokumen dipilih: {doc.label}", "OK")
        update_nav()
        page.update()

    def make_doc_card(doc):
        icon = DOC_ICONS.get(doc.group, ft.Icons.INSERT_DRIVE_FILE_OUTLINED)
        card = ft.Container(
            data={"group": doc.group},
            content=ft.Row(
                [
                    ft.Icon(icon, size=26, color=ft.Colors.BLUE_GREY_300),
                    ft.Column(
                        [
                            ft.Text(doc.label, size=14, weight=ft.FontWeight.BOLD,
                                    color=ft.Colors.GREY_900),
                            ft.Text(doc.description, size=11,
                                    color=ft.Colors.GREY_600, max_lines=2),
                        ],
                        spacing=2,
                        expand=True,
                    ),
                ],
                spacing=12,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            width=460, height=78, padding=14, border_radius=10,
            # jarak kanan agar kartu tidak tertutup scrollbar
            margin=ft.Margin.only(right=14),
            bgcolor=ft.Colors.WHITE,
            border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
            on_click=lambda e, d=doc.id: select_doc(d),
            animate=150,
        )
        doc_cards[doc.id] = card
        return card

    doc_grid_controls = []
    for group in GROUP_ORDER:
        doc_grid_controls.append(
            ft.Text(group.upper(), size=12, weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_GREY_400)
        )
        doc_grid_controls.append(
            ft.Row(
                [make_doc_card(d) for d in documents_by_group(group)],
                wrap=True, spacing=10, run_spacing=10,
            )
        )

    panel1 = ft.Column(
        [
            ft.Text("Pilih Dokumen yang Ingin Dibuat", size=20,
                    weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
            ft.Text("Langkah 1 dari 5 — pilih satu jenis dokumen untuk seluruh sesi ini.",
                    size=13, color=ft.Colors.GREY_600),
            ft.Divider(height=1, color=ft.Colors.GREY_300),
            ft.Container(height=4),
            ft.Column(doc_grid_controls, spacing=10, expand=True,
                      scroll=ft.ScrollMode.ALWAYS),
            ft.Container(
                content=step1_summary, padding=12, border_radius=8,
                bgcolor=ft.Colors.GREY_50,
                border=ft.Border.all(1, ft.Colors.GREY_200),
            ),
        ],
        spacing=10, expand=True,
    )

    # ------------------------------------------------------------------ #
    # LANGKAH 3 - Upload & verifikasi data                                #
    # ------------------------------------------------------------------ #
    data_file_chip = ft.Text("Belum ada file dipilih.", italic=True,
                             color=ft.Colors.GREY_600, size=13, expand=True)
    verify_area = ft.Column(visible=False, spacing=12)

    async def download_input_template(e=None):
        """Simpan template Excel data ke komputer user.

        Dipakai bersama oleh Langkah 1 (kartu format Lampiran SPK) dan
        Langkah 3 (tombol Download Template Excel). Untuk grup Lampiran SPK
        menyalin bundel input/00_input_lampiran_spk.xlsx apa adanya; grup
        lain memakai generator generik.
        """
        doc = current_doc()
        bundled = doc.input_template_path if doc else None
        suggested = (os.path.basename(bundled) if bundled
                     else "template_database_administrasi.xlsx")
        save_path = await save_tpl_picker.save_file(
            dialog_title="Simpan Template Excel Data",
            file_name=suggested,
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
        )
        if not save_path:
            return
        save_path = ensure_extension(save_path, "xlsx")
        try:
            if bundled:
                # Grup Lampiran SPK → salin template input bawaan apa adanya
                shutil.copyfile(bundled, save_path)
                log(f"Template input bawaan disimpan: {save_path} "
                    f"(format {doc.group})", "OK")
            elif doc and doc.group == "Bukti Terima":
                bukti_terima.generate_input_template(save_path)
                log(f"Template Bukti Terima disimpan: {save_path}", "OK")
            else:
                generate_template(save_path)
                log(f"Template data disimpan: {save_path}", "OK")
            show_snackbar("Template Excel berhasil disimpan.", ft.Colors.GREEN_700)
        except Exception as ex:
            log(f"Gagal menyimpan template: {ex}", "ERROR")
            show_snackbar(f"Gagal menyimpan template: {ex}", ft.Colors.RED_700)

    def process_data_file(path: str):
        state["file_path"] = path
        state["data_ok"] = False
        doc = current_doc()
        template_path = state["template_path"]
        if template_path == "__blank__":
            template_path = None
        ok, errors, dfs = validate_document_input(
            doc, path, template_path=template_path
        )
        state["dfs"], state["errors"] = dfs, errors
        name = os.path.basename(path)
        if not ok:
            verify_area.controls = [
                ft.Container(
                    content=ft.Column(
                        [ft.Row([
                            ft.Icon(ft.Icons.ERROR_OUTLINE, color=ft.Colors.RED_600),
                            ft.Text("File tidak valid:", weight=ft.FontWeight.BOLD,
                                    color=ft.Colors.RED_900),
                        ]),
                        *[ft.Text(f"•  {err}", size=12, color=ft.Colors.RED_800)
                          for err in errors]],
                        spacing=6,
                    ),
                    bgcolor=ft.Colors.RED_50,
                    border=ft.Border.all(1.5, ft.Colors.RED_300),
                    border_radius=8, padding=14,
                )
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} — TIDAK VALID"
            data_file_chip.color = ft.Colors.RED_700
            log(f"Data diverifikasi: {name} → GAGAL", "ERROR")
            for err in errors:
                log(f"  {err}", "WARN")
            update_nav()
            page.update()
            return

        state["data_ok"] = True
        if doc and doc.id in ("lampiran_spk_ppl", "lampiran_spk_pml"):
            # Ringkasan khusus format Lampiran SPK (per-sheet + jumlah petugas)
            try:
                ctx = lampiran_spk.prepare_context(dfs)
                n_ppl = len(lampiran_spk._officer_list(ctx, "nik_ppl"))
                n_pml = len(lampiran_spk._officer_list(ctx, "nik_pml"))
            except Exception:
                n_ppl = n_pml = 0
            verify_area.controls = [
                ft.Row(
                    [
                        stat_box("Petugas Lapangan (PPL)", str(n_ppl), good=True),
                        stat_box("Petugas Pemeriksa Lapangan (PML)", str(n_pml), good=True),
                        stat_box("Sheet Wajib Terbaca",
                                 f"{sum(1 for s in lampiran_spk.REQUIRED_SCHEMA if s in dfs)}"
                                 f"/{len(lampiran_spk.REQUIRED_SCHEMA)}",
                                 good=all(s in dfs for s in lampiran_spk.REQUIRED_SCHEMA)),
                    ],
                    spacing=10,
                ),
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} — Lampiran SPK"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} → VALID (format Lampiran SPK)", "OK")
            log(f"  Petugas terdaftar — PPL: {n_ppl} orang, "
                f"PML: {n_pml} orang", "INFO")
            update_nav()
            page.update()
            return

        if doc and doc.group == "BAPP Termin 1":
            # Ringkasan khusus format BAPP T1
            bapp_mod = bapp_ppl if doc.kind == "ppl" else bapp_pml
            df_input = dfs.get(bapp_mod.SHEET_NAME)
            n_rows = len(df_input) if df_input is not None else 0
            n_with_links = 0
            if df_input is not None and bapp_mod.LINK_COLUMN in df_input.columns:
                n_with_links = sum(
                    1 for v in df_input[bapp_mod.LINK_COLUMN]
                    if str(v).strip() and str(v).strip().lower() != "nan"
                )
            role_label = "PPL" if doc.kind == "ppl" else "PML"
            verify_area.controls = [
                ft.Row(
                    [
                        stat_box(f"Total {role_label}", str(n_rows), good=n_rows > 0),
                        stat_box("Dengan Screenshot", str(n_with_links),
                                 good=True),
                        stat_box("Sheet", bapp_mod.SHEET_NAME, good=True),
                    ],
                    spacing=10,
                ),
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} \u2014 BAPP T1 {role_label}"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} \u2192 VALID "
                f"({n_rows} baris, {n_with_links} dengan screenshot)", "OK")
            update_nav()
            page.update()
            return

        if doc and doc.group == "BAPP Termin 2":
            # Ringkasan khusus format BAPP T2
            bapp_mod = bapp_ppl_t2 if doc.kind == "ppl" else bapp_pml_t2
            df_input = dfs.get(bapp_mod.SHEET_NAME)
            n_rows = len(df_input) if df_input is not None else 0
            n_with_links = 0
            if df_input is not None and bapp_mod.LINK_COLUMN in df_input.columns:
                n_with_links = sum(
                    1 for v in df_input[bapp_mod.LINK_COLUMN]
                    if str(v).strip() and str(v).strip().lower() != "nan"
                )
            role_label = "PPL" if doc.kind == "ppl" else "PML"
            verify_area.controls = [
                ft.Row(
                    [
                        stat_box(f"Total {role_label}", str(n_rows), good=n_rows > 0),
                        stat_box("Dengan Screenshot", str(n_with_links),
                                 good=True),
                        stat_box("Sheet", bapp_mod.SHEET_NAME, good=True),
                    ],
                    spacing=10,
                ),
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} \u2014 BAPP T2 {role_label}"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} \u2192 VALID "
                f"({n_rows} baris, {n_with_links} dengan screenshot)", "OK")
            update_nav()
            page.update()
            return

        if doc and doc.group == "SPP Termin 1":
            # Ringkasan khusus format SPP
            df_mitra = dfs.get(spp.SHEET_DATA_MITRA)
            n_ppl = 0
            n_pml = 0
            if df_mitra is not None and not df_mitra.empty:
                n_ppl = sum(
                    1 for v in df_mitra[spp.COL_JABATAN]
                    if str(v).strip().upper() == "PPL"
                )
                n_pml = sum(
                    1 for v in df_mitra[spp.COL_JABATAN]
                    if str(v).strip().upper() == "PML"
                )
            role_label = "PPL" if doc.kind == "ppl" else "PML"
            n_target = n_ppl if doc.kind == "ppl" else n_pml
            df_alok = dfs.get(spp.SHEET_ALOKASI)
            verify_area.controls = [
                ft.Row(
                    [
                        stat_box(f"Total {role_label}", str(n_target),
                                 good=n_target > 0),
                        stat_box("Sheet data_mitra",
                                 str(len(df_mitra)) if df_mitra is not None else "0",
                                 good=df_mitra is not None and len(df_mitra) > 0),
                        stat_box("Sheet alokasi_usaha",
                                 str(len(df_alok)) if df_alok is not None else "0",
                                 good=True),
                    ],
                    spacing=10,
                ),
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} \u2014 SPP Termin 1 {role_label}"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} \u2192 VALID "
                f"(PPL: {n_ppl}, PML: {n_pml})", "OK")
            update_nav()
            page.update()
            return

        if doc and doc.group == "SPP Termin 2":
            df_spp_t2 = dfs.get(spp_t2.SHEET_DATA_MITRA)
            n_rows = 0
            role_label = "PPL" if doc.kind == "ppl" else "PML"
            if df_spp_t2 is not None and not df_spp_t2.empty:
                n_rows = sum(
                    1 for value in df_spp_t2[spp.COL_JABATAN]
                    if str(value).strip().upper() == role_label
                )
            df_alokasi = dfs.get(spp_t2.SHEET_ALOKASI)
            verify_area.controls = [
                ft.Row(
                    [
                        stat_box(f"Total {role_label}", str(n_rows), good=n_rows > 0),
                        stat_box("Sheet data_mitra",
                                 str(len(df_spp_t2)) if df_spp_t2 is not None else "0",
                                 good=df_spp_t2 is not None and len(df_spp_t2) > 0),
                        stat_box("Sheet alokasi_usaha",
                                 str(len(df_alokasi)) if df_alokasi is not None else "0",
                                 good=df_alokasi is not None),
                    ],
                    spacing=10,
                ),
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} — SPP Termin 2 {role_label}"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} → VALID ({n_rows} baris)", "OK")
            update_nav()
            page.update()
            return

        if doc and doc.group == "BAST":
            # Ringkasan khusus format BAST
            df_mitra_b = dfs.get(bast.SHEET_NAME)
            n_ppl_b = 0
            n_pml_b = 0
            if df_mitra_b is not None and not df_mitra_b.empty:
                n_ppl_b = sum(
                    1 for v in df_mitra_b["jabatan"]
                    if str(v).strip().lower() == "ppl"
                )
                n_pml_b = sum(
                    1 for v in df_mitra_b["jabatan"]
                    if str(v).strip().lower() == "pml"
                )
            role_label = "PPL" if doc.kind == "ppl" else "PML"
            n_target = n_ppl_b if doc.kind == "ppl" else n_pml_b
            n_tugas = len(dfs["alokasi_tugas"]) if "alokasi_tugas" in dfs else 0
            verify_area.controls = [
                ft.Row(
                    [
                        stat_box(f"Total {role_label}", str(n_target),
                                 good=n_target > 0),
                        stat_box("Sheet data_mitra",
                                 str(len(df_mitra_b)) if df_mitra_b is not None else "0",
                                 good=df_mitra_b is not None and len(df_mitra_b) > 0),
                        stat_box("Baris alokasi_tugas",
                                 str(n_tugas),
                                 good=n_tugas > 0),
                    ],
                    spacing=10,
                ),
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} \u2014 BAST {role_label}"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} \u2192 VALID "
                f"(PPL: {n_ppl_b}, PML: {n_pml_b})", "OK")
            update_nav()
            page.update()
            return

        if doc and doc.group == "Bukti Terima":
            # Ringkasan khusus format Bukti Terima
            df_bt = dfs.get(bukti_terima.SHEET_NAME)
            n_rows = len(df_bt) if df_bt is not None else 0
            n_with_foto = 0
            if df_bt is not None and bukti_terima.REQUIRED_COLUMNS[-1] in df_bt.columns:
                n_with_foto = sum(
                    1 for v in df_bt[bukti_terima.REQUIRED_COLUMNS[-1]]
                    if str(v).strip() and str(v).strip().lower() != "nan"
                )
            verify_area.controls = [
                ft.Row(
                    [
                        stat_box("Total Petugas", str(n_rows), good=n_rows > 0),
                        stat_box("Dengan Link Foto", str(n_with_foto),
                                 good=n_with_foto > 0),
                        stat_box("Sheet", bukti_terima.SHEET_NAME, good=True),
                    ],
                    spacing=10,
                ),
            ]
            verify_area.visible = True
            data_file_chip.value = f"{name} \u2014 Bukti Terima"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} \u2192 VALID "
                f"({n_rows} petugas, {n_with_foto} dengan link foto)", "OK")
            update_nav()
            page.update()
            return

        analysis = analyze_nulls(dfs.get("data_petugas"))
        warn = analysis["rows_with_null"] > 0 or len(analysis["cols_with_null"]) > 0
        verify_area.controls = [
            ft.Row(
                [
                    stat_box("Total Kolom", str(analysis["total_cols"]), good=True),
                    stat_box("Baris dengan Null", f"{analysis['rows_with_null']}",
                             good=analysis["rows_with_null"] == 0),
                    stat_box("Kolom dengan Null", f"{len(analysis['cols_with_null'])}",
                             good=len(analysis["cols_with_null"]) == 0),
                ],
                spacing=10,
            ),
        ]
        verify_area.visible = True
        data_file_chip.value = f"{name}"
        data_file_chip.color = ft.Colors.GREY_900
        log(f"Data diverifikasi: {name} → VALID "
            f"({analysis['total_rows']} baris, {analysis['total_cols']} kolom)", "OK")
        if warn:
            log(f"Nilai kosong: {analysis['rows_with_null']} baris, "
                f"{len(analysis['cols_with_null'])} kolom", "WARN")
        update_nav()
        page.update()

    async def on_pick_data(e):
        files = await data_picker.pick_files(
            dialog_title="Pilih File Excel Database Administrasi",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xls"],
        )
        if files:
            process_data_file(files[0].path)

    panel2 = ft.Column(
        [
            ft.Text("Upload Data", size=20, weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_900),
            ft.Text("Langkah 3 dari 5 — siapkan database administrasi dalam format Excel.",
                    size=13, color=ft.Colors.GREY_600),
            ft.Divider(height=1, color=ft.Colors.GREY_300),
            ft.Container(height=4),
            ft.Row(
                [
                    ft.Button(
                        "Download Template Excel",
                        icon=ft.Icons.DOWNLOAD_ROUNDED,
                        style=ft.ButtonStyle(
                            color=ft.Colors.BLUE_800, bgcolor=ft.Colors.BLUE_50,
                            side=ft.BorderSide(1, ft.Colors.BLUE_300),
                            shape=ft.RoundedRectangleBorder(radius=8),
                            padding=ft.Padding.symmetric(horizontal=15, vertical=12),
                        ),
                        on_click=download_input_template,
                    ),
                    ft.Button(
                        "Pilih File Excel (.xlsx)",
                        icon=ft.Icons.FOLDER_OPEN_ROUNDED,
                        style=ft.ButtonStyle(
                            color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_800,
                            shape=ft.RoundedRectangleBorder(radius=8),
                            padding=ft.Padding.symmetric(horizontal=18, vertical=12),
                        ),
                        on_click=on_pick_data,
                    ),
                ],
                spacing=10,
            ),
            ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.ATTACH_FILE_ROUNDED, size=18,
                            color=ft.Colors.GREY_600),
                    data_file_chip,
                ]),
                bgcolor=ft.Colors.GREY_50,
                border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
                border_radius=8, padding=12,
            ),
            verify_area,
        ],
        spacing=10, scroll=ft.ScrollMode.AUTO, expand=True,
    )

    # ------------------------------------------------------------------ #
    # LANGKAH 2 - Template dokumen                                        #
    # ------------------------------------------------------------------ #
    tpl_status_chip = ft.Text("Belum ada template dipilih.", italic=True,
                              color=ft.Colors.GREY_600, size=13, expand=True)
    tpl_expected_text = ft.Text("", size=13, color=ft.Colors.GREY_700)
    # --- Panel pratinjau isi template (.docx) yang diunggah ------------- #
    tpl_preview_stats = ft.Text("", size=12, color=ft.Colors.GREY_700)
    tpl_preview_ph_row = ft.Row(spacing=6, run_spacing=6, wrap=True)
    tpl_preview_body = ft.Column(spacing=3, scroll=ft.ScrollMode.AUTO)
    tpl_preview_area = ft.Container(
        visible=False,
        bgcolor=ft.Colors.BLUE_GREY_50,
        border=ft.Border.all(1, ft.Colors.BLUE_GREY_200),
        border_radius=8,
        padding=12,
        content=ft.Column(
            [
                ft.Row([
                    ft.Icon(ft.Icons.PREVIEW_ROUNDED, size=18,
                            color=ft.Colors.BLUE_700),
                    ft.Text("Pratinjau Template", size=13,
                            weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                ], spacing=8),
                tpl_preview_stats,
                tpl_preview_ph_row,
                ft.Container(
                    content=tpl_preview_body,
                    height=260,
                    bgcolor=ft.Colors.WHITE,
                    border=ft.Border.all(1, ft.Colors.GREY_200),
                    border_radius=6,
                    padding=10,
                ),
            ],
            spacing=8, tight=True,
        ),
    )

    def render_template_preview(path: str):
        """Tampilkan isi .docx terpilih pada panel pratinjau langkah 2."""
        tpl_preview_area.visible = False
        tpl_preview_stats.value = ""
        tpl_preview_ph_row.controls = []
        tpl_preview_body.controls = []
        if not HAS_DOCX:
            tpl_preview_stats.value = ("Pratinjau tidak tersedia — paket "
                                       "python-docx belum terpasang.")
            tpl_preview_stats.color = ft.Colors.AMBER_800
            tpl_preview_area.visible = True
            return
        try:
            from docx import Document
            from docx.table import Table
            from docx.text.paragraph import Paragraph

            d = Document(path)
            blocks, placeholders, seen = [], [], set()
            for child in d.element.body.iterchildren():
                tag = child.tag.rsplit("}", 1)[-1]
                if tag == "p":
                    txt = Paragraph(child, d).text.strip()
                    if txt:
                        blocks.append(txt)
                        for ph in PLACEHOLDER_RE.findall(txt):
                            if ph not in seen:
                                seen.add(ph)
                                placeholders.append(ph)
                elif tag == "tbl":
                    t = Table(child, d)
                    blocks.append(
                        f"[ Tabel {len(t.rows)} baris × {len(t.columns)} kolom ]")

            n_tables = sum(1 for b in blocks if b.startswith("[ Tabel"))
            n_paras = len(blocks) - n_tables
            tpl_preview_stats.value = (
                f"{os.path.basename(path)} — {n_paras} paragraf, {n_tables} tabel."
            )
            tpl_preview_stats.color = ft.Colors.GREY_700
            if placeholders:
                for ph in placeholders[:12]:
                    tpl_preview_ph_row.controls.append(ft.Container(
                        content=ft.Text(f"{{{{{ph}}}}}", size=11,
                                        color=ft.Colors.INDIGO_800),
                        bgcolor=ft.Colors.INDIGO_50,
                        border=ft.Border.all(1, ft.Colors.INDIGO_100),
                        border_radius=6,
                        padding=ft.Padding.only(left=8, right=8, top=3, bottom=3),
                    ))
                if len(placeholders) > 12:
                    tpl_preview_ph_row.controls.append(ft.Text(
                        f"+{len(placeholders) - 12} penanda lainnya",
                        size=11, color=ft.Colors.GREY_600))
            else:
                tpl_preview_ph_row.controls.append(ft.Text(
                    "Tidak ada penanda {{kolom}} ditemukan — data tidak akan "
                    "terisi otomatis.", size=11, color=ft.Colors.AMBER_800))

            MAX_BLOCKS = 200
            shown = blocks[:MAX_BLOCKS]
            for b in shown:
                is_tbl = b.startswith("[ Tabel")
                tpl_preview_body.controls.append(ft.Text(
                    b,
                    size=12,
                    color=(ft.Colors.BLUE_GREY_400 if is_tbl
                           else ft.Colors.GREY_900),
                    italic=is_tbl,
                ))
            if len(blocks) > MAX_BLOCKS:
                tpl_preview_body.controls.append(ft.Text(
                    f"… {len(blocks) - MAX_BLOCKS} bagian lainnya tidak "
                    "ditampilkan.", size=11, italic=True,
                    color=ft.Colors.GREY_500))
        except Exception as ex:
            log(f"Gagal membaca pratinjau template: {ex}", "WARN")
            tpl_preview_stats.value = f"Pratinjau gagal dibaca ({ex})."
            tpl_preview_stats.color = ft.Colors.RED_700
        tpl_preview_area.visible = True

    async def on_pick_template(e):
        files = await tpl_picker.pick_files(
            dialog_title="Pilih Template Dokumen (.docx)",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["docx"],
        )
        if not files:
            return
        path = files[0].path
        doc = current_doc()
        if doc and doc.builtin_template_path:
            try:
                expected_placeholders = extract_template_placeholders(
                    doc.builtin_template_path
                )
                validation = validate_template_placeholders(
                    path, expected_placeholders
                )
            except Exception as ex:
                log(f"Gagal memvalidasi template dokumen: {ex}", "ERROR")
                show_snackbar(
                    f"Template tidak dapat divalidasi: {ex}", ft.Colors.RED_700
                )
                return

            if not validation["is_valid"]:
                details = []
                if validation["missing"]:
                    details.append(
                        "penanda wajib hilang: "
                        + ", ".join(
                            f"{{{{{key}}}}}" for key in validation["missing"]
                        )
                    )
                if validation["unexpected"]:
                    details.append(
                        "penanda tidak dikenal: "
                        + ", ".join(
                            f"{{{{{key}}}}}" for key in validation["unexpected"]
                        )
                    )
                message = "Template ditolak — " + "; ".join(details) + "."
                log(message, "WARN")
                show_snackbar(message, ft.Colors.RED_700)
                return

        state["template_path"] = path
        state["template_source"] = "unggahan"
        name = os.path.basename(path)
        tpl_status_chip.value = f"{name}  (unggahan Anda)"
        tpl_status_chip.color = ft.Colors.GREY_900
        log(f"Template dokumen diunggah: {name}", "OK")
        render_template_preview(files[0].path)
        update_nav()
        page.update()

    def refresh_step3():
        doc = current_doc()
        if doc is None:
            return
        if doc.builtin_template_path:
            tpl_expected_text.value = (
                f"Dokumen terpilih: {doc.label}. Klik Download Template Word "
                f"untuk menyimpan {doc.template_filename}, sunting di Microsoft "
                f"Word, lalu unggah kembali di sini."
            )
            tpl_expected_text.color = ft.Colors.GREY_700
            btn_download_tpl.disabled = False
            btn_download_tpl.tooltip = (
                f"Simpan salinan {doc.template_filename} dari folder template/"
            )
        else:
            tpl_expected_text.value = (
                f"Dokumen terpilih: {doc.label}. Template bawaan "
                f"({doc.template_filename}) BELUM tersedia — siapkan template "
                f".docx Anda sendiri dan unggah di sini."
            )
            tpl_expected_text.color = ft.Colors.AMBER_900
            btn_download_tpl.disabled = True
            btn_download_tpl.tooltip = "Template bawaan akan disediakan kemudian"

    async def download_doc_template(e=None):
        """Simpan template dokumen bawaan ke komputer user.

        Mengikuti pola download_input_template (langkah 3): pengguna memilih
        lokasi lewat dialog Save, lalu berkas dari folder template/ disalin.
        """
        doc = current_doc()
        if not doc or not doc.builtin_template_path:
            show_snackbar("Template bawaan belum tersedia untuk dokumen ini.",
                          ft.Colors.AMBER_800)
            return
        save_path = await save_tpl_picker.save_file(
            dialog_title="Simpan Template Dokumen",
            file_name=doc.template_filename,
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["docx"],
        )
        if not save_path:
            return
        save_path = ensure_extension(save_path, "docx")
        try:
            shutil.copyfile(doc.builtin_template_path, save_path)
            log(f"Template dokumen disimpan: {save_path}", "OK")
            show_snackbar("Template dokumen berhasil disimpan — sunting di "
                          "Word, lalu unggah kembali di langkah ini.",
                          ft.Colors.GREEN_700)
        except Exception as ex:
            log(f"Gagal menyimpan template dokumen: {ex}", "ERROR")
            show_snackbar(f"Gagal menyimpan template dokumen: {ex}",
                          ft.Colors.RED_700)

    btn_download_tpl = ft.Button(
        "Download Template Word",
        icon=ft.Icons.DOWNLOAD_ROUNDED,
        style=ft.ButtonStyle(
            color=ft.Colors.BLUE_800, bgcolor=ft.Colors.BLUE_50,
            side=ft.BorderSide(1, ft.Colors.BLUE_300),
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.Padding.symmetric(horizontal=15, vertical=12),
        ),
        on_click=download_doc_template,
    )

    panel3 = ft.Column(
        [
            ft.Text("Template Dokumen", size=20, weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_900),
            ft.Text("Langkah 2 dari 5 — unduh template bawaan, sunting di "
                    "Word, lalu unggah kembali.",
                    size=13, color=ft.Colors.GREY_600),
            ft.Divider(height=1, color=ft.Colors.GREY_300),
            ft.Container(height=4),
            ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.INFO_OUTLINE_ROUNDED, size=18,
                            color=ft.Colors.BLUE_700),
                    tpl_expected_text, 
                ]),
                bgcolor=ft.Colors.BLUE_50,
                border=ft.Border.all(1, ft.Colors.BLUE_200),
                border_radius=8, padding=12,
            ),
            ft.Text(
                "Alur kerja: klik Download Template Word, sunting berkasnya di "
                "Microsoft Word lalu simpan, dan unggah hasilnya lewat tombol "
                "Pilih Template Word. Penanda {{nama_kolom}} pada template akan "
                "diganti dengan kolom dari file data Anda.",
                size=12, color=ft.Colors.GREY_600,
            ),
            ft.Row(
                [
                    btn_download_tpl,
                    ft.Button(
                        "Pilih Template Word (.docx)",
                        icon=ft.Icons.UPLOAD_FILE_ROUNDED,
                        style=ft.ButtonStyle(
                            color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_800,
                            shape=ft.RoundedRectangleBorder(radius=8),
                            padding=ft.Padding.symmetric(horizontal=18, vertical=12),
                        ),
                        on_click=on_pick_template,
                    ),
                ],
                spacing=10, wrap=True,
            ),
            ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.DESCRIPTION_ROUNDED, size=18,
                            color=ft.Colors.GREY_600),
                    tpl_status_chip,
                ]),
                bgcolor=ft.Colors.GREY_50,
                border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
                border_radius=8, padding=12,
            ),
            tpl_preview_area,
        ],
        spacing=10, scroll=ft.ScrollMode.AUTO, expand=True,
    )

    # ------------------------------------------------------------------ #

    # ------------------------------------------------------------------ #
    # LANGKAH 4 - Generate                                                #
    # ------------------------------------------------------------------ #
    gen_summary_text = ft.Text("Lengkapi langkah sebelumnya.", size=13,
                               color=ft.Colors.GREY_600)
    gen_progress = ft.ProgressBar(value=0, visible=False, bar_height=8,
                                  color=ft.Colors.BLUE_700,
                                  bgcolor=ft.Colors.BLUE_GREY_100)
    gen_status = ft.Text("", size=13, color=ft.Colors.BLUE_900, visible=False)
    gen_duration_box = ft.Container(visible=False)
    _gen_timer_task = None
    _gen_start_time = 0.0

    bapp_t1_link_warning = ft.Container(
        visible=False,
        content=ft.Row(
            [
                ft.Icon(ft.Icons.WARNING_AMBER_ROUNDED,
                        color=ft.Colors.AMBER_800, size=20),
                ft.Text(
                    "PENTING: Pastikan tautan file Google Drive bukti dukung "
                    "(gambar/PDF) memiliki akses "
                    "\"Anyone with the link\" "
                    "(Siapa saja yang memiliki tautan). "
                    "Jika tautan dikumpulkan dalam folder dari "
                    "Google Forms (gForm), pastikan folder tersebut "
                    "juga diatur \"Anyone with the link\".",
                    size=12, color=ft.Colors.AMBER_900, expand=True,
                ),
            ],
            spacing=8,
            vertical_alignment=ft.CrossAxisAlignment.START,
        ),
        bgcolor=ft.Colors.AMBER_50,
        border=ft.Border.all(1, ft.Colors.AMBER_300),
        border_radius=8,
        padding=12,
    )
    evidence_orientation_group = ft.RadioGroup(
        value="automatic",
        disabled=True,
        content=ft.Column(
            [
                ft.Radio(value="automatic", label="Otomatis"),
                ft.Radio(value="portrait", label="Potret"),
                ft.Radio(value="landscape", label="Lanskap"),
            ],
            spacing=2,
        ),
        on_change=lambda e: state.update(
            {"evidence_image_orientation": e.control.value}
        ),
    )
    evidence_orientation_option = ft.Column(
        [
            ft.Divider(height=12, color=ft.Colors.BLUE_200),
            ft.Text(
                "Orientasi Gambar pada Halaman Khusus",
                size=13,
                weight=ft.FontWeight.BOLD,
                color=ft.Colors.BLUE_900,
            ),
            evidence_orientation_group,
            ft.Text(
                "Petunjuk: Pilih Otomatis agar SIOMAY menentukan orientasi "
                "berdasarkan perbandingan lebar dan tinggi gambar. "
                "Pilih Potret untuk mempertahankan tampilan lama. "
                "Pilih Lanskap untuk memutar setiap gambar 90° searah jarum jam "
                "dan memperbesarnya.",
                size=12,
                color=ft.Colors.GREY_700,
                italic=True,
            ),
        ],
        spacing=4,
        visible=False,
    )

    def on_evidence_layout_change(e):
        state["evidence_image_layout"] = e.control.value
        show_orientation = e.control.value == "dedicated_pages"
        evidence_orientation_option.visible = show_orientation
        evidence_orientation_group.disabled = not show_orientation or state["busy"]
        page.update()

    evidence_layout_group = ft.RadioGroup(
        value="grid",
        content=ft.Column(
            [
                ft.Radio(
                    value="grid",
                    label="Grid adaptif — maksimal 5 gambar per halaman, ukuran menyesuaikan",
                ),
                ft.Radio(
                    value="dedicated_pages",
                    label="Halaman khusus — satu gambar per halaman",
                ),
            ],
            spacing=2,
        ),
        on_change=on_evidence_layout_change,
    )
    evidence_layout_option = ft.Container(
        visible=False,
        content=ft.Column(
            [
                ft.Text(
                    "Tata Letak Bukti Dukung",
                    size=13,
                    weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_900,
                ),
                ft.Text(
                    "Pilih cara gambar ditempatkan. Bukti dukung berupa .PDF akan selalu "
                    "ditempatkan pada halaman khusus.",
                    size=12,
                    color=ft.Colors.GREY_700,
                ),
                evidence_layout_group,
                evidence_orientation_option,
            ],
            spacing=4,
        ),
        bgcolor=ft.Colors.BLUE_50,
        border=ft.Border.all(1, ft.Colors.BLUE_200),
        border_radius=8,
        padding=12,
    )
    btn_generate = None

    def refresh_step4():
        doc = current_doc()
        if doc is None or not state["data_ok"] or not state["template_path"]:
            return
        # Tampilkan peringatan akses link untuk workflow dengan bukti Drive.
        bapp_t1_link_warning.visible = (
            doc.group in ("BAPP Termin 1", "BAPP Termin 2", "BAST", "Bukti Terima")
        )
        evidence_layout_option.visible = doc.group in ("BAPP Termin 2", "BAST")
        evidence_layout_group.value = state["evidence_image_layout"]
        evidence_orientation_group.value = state["evidence_image_orientation"]
        evidence_orientation_option.visible = (
            state["evidence_image_layout"] == "dedicated_pages"
        )
        evidence_orientation_group.disabled = (
            state["evidence_image_layout"] != "dedicated_pages" or state["busy"]
        )
        if doc.group == "Lampiran SPK" and doc.kind:
            try:
                ctx = lampiran_spk.prepare_context(state["dfs"])
                n_rows = len(lampiran_spk._officer_list(
                    ctx, f"nik_{doc.kind}"))
                extra = f" ({n_rows} petugas {doc.kind.upper()})"
            except Exception:
                extra = " (format Lampiran SPK)"
        elif doc.group == "BAPP Termin 1":
            df_input = state["dfs"].get(bapp_pml.SHEET_NAME)
            n_rows = len(df_input) if df_input is not None else 0
            extra = f" ({n_rows} petugas PML)"
        elif doc.group == "BAPP Termin 2":
            df_input = state["dfs"].get(bapp_pml_t2.SHEET_NAME)
            n_rows = len(df_input) if df_input is not None else 0
            extra = f" ({n_rows} petugas PML)"
        elif doc.group == "SPP Termin 1":
            df_mitra = state["dfs"].get(spp.SHEET_DATA_MITRA)
            n_total = len(df_mitra) if df_mitra is not None else 0
            extra = f" ({n_total} petugas total)"
        elif doc.group == "SPP Termin 2":
            df_input = state["dfs"].get(spp_t2.SHEET_DATA_MITRA)
            n_rows = 0
            if df_input is not None and not df_input.empty:
                n_rows = sum(
                    1 for value in df_input[spp.COL_JABATAN]
                    if str(value).strip().lower() == doc.kind
                )
            extra = f" ({n_rows} petugas {doc.kind.upper()})"
        elif doc.group == "BAST":
            df_mitra_b = state["dfs"].get(bast.SHEET_NAME)
            n_target = 0
            if df_mitra_b is not None and not df_mitra_b.empty:
                n_target = sum(
                    1 for v in df_mitra_b["jabatan"]
                    if str(v).strip().lower() == doc.kind
                )
            extra = f" ({n_target} {doc.kind.upper()})"
        elif doc.group == "Bukti Terima":
            df_bt = state["dfs"].get(bukti_terima.SHEET_NAME)
            n_rows = len(df_bt) if df_bt is not None else 0
            extra = f" ({n_rows} petugas)"
        else:
            n_rows = len(state["dfs"].get("data_petugas", []))
            extra = f" ({n_rows} baris)"
        template_label = (
            "— (dibuat dari dokumen kosong)"
            if state["template_path"] == "__blank__"
            else os.path.basename(state["template_path"])
        )
        gen_summary_text.value = (
            f"Dokumen : {doc.label}\n"
            f"Sumber data : {os.path.basename(state['file_path'])}{extra}\n"
            f"Template : {template_label}"
        )

    async def _gen_timer_loop():
        """Loop asinkron untuk memperbarui timer live pada langkah 4."""
        nonlocal _gen_start_time
        try:
            while state["busy"]:
                elapsed = time.time() - _gen_start_time
                clock_str = format_timer_clock(elapsed)
                btn_generate.text = f"Sedang Memproses Dokumen… ({clock_str})"
                page.update()
                await asyncio.sleep(0.5)
        except asyncio.CancelledError:
            pass

    def _gen_ui_start():
        """Persiapan UI/state yang sama untuk semua jalur generate."""
        nonlocal _gen_timer_task, _gen_start_time
        state["busy"] = True
        state["generated_files"] = []
        state["generation_done"] = False
        state["gen_duration"] = None
        _gen_start_time = time.time()
        gen_duration_box.visible = False

        btn_generate.disabled = True
        evidence_layout_group.disabled = True
        evidence_orientation_group.disabled = True
        btn_generate.text = "Sedang Memproses Dokumen… (00:00)"
        btn_generate.icon = ft.Icons.HOURGLASS_TOP_ROUNDED
        gen_progress.visible = True
        gen_progress.value = 0
        gen_status.visible = True
        if _gen_timer_task and not _gen_timer_task.done():
            _gen_timer_task.cancel()
        _gen_timer_task = asyncio.create_task(_gen_timer_loop())
        update_nav()
        page.update()

    def _gen_ui_finish():
        """Reset UI/state setelah proses generate selesai / error."""
        nonlocal _gen_timer_task, _gen_start_time
        if _gen_timer_task and not _gen_timer_task.done():
            _gen_timer_task.cancel()
        duration = time.time() - _gen_start_time
        state["gen_duration"] = duration
        state["busy"] = False
        btn_generate.disabled = False
        evidence_layout_group.disabled = False
        evidence_orientation_group.disabled = (
            state["evidence_image_layout"] != "dedicated_pages"
        )
        btn_generate.text = "Mulai Generate Dokumen"
        btn_generate.icon = ft.Icons.PLAY_ARROW_ROUNDED

        # Tampilkan kotak lilac info durasi jika dokumen berhasil dibuat
        if state["generation_done"] and len(state["generated_files"]) > 0:
            doc = current_doc()
            doc_name = doc.label if doc else "Dokumen"
            n_files = len(state["generated_files"])
            gen_duration_box.content = duration_info_box(
                title="Waktu Pembuatan Dokumen Selesai",
                items=[
                    ("Dokumen", f"{doc_name} ({n_files} berkas)"),
                    ("Waktu pembuatan (generate)", format_duration(duration)),
                ],
                icon=ft.Icons.TIMER_ROUNDED,
            )
            gen_duration_box.visible = True
        else:
            gen_duration_box.visible = False

        update_nav()
        page.update()

    async def generate_lampiran_spk(doc, out_dir):
        """Populasi Lampiran SPK (PPL/PML) — port notebook generator/.

        Memakai src/lampiran_spk.iter_generate() yang melempar event
        {t: log|file|progress|done}; setiap event log ditampilkan ke panel
        dengan gaya yang sama persis seperti di Jupyter notebook.
        """
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen…"
        page.update()

        kind = doc.kind  # 'ppl' | 'pml' (dari pilihan Langkah 1)
        log("=" * 46, "STEP")
        log(f"MEMULAI GENERATE — {doc.label} ({kind.upper()})", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")

        try:
            ctx = lampiran_spk.prepare_context(state["dfs"])
            for ev in lampiran_spk.iter_generate(
                    kind, ctx, state["template_path"], out_dir):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}…")
                    page.update()
                    await asyncio.sleep(0)  # beri kesempatan UI bernapas
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI — {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder sementara: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_bapp_pml():
        """Populasi BAPP T1 PML — screenshot grid bukti dukung."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        out_dir = tempfile.mkdtemp(prefix="gen_bapp_pml_")
        log("=" * 46, "STEP")
        log("MEMULAI GENERATE \u2014 BAPP T1 PML", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")

        try:
            for ev in bapp_pml.iter_generate(
                    state["dfs"], state["template_path"], out_dir):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}\u2026")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI \u2014 {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_bapp_ppl():
        """Populasi BAPP T1 PPL — screenshot grid bukti dukung."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        out_dir = tempfile.mkdtemp(prefix="gen_bapp_ppl_")
        log("=" * 46, "STEP")
        log("MEMULAI GENERATE \u2014 BAPP T1 PPL", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")

        try:
            for ev in bapp_ppl.iter_generate(
                    state["dfs"], state["template_path"], out_dir):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}\u2026")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI \u2014 {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_bapp_pml_t2():
        """Populasi BAPP T2 PML dengan tata letak bukti pilihan pengguna."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        out_dir = tempfile.mkdtemp(prefix="gen_bapp_pml_t2_")
        log("=" * 46, "STEP")
        log("MEMULAI GENERATE \u2014 BAPP T2 PML", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")
        image_layout = state["evidence_image_layout"]
        image_orientation = state["evidence_image_orientation"]
        layout_label = "Grid adaptif" if image_layout == "grid" else "Halaman khusus"
        orientation_label = {
            "portrait": "Potret", "landscape": "Lanskap", "automatic": "Otomatis",
        }[image_orientation]
        log(f"Gambar   : {layout_label}", "INFO")
        log(f"Orientasi: {orientation_label}", "INFO")

        try:
            for ev in bapp_pml_t2.iter_generate(
                    state["dfs"], state["template_path"], out_dir,
                    image_layout=image_layout,
                    image_orientation=image_orientation):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}\u2026")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI \u2014 {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_bapp_ppl_t2():
        """Populasi BAPP T2 PPL dengan tata letak bukti pilihan pengguna."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        out_dir = tempfile.mkdtemp(prefix="gen_bapp_ppl_t2_")
        log("=" * 46, "STEP")
        log("MEMULAI GENERATE \u2014 BAPP T2 PPL", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")
        image_layout = state["evidence_image_layout"]
        image_orientation = state["evidence_image_orientation"]
        layout_label = "Grid adaptif" if image_layout == "grid" else "Halaman khusus"
        orientation_label = {
            "portrait": "Potret", "landscape": "Lanskap", "automatic": "Otomatis",
        }[image_orientation]
        log(f"Gambar   : {layout_label}", "INFO")
        log(f"Orientasi: {orientation_label}", "INFO")

        try:
            for ev in bapp_ppl_t2.iter_generate(
                    state["dfs"], state["template_path"], out_dir,
                    image_layout=image_layout,
                    image_orientation=image_orientation):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}\u2026")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI \u2014 {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_spp():
        """Populasi SPP (PPL & PML) -- Surat Pernyataan Penyelesaian."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        doc = current_doc()
        kind = doc.kind  # 'ppl' | 'pml'
        out_dir = tempfile.mkdtemp(prefix="gen_spp_")
        log("=" * 46, "STEP")
        log(f"MEMULAI GENERATE \u2014 SPP {kind.upper()}", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")

        try:
            for ev in spp.iter_generate(
                    kind, state["dfs"], state["template_path"], out_dir):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}\u2026")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI \u2014 {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_spp_t2():
        """Populasi SPP Termin 2 untuk PPL atau PML."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen…"
        page.update()

        doc = current_doc()
        kind = doc.kind
        out_dir = tempfile.mkdtemp(prefix="gen_spp_t2_")
        log("=" * 46, "STEP")
        log(f"MEMULAI GENERATE — SPP Termin 2 {kind.upper()}", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")

        try:
            for ev in spp_t2.iter_generate(
                    kind, state["dfs"], state["template_path"], out_dir):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}…")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI — {n_ok} dokumen.", "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_bast():
        """Populasi BAST (PPL & PML) -- Berita Acara Serah Terima."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        doc = current_doc()
        kind = doc.kind  # 'ppl' | 'pml'
        out_dir = tempfile.mkdtemp(prefix="gen_bast_")
        log("=" * 46, "STEP")
        log(f"MEMULAI GENERATE \u2014 BAST {kind.upper()}", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")
        image_layout = state["evidence_image_layout"]
        image_orientation = state["evidence_image_orientation"]
        layout_label = "Grid adaptif" if image_layout == "grid" else "Halaman khusus"
        orientation_label = {
            "portrait": "Potret", "landscape": "Lanskap", "automatic": "Otomatis",
        }[image_orientation]
        log(f"Bukti    : {layout_label}", "INFO")
        log(f"Orientasi: {orientation_label}", "INFO")

        try:
            for ev in bast.iter_generate(
                    kind, state["dfs"], state["template_path"], out_dir,
                    image_layout=image_layout,
                    image_orientation=image_orientation):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Mengisi dokumen {ev.get('done', 0)} dari {total}\u2026")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI \u2014 {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def generate_bukti_terima():
        """Populasi Bukti Terima Paket Internet — grid foto 2x2 per halaman A4."""
        _gen_ui_start()
        gen_status.value = "Mempersiapkan dokumen Bukti Terima\u2026"
        page.update()

        out_dir = tempfile.mkdtemp(prefix="gen_bt_")
        log("=" * 46, "STEP")
        log("MEMULAI GENERATE \u2014 Bukti Terima Paket Internet", "STEP")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")
        log("Template : \u2014 (dokumen dibuat dari halaman kosong)", "INFO")

        try:
            for ev in bukti_terima.iter_generate(state["dfs"], out_dir):
                t = ev.get("t")
                if t == "log":
                    log(ev["msg"], ev.get("level", "INFO"))
                elif t == "file":
                    state["generated_files"].append(ev["path"])
                elif t == "progress":
                    total = max(ev.get("total", 1), 1)
                    gen_progress.max = total
                    gen_progress.value = ev.get("done", 0) / total
                    gen_status.value = (
                        f"Memproses petugas {ev.get('done', 0)} dari {total}\u2026")
                    page.update()
                    await asyncio.sleep(0)
                elif t == "done":
                    state["generated_files"] = list(ev.get("generated", []))

            n_ok = len(state["generated_files"])
            state["generation_done"] = n_ok > 0
            gen_progress.value = 1
            gen_status.value = f"Selesai: {n_ok} dokumen berhasil dibuat."
            log(f"GENERATE SELESAI \u2014 {n_ok} dokumen.",
                "OK" if n_ok else "ERROR")
            log(f"Folder output: {out_dir}", "INFO")
        except Exception as ex:
            gen_status.value = f"Generasi gagal: {ex}"
            log(f"Generasi gagal: {ex}", "ERROR")
        finally:
            _gen_ui_finish()

    async def run_generation(e=None):
        if state["busy"]:
            return
        doc = current_doc()

        # ── Jalur khusus: grup Lampiran SPK (PPL / PML) ────────────────
        if doc and doc.id in ("lampiran_spk_ppl", "lampiran_spk_pml"):
            await generate_lampiran_spk(
                doc, tempfile.mkdtemp(prefix="gen_spk_"))
            return

        # ── Jalur khusus: grup BAPP Termin 1 ────────────────
        if doc and doc.id == "bapp_ppl_t1":
            await generate_bapp_ppl()
            return
        if doc and doc.id == "bapp_pml_t1":
            await generate_bapp_pml()
            return

        # ── Jalur khusus: grup BAPP Termin 2 ────────────────
        if doc and doc.id == "bapp_ppl_t2":
            await generate_bapp_ppl_t2()
            return
        if doc and doc.id == "bapp_pml_t2":
            await generate_bapp_pml_t2()
            return

        # ── Jalur khusus: grup SPP Termin 1 ────────────────
        if doc and doc.id in ("spp_ppl", "spp_pml"):
            await generate_spp()
            return

        # ── Jalur khusus: grup SPP Termin 2 ────────────────
        if doc and doc.id in ("spp_t2_ppl", "spp_t2_pml"):
            await generate_spp_t2()
            return

        # ── Jalur khusus: grup BAST ────────────────
        if doc and doc.id in ("bast_ppl", "bast_pml"):
            await generate_bast()
            return

        # ── Jalur khusus: grup Bukti Terima ────────────────
        if doc and doc.id == "bukti_terima":
            await generate_bukti_terima()
            return

        df = state["dfs"].get("data_petugas")
        if df is None or df.empty:
            log("Tidak ada baris data pada sheet 'data_petugas'.", "ERROR")
            show_snackbar("Data kosong!", ft.Colors.RED_700)
            return

        _gen_ui_start()
        out_dir = tempfile.mkdtemp(prefix="gen_docs_")

        log("=" * 46, "STEP")
        log(f"MEMULAI GENERATE — {doc.label} — {len(df)} dokumen", "STEP")
        if not HAS_DOCX:
            log("python-docx tidak tersedia; template hanya akan disalin tanpa diisi.", "WARN")

        rows = df.to_dict("records")
        total = len(rows)
        error_count = 0

        for i, record in enumerate(rows):
            values = build_values(record)
            who = values.get("nama_lengkap") or values.get("custNoRef") or f"baris {i + 1}"
            log(f"[{i + 1}/{total}] Menyiapkan dokumen untuk {who}", "INFO")
            out_name = f"{doc.prefix}_{i + 1:03d}_{slugify(who)}.docx"
            out_path = os.path.join(out_dir, out_name)
            try:
                if HAS_DOCX:
                    result = fill_row(state["template_path"], values, out_path)
                    filled = result["filled"]
                    unresolved = result["unresolved"]
                    preview = ", ".join(filled[:6]) + ("…" if len(filled) > 6 else "")
                    if filled:
                        log(f"    Mengisi {len(filled)} field: {preview}")
                    for key in unresolved:
                        log(f"    Placeholder {{{{{key}}}}} tidak terisi "
                            f"(kolom tidak ada / nilai kosong)", "WARN")
                else:
                    copy_template(state["template_path"], out_path)
                state["generated_files"].append(out_path)
                log(f"    Tersimpan: {out_name}", "OK")
            except Exception as ex:
                error_count += 1
                log(f"    GAGAL memproses baris {i + 1}: {ex}", "ERROR")
            gen_progress.value = (i + 1) / total
            gen_status.value = f"Memproses {i + 1} dari {total} dokumen…"
            page.update()
            await asyncio.sleep(0)  # beri kesempatan UI bernapas

        state["generation_done"] = (
            len(state["generated_files"]) > 0 and error_count < total
        )
        gen_status.value = (
            f"Selesai: {len(state['generated_files'])}/{total} dokumen berhasil."
        )
        log(f"GENERATE SELESAI — {len(state['generated_files'])} dokumen, "
            f"{error_count} gagal.", "OK" if state["generation_done"] else "ERROR")
        _gen_ui_finish()

    btn_generate = ft.Button(
        "Mulai Generate Dokumen",
        icon=ft.Icons.PLAY_ARROW_ROUNDED,
        style=ft.ButtonStyle(
            color={
                ft.ControlState.DISABLED: ft.Colors.GREY_600,
                ft.ControlState.DEFAULT: ft.Colors.WHITE,
            },
            bgcolor={
                ft.ControlState.DISABLED: ft.Colors.GREY_300,
                ft.ControlState.DEFAULT: ft.Colors.GREEN_700,
            },
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.Padding.symmetric(horizontal=22, vertical=14),
        ),
        on_click=run_generation,
    )

    panel4 = ft.Column(
        [
            ft.Text("Generate Dokumen", size=20, weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_900),
            ft.Text("Langkah 4 dari 5 — data diisi ke template. Pantau prosesnya di panel log.",
                    size=13, color=ft.Colors.GREY_600),
            ft.Divider(height=1, color=ft.Colors.GREY_300),
            ft.Container(height=4),
            ft.Container(
                content=gen_summary_text, bgcolor=ft.Colors.GREY_50,
                border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
                border_radius=8, padding=12,
            ),
            bapp_t1_link_warning,
            evidence_layout_option,
            btn_generate,
            ft.Container(height=6),
            gen_progress,
            gen_status,
            gen_duration_box,
        ],
        spacing=10, scroll=ft.ScrollMode.AUTO, expand=True,
    )

    # ------------------------------------------------------------------ #
    # LANGKAH 5 - Simpan hasil                                            #
    # ------------------------------------------------------------------ #
    save_info_text = ft.Text("", size=13, color=ft.Colors.GREY_700)
    fmt_options_col = ft.Column(spacing=4)
    fmt_group = ft.RadioGroup(value="zip", content=fmt_options_col)
    save_result_area = ft.Column(visible=False, spacing=8)
    # Indikator proses simpan (ZIP/PDF): bar tak tentu + teks status
    save_progress = ft.ProgressBar(value=None, visible=False, bar_height=8,
                                   color=ft.Colors.BLUE_700,
                                   bgcolor=ft.Colors.BLUE_GREY_100)
    save_status = ft.Text("", size=13, color=ft.Colors.BLUE_900, visible=False)
    save_duration_box = ft.Container(visible=False)
    _save_timer_task = None
    _save_start_time = 0.0
    _conversion_start_time = 0.0
    _conversion_estimate_seconds = None

    # Dialog modal: ditampilkan saat konversi DOCX→PDF berlangsung
    _save_dialog_msg = ft.Text("Menyiapkan data…", size=14,
                               color=ft.Colors.BLUE_900,
                               text_align=ft.TextAlign.CENTER)
    _save_estimate_max = ft.Text("", size=13, visible=False,
                                 color=ft.Colors.BLUE_GREY_700,
                                 text_align=ft.TextAlign.CENTER)
    _save_estimate_remaining = ft.Text("", size=13, visible=False,
                                       color=ft.Colors.BLUE_900,
                                       weight=ft.FontWeight.W_600,
                                       text_align=ft.TextAlign.CENTER)
    save_preparing_dialog = ft.AlertDialog(
        modal=True,
        content=ft.Column(
            [
                ft.ProgressRing(width=48, height=48, stroke_width=4,
                                color=ft.Colors.BLUE_700),
                _save_dialog_msg,
                ft.Column(
                    [_save_estimate_max, _save_estimate_remaining],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=4,
                    tight=True,
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=16, tight=True,
        ),
        actions=[],
    )

    def refresh_step5():
        n = len(state["generated_files"])
        if n == 0:
            return
        doc = current_doc()
        is_bukti_terima = doc and doc.group == "Bukti Terima"
        docx_label = (f"Arsip ZIP {n} dokumen DOCX asli (cepat)"
                      if n > 1 else
                      "Arsip ZIP 1 dokumen DOCX asli (cepat)")
        options = [ft.Radio(value="docx_zip", label=docx_label)]
        if PDF_AVAILABLE:
            if is_bukti_terima:
                zip_label = "Arsip ZIP, 1 berkas PDF multi-halaman, (normal)"
            else:
                zip_label = (f"Arsip ZIP, {n} berkas PDF, (normal)" if n > 1
                             else "Arsip ZIP, 1 berkas PDF, (normal)")
            options.append(ft.Radio(value="zip", label=zip_label))
        if PDF_AVAILABLE and MERGE_AVAILABLE:
            if is_bukti_terima:
                merged_label = "PDF gabungan, 1 berkas multi-halaman (siap cetak, lebih lambat)"
            elif n > 1:
                merged_label = (f"PDF gabungan, 1 berkas multi-halaman (siap cetak, lebih lambat)")
            else:
                merged_label = "PDF gabungan, 1 berkas multi-halaman (siap cetak, lebih lambat)"
            options.append(ft.Radio(value="merged", label=merged_label))
        fmt_options_col.controls = options
        fmt_group.value = "zip" if PDF_AVAILABLE else "docx_zip"
        if is_bukti_terima:
            hint_zip = ("Pilih DOCX untuk menyimpan langsung tanpa konversi. "
                        "Dokumen multi-halaman berisi semua petugas dalam grid "
                        "2\u00d72 per halaman A4.")
        elif PDF_AVAILABLE and not MERGE_AVAILABLE:
            hint_zip = ("Pilih DOCX untuk menyimpan langsung, atau PDF per "
                        "petugas (PDF gabungan butuh paket 'pypdf').")
        elif PDF_AVAILABLE:
            hint_zip = ("Pilih DOCX tanpa konversi, PDF per petugas, atau "
                        "PDF gabungan untuk mencetak dari satu berkas.")
        else:
            hint_zip = ("DOCX disimpan langsung tanpa konversi "
                        "(LibreOffice bundel tidak ditemukan).")
        save_info_text.value = (
            f"{n} dokumen '{doc.label}' siap disimpan. "
            f"{hint_zip} Pilih format keluaran lalu klik Simpan."
        )
        page.update()

    def open_output_folder(e=None):
        if state["saved_path"]:
            open_in_explorer(state["saved_path"])

    def restart_workflow(e=None):
        nonlocal _gen_timer_task, _save_timer_task
        if _gen_timer_task and not _gen_timer_task.done():
            _gen_timer_task.cancel()
        if _save_timer_task and not _save_timer_task.done():
            _save_timer_task.cancel()
        state.update({
            "step": 0, "max_step": 0, "doc_id": None, "file_path": None,
            "dfs": {}, "data_ok": False, "errors": [],
            "template_path": None, "template_source": None,
            "generated_files": [], "generation_done": False,
            "gen_duration": None, "conv_duration": None,
            "saved_path": None, "saved": False,
            "evidence_image_layout": "grid",
            "evidence_image_orientation": "automatic", "busy": False,
        })
        restyle_doc_cards()
        step1_summary.value = "Belum ada dokumen dipilih."
        step1_summary.color = ft.Colors.GREY_600
        data_file_chip.value = "Belum ada file dipilih."
        data_file_chip.color = ft.Colors.GREY_600
        verify_area.visible = False
        tpl_status_chip.value = "Belum ada template dipilih."
        tpl_status_chip.color = ft.Colors.GREY_600
        tpl_preview_area.visible = False
        gen_progress.visible = False
        gen_progress.value = 0
        gen_status.visible = False
        gen_duration_box.visible = False
        bapp_t1_link_warning.visible = False
        evidence_layout_option.visible = False
        evidence_layout_group.value = "grid"
        evidence_orientation_group.value = "automatic"
        evidence_orientation_group.disabled = True
        evidence_orientation_option.visible = False
        gen_summary_text.value = "Lengkapi langkah sebelumnya."
        save_result_area.visible = False
        save_duration_box.visible = False
        save_progress.visible = False
        save_status.visible = False
        fmt_options_col.disabled = False
        try:
            page.pop_dialog()
        except Exception:
            pass
        log("── Sesi baru dimulai ──", "STEP")
        goto(0)

    async def _save_timer_loop():
        """Loop asinkron untuk memperbarui timer live pada dialog dan status langkah 5."""
        nonlocal _save_start_time
        try:
            while state["busy"]:
                elapsed = time.time() - _save_start_time
                clock_str = format_timer_clock(elapsed)
                btn_save.text = f"Menyimpan… ({clock_str})"
                # Update status teks jika ada base message
                base_msg = save_status.data or "Memproses penyimpanan…"
                save_status.value = f"{base_msg} ({clock_str})"
                # Update pesan dialog jika dialog sedang aktif
                dialog_base = _save_dialog_msg.data or "Menyiapkan data…"
                _save_dialog_msg.value = f"{dialog_base} ({clock_str})"
                if (_conversion_estimate_seconds is not None
                        and _save_estimate_max.visible):
                    conversion_elapsed = time.time() - _conversion_start_time
                    maximum, remaining = conversion_estimate_messages(
                        _conversion_estimate_seconds, conversion_elapsed)
                    _save_estimate_max.value = maximum
                    _save_estimate_remaining.value = remaining
                page.update()
                await asyncio.sleep(0.5)
        except asyncio.CancelledError:
            pass

    def save_ui_start(message: str = "Menyiapkan penyimpanan…"):
        """Tampilkan indikator loading dan kunci tombol selama menyimpan."""
        state["busy"] = True
        btn_save.disabled = True
        btn_save.text = "Menyimpan…"
        fmt_options_col.disabled = True
        save_status.data = message
        save_status.value = message
        save_progress.visible = True
        save_status.visible = True
        save_result_area.visible = False

    def save_ui_done():
        """Kembalikan tombol ke normal dan sembunyikan indikator proses."""
        nonlocal _save_timer_task
        if _save_timer_task and not _save_timer_task.done():
            _save_timer_task.cancel()
        state["busy"] = False
        btn_save.disabled = False
        btn_save.text = "Simpan Hasil"
        fmt_options_col.disabled = False
        save_progress.visible = False
        save_status.visible = False

    async def on_save_output(e):
        nonlocal _save_timer_task, _save_start_time
        nonlocal _conversion_start_time, _conversion_estimate_seconds
        if state["busy"] or not state["generated_files"]:
            return
        files = state["generated_files"]
        doc = current_doc()
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        fmt = fmt_group.value or "zip"

        suggested, ext_filter = save_dialog_options(fmt, doc.prefix, stamp)

        save_path = await save_out_picker.save_file(
            dialog_title="Simpan Hasil",
            file_name=suggested,
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=ext_filter,
        )
        if not save_path:
            return

        n_total = len(files)
        # ── Kunci UI & tampilkan dialog modal ──────────────────────
        state["busy"] = True
        state["conv_duration"] = None
        _save_start_time = time.time()
        save_duration_box.visible = False

        btn_save.disabled = True
        btn_save.text = "Menyimpan… (00:00)"
        fmt_options_col.disabled = True
        _save_dialog_msg.data = "Menyiapkan data…"
        _save_dialog_msg.value = "Menyiapkan data… (00:00)"
        _conversion_estimate_seconds = None
        _save_estimate_max.visible = False
        _save_estimate_remaining.visible = False
        page.show_dialog(save_preparing_dialog)
        if _save_timer_task and not _save_timer_task.done():
            _save_timer_task.cancel()
        _save_timer_task = asyncio.create_task(_save_timer_loop())
        page.update()
        try:
            if fmt == "merged" and not PDF_AVAILABLE:
                raise RuntimeError(
                    "Konversi PDF tidak tersedia "
                    "(LibreOffice bundel tidak ditemukan).")
            elif fmt == "merged" and not MERGE_AVAILABLE:
                raise RuntimeError(
                    "Penggabungan PDF tidak tersedia — "
                    "install dulu: pip install pypdf")
            elif fmt == "docx_zip":
                # Jalur cepat: kemas hasil generator apa adanya tanpa LibreOffice.
                page.pop_dialog()
                save_ui_start(
                    f"Mengemas {n_total} dokumen DOCX ke dalam ZIP…")
                page.update()
                log("Konversi PDF dilewati — ZIP diisi dokumen DOCX asli.", "INFO")
                final = ensure_extension(save_path, "zip")
                for arcname in await asyncio.to_thread(zip_files, files, final):
                    log(f"    + {arcname}")
                state["saved_path"] = final
                log(f"Hasil disimpan (ZIP/DOCX): {final} "
                    f"({len(files)} dokumen)", "OK")
            else:
                # ── Urutkan berkas berdasar no_urut_spk ─────────────────
                # Prefix numerik nama berkas = no_urut_spk / urutan baris.
                files_sorted = sorted(files, key=file_order_key)
                log("Urutan dokumen (mengikuti no_urut_spk):", "INFO")
                for i, f in enumerate(files_sorted, start=1):
                    log(f"    {i:02d}. {os.path.basename(f)}")

                # ── Konversi batch: satu proses LibreOffice untuk semua DOCX ──
                tmp_pdf_dir = tempfile.mkdtemp(prefix="gen_pdf_")
                conv_msg = (f"Mengonversi {n_total} dokumen ke dalam satu PDF...")
                _save_dialog_msg.data = conv_msg
                _save_dialog_msg.value = conv_msg
                recipient_count = None
                if doc.id == "bukti_terima":
                    recipient_rows = state["dfs"].get(bukti_terima.SHEET_NAME)
                    recipient_count = len(recipient_rows) if recipient_rows is not None else 0
                workload = conversion_workload(
                    doc.id, n_total, recipient_count=recipient_count)
                _conversion_estimate_seconds = estimate_conversion_seconds(
                    doc.id, workload)
                _conversion_start_time = time.time()
                maximum, remaining = conversion_estimate_messages(
                    _conversion_estimate_seconds, 0)
                _save_estimate_max.value = maximum
                _save_estimate_remaining.value = remaining
                _save_estimate_max.visible = True
                _save_estimate_remaining.visible = True
                log(f"Konversi PDF batch dimulai: {n_total} dokumen.", "INFO")
                page.update()
                pdfs, failed_paths = await asyncio.to_thread(
                    convert_docx_files_to_pdf, files_sorted, tmp_pdf_dir)
                gagal_konversi = [os.path.basename(path) for path in failed_paths]
                if gagal_konversi:
                    log(f"{len(gagal_konversi)} dokumen gagal dikonversi ke "
                        "PDF dan tidak ikut dalam hasil.", "WARN")
                if not pdfs:
                    raise RuntimeError(
                        "Tidak ada dokumen yang berhasil dikonversi ke PDF.")

                # ── Tutup dialog, lanjutkan dengan progress bar ─────────
                page.pop_dialog()
                _save_estimate_max.visible = False
                _save_estimate_remaining.visible = False

                if fmt == "merged":
                    # Semua PDF digabung jadi SATU berkas siap cetak,
                    # urutan halaman mengikuti no_urut_spk.
                    final = ensure_extension(save_path, "pdf")
                    save_ui_start(
                        f"Menggabungkan {len(pdfs)} PDF menjadi 1 berkas…")
                    page.update()
                    await asyncio.to_thread(merge_pdfs, pdfs, final)
                    state["saved_path"] = final
                    log(f"Hasil disimpan (PDF gabungan, {len(pdfs)} dokumen "
                        f"terurut no_urut_spk): {final}", "OK")
                else:
                    # Arsip ZIP: satu berkas PDF untuk tiap petugas
                    final = ensure_extension(save_path, "zip")
                    save_ui_start(
                        f"Mengemas {len(pdfs)} berkas PDF "
                        f"(satu per petugas) ke dalam ZIP…")
                    page.update()
                    for arcname in await asyncio.to_thread(
                            zip_files, pdfs, final):
                        log(f"    + {arcname}")
                    state["saved_path"] = final
                    log(f"Hasil disimpan (ZIP): {final} "
                        f"({len(pdfs)} berkas PDF)", "OK")

            # Hitung durasi proses konversi/penyimpanan
            conv_duration = time.time() - _save_start_time
            state["conv_duration"] = conv_duration
            if _save_timer_task and not _save_timer_task.done():
                _save_timer_task.cancel()
            size_kb = os.path.getsize(state["saved_path"]) / 1024
            save_result_area.controls = [
                ft.Container(
                    content=ft.Column([
                        ft.Row([
                            ft.Icon(ft.Icons.CHECK_CIRCLE_ROUNDED,
                                    color=ft.Colors.GREEN_700, size=28),
                            ft.Text("Selesai! Berhasil disimpan:",
                                    weight=ft.FontWeight.BOLD, size=15,
                                    color=ft.Colors.GREEN_900),
                        ], spacing=8),
                        ft.Text(state["saved_path"], size=12,
                                color=ft.Colors.GREY_800, selectable=True),
                        ft.Text(f"Ukuran: {size_kb:.1f} KB", size=12,
                                color=ft.Colors.GREY_600),
                        ft.Row([
                            ft.Button(
                                "Buka Folder",
                                icon=ft.Icons.FOLDER_OPEN_ROUNDED,
                                style=ft.ButtonStyle(
                                    color=ft.Colors.BLUE_800,
                                    bgcolor=ft.Colors.BLUE_50,
                                    side=ft.BorderSide(1, ft.Colors.BLUE_300),
                                    shape=ft.RoundedRectangleBorder(radius=8),
                                    padding=ft.Padding.symmetric(horizontal=15, vertical=12),
                                ),
                                on_click=open_output_folder,
                            ),
                            ft.Button(
                                "Mulai Ulang dari Awal",
                                icon=ft.Icons.REFRESH_ROUNDED,
                                style=ft.ButtonStyle(
                                    color=ft.Colors.WHITE,
                                    bgcolor=ft.Colors.GREEN_700,
                                    shape=ft.RoundedRectangleBorder(radius=8),
                                    padding=ft.Padding.symmetric(horizontal=18, vertical=12),
                                ),
                                on_click=restart_workflow,
                            ),
                        ], spacing=10),
                    ], spacing=6),
                    bgcolor=ft.Colors.GREEN_50,
                    border=ft.Border.all(1.5, ft.Colors.GREEN_300),
                    border_radius=10, padding=16,
                )
            ]
            save_result_area.visible = True

            # Tampilkan kotak lilac info durasi di Step 5
            doc_name = doc.label if doc else "Dokumen"
            dur_items = [
                ("Dokumen", f"{doc_name} ({n_total} berkas)"),
            ]
            if state.get("gen_duration") is not None:
                dur_items.append(("Waktu pembuatan (generate)", format_duration(state["gen_duration"])))
            duration_label = ("Waktu penyimpanan (tanpa konversi PDF)"
                              if fmt == "docx_zip" else
                              "Waktu penyimpanan (konversi PDF & simpan)")
            dur_items.append((duration_label, format_duration(conv_duration)))
            if state.get("gen_duration") is not None:
                total_duration = state["gen_duration"] + conv_duration
                dur_items.append(("Total waktu proses", format_duration(total_duration)))

            save_duration_box.content = duration_info_box(
                title="Ringkasan Waktu Proses",
                items=dur_items,
                icon=ft.Icons.TIMER_ROUNDED,
            )
            save_duration_box.visible = True

            state["saved"] = True      # penanda sesi langkah 5 tuntas
            update_nav()
            fmt_label = {"zip": "Konversi dan pengemasan PDF",
                         "docx_zip": "Pengemasan DOCX",
                         "merged": "Penggabungan PDF"}.get(fmt, "Penyimpanan")
            show_snackbar(f"{fmt_label} selesai.", ft.Colors.GREEN_700)
        except Exception as ex:
            log(f"Gagal menyimpan hasil: {ex}", "ERROR")
            show_snackbar(f"Gagal menyimpan: {ex}", ft.Colors.RED_700)
            save_duration_box.visible = False
        finally:
            # Pastikan dialog modal tertutup (siapapun path-nya)
            try:
                page.pop_dialog()
            except Exception:
                pass
            save_ui_done()
            page.update()

    btn_save = ft.Button(
        "Simpan Hasil",
        icon=ft.Icons.SAVE_ROUNDED,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE, bgcolor=ft.Colors.GREEN_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.Padding.symmetric(horizontal=22, vertical=14),
        ),
        on_click=on_save_output,
    )

    panel5 = ft.Column(
        [
            ft.Text("Simpan Hasil", size=20, weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_900),
            ft.Text("Langkah 5 dari 5 — simpan hasil generate ke komputer Anda.",
                    size=13, color=ft.Colors.GREY_600),
            ft.Divider(height=1, color=ft.Colors.GREY_300),
            ft.Container(height=4),
            save_info_text,
            fmt_group,
            btn_save,
            ft.Container(height=6),
            save_progress,
            save_status,
            save_duration_box,
            save_result_area,
        ],
        spacing=10, scroll=ft.ScrollMode.AUTO, expand=True,
    )

    # ------------------------------------------------------------------ #
    # Navigasi antar langkah                                              #
    # ------------------------------------------------------------------ #
    # Urutan wizard: Pilih Dokumen → Template Dokumen → Upload Data →
    # Generate → Simpan & Selesai.
    # (panel2 = layar upload data; panel3 = layar template dokumen.)
    panels = [panel1, panel3, panel2, panel4, panel5]
    for p in panels[1:]:
        p.visible = False

    back_btn = ft.Button(
        "Kembali",
        icon=ft.Icons.ARROW_BACK_ROUNDED,
        style=ft.ButtonStyle(
            color=ft.Colors.GREY_800, bgcolor=ft.Colors.WHITE,
            side=ft.BorderSide(1, ft.Colors.GREY_300),
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.Padding.symmetric(horizontal=15, vertical=12),
        ),
        on_click=lambda e: goto(state["step"] - 1),
    )
    next_btn = ft.Button(
        "Lanjut",
        icon=ft.Icons.ARROW_FORWARD_ROUNDED,
        style=ft.ButtonStyle(
            # Warna khusus saat disabled agar terlihat jelas abu-abu
            color={
                ft.ControlState.DEFAULT: ft.Colors.WHITE,
                ft.ControlState.DISABLED: ft.Colors.GREY_500,
            },
            bgcolor={
                ft.ControlState.DEFAULT: ft.Colors.BLUE_700,
                ft.ControlState.DISABLED: ft.Colors.GREY_200,
            },
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.Padding.symmetric(horizontal=22, vertical=12),
        ),
    )
    next_click_gate = DebounceGate(0.5)

    async def quit_app(e=None):
        """Keluar dari aplikasi (tutup jendela; fallback paksa destroy)."""
        log("Menutup aplikasi…", "STEP")
        page.update()
        await close_window(page)

    quit_btn = ft.Button(
        "Keluar",
        icon=ft.Icons.EXIT_TO_APP_ROUNDED,
        tooltip="Tutup aplikasi",
        style=ft.ButtonStyle(
            color=ft.Colors.RED_700,
            bgcolor=ft.Colors.WHITE,
            side=ft.BorderSide(1, ft.Colors.RED_200),
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.Padding.symmetric(horizontal=15, vertical=12),
        ),
        on_click=quit_app,
    )
    step_hint = ft.Text("Langkah 1 dari 5", size=12, color=ft.Colors.GREY_600,
                        expand=True, text_align=ft.TextAlign.END)

    # Alasan kenapa "Lanjut" masih abu-abu, per langkah (tooltip)
    NAV_HINTS = {
        0: "Pilih salah satu dokumen untuk mengaktifkan tombol ini",
        1: "Pilih atau unggah template dokumen terlebih dahulu",
        2: "Selesaikan unggah & verifikasi file Excel terlebih dahulu",
        3: "Jalankan Generate hingga selesai terlebih dahulu",
        4: "Ini langkah terakhir — tidak ada langkah berikutnya",
    }

    def update_nav():
        st = state["step"]
        doc = current_doc()
        back_btn.visible = st > 0
        next_btn.disabled = False
        if st == 0:
            # Ajust label next button bila dokumen tidak perlu template
            if doc and doc.no_template and state["doc_id"]:
                next_btn.text = "Lanjut ke Upload Data"
            else:
                next_btn.text = "Lanjut ke Template Dokumen"
            next_btn.icon = ft.Icons.ARROW_FORWARD_ROUNDED
            next_btn.disabled = state["doc_id"] is None
        elif st == 1:
            next_btn.text = "Lanjut ke Upload Data"
            next_btn.icon = ft.Icons.ARROW_FORWARD_ROUNDED
            next_btn.disabled = not state["template_path"]
        elif st == 2:
            next_btn.text = "Lanjut ke Generate"
            next_btn.icon = ft.Icons.ARROW_FORWARD_ROUNDED
            next_btn.disabled = not (state["data_ok"] and state["template_path"])
        elif st == 3:
            next_btn.text = "Lanjut ke Simpan Hasil"
            next_btn.icon = ft.Icons.ARROW_FORWARD_ROUNDED
            next_btn.disabled = not state["generation_done"]
        else:
            # Langkah terakhir: tidak ada langkah berikutnya — tombol
            # dinonaktifkan. Sesi baru tetap bisa dimulai lewat tombol
            # "Mulai Ulang dari Awal" pada panel hasil (langkah 5).
            next_btn.text = "Lanjut"
            next_btn.icon = ft.Icons.ARROW_FORWARD_ROUNDED
            next_btn.disabled = True
        # Saat dikunci: abu-abu (via style) + tooltip alasan
        next_btn.tooltip = NAV_HINTS.get(st) if next_btn.disabled else None
        step_hint.value = f"Langkah {st + 1} dari 5 — {STEP_DEFS[st]['label']}"

    async def on_next(e):
        if not next_click_gate.accept():
            return
        st = state["step"]
        doc = current_doc()
        # Jika dari Step 1 ke Step 2 dan dokumen tidak perlu template,
        # lewati Step 2 langsung ke Step 3 (Upload Data).
        if st == 0 and doc and doc.no_template:
            state["max_step"] = max(state["max_step"], 2)
            goto(2, force=True)
            return
        if st == 2:
            refresh_step4()
        if st == 3:
            refresh_step5()
        goto(st + 1, force=True)  # syarat langkah sudah dicek di update_nav()

    next_btn.on_click = on_next

    def goto(index: int, force: bool = False):
        if index < 0 or index >= len(panels):
            return
        # Gerbang navigasi: klik manual pada tracker tidak boleh melompat
        # melewati max_step. Tombol "Lanjut" memakai force=True karena
        # syaratnya masing-masing sudah digerbangi oleh update_nav().
        if index > state["max_step"] and not force:
            return
        state["step"] = index
        state["max_step"] = max(state["max_step"], index)
        for i, p in enumerate(panels):
            p.visible = i == index
        if index == 1:
            refresh_step3()
        if index == 3:
            refresh_step4()
        if index == 4:
            refresh_step5()
        update_nav()
        render_tracker()

    # ------------------------------------------------------------------ #
    # Perakitan halaman                                                   #
    # ------------------------------------------------------------------ #
    header = ft.Container(
        content=ft.Row(
            [
                ft.Container(
                    content=ft.Icon(ft.Icons.APPS_ROUNDED, color=ft.Colors.WHITE, size=24),
                    bgcolor=ft.Colors.BLUE_800, border_radius=8, padding=8,
                ),
                ft.Column(
                    [
                        ft.Text(APP_FULL_NAME,
                                size=16, weight=ft.FontWeight.BOLD,
                                color=ft.Colors.BLUE_900),
                        ft.Text("Panduan langkah demi langkah pembuatan dokumen",
                                size=11, color=ft.Colors.GREY_600),
                    ],
                    spacing=0,
                ),
                ft.Container(expand=True),
                ft.IconButton(
                    icon=ft.Icons.SYSTEM_UPDATE_ALT_ROUNDED,
                    icon_color=ft.Colors.BLUE_800,
                    tooltip="Periksa pembaruan",
                    on_click=on_check_updates,
                ),
                ft.IconButton(
                    icon=ft.Icons.INFO_OUTLINE_ROUNDED,
                    icon_color=ft.Colors.BLUE_800,
                    tooltip="Tentang SIOMAY",
                    on_click=show_about,
                ),
                ft.IconButton(
                    icon=ft.Icons.FEEDBACK_OUTLINED,
                    icon_color=ft.Colors.BLUE_800,
                    tooltip="Kirim feedback",
                    on_click=lambda _: open_external_url(FEEDBACK_URL, page),
                ),
                ft.Container(
                    content=ft.Text(f"{DISPLAY_VERSION} · {RELEASE_CHANNEL.title()}", size=11,
                                    weight=ft.FontWeight.W_600,
                                    color=ft.Colors.BLUE_800),
                    bgcolor=ft.Colors.BLUE_50, border_radius=20,
                    border=ft.Border.all(1, ft.Colors.BLUE_200),
                    padding=ft.Padding.symmetric(horizontal=12, vertical=5),
                ),
            ],
            spacing=10,
        ),
        bgcolor=ft.Colors.WHITE,
        padding=ft.Padding.symmetric(horizontal=24, vertical=12),
    )

    tracker_bar = ft.Container(
        content=tracker_box,
        bgcolor=ft.Colors.WHITE,
        border=ft.Border.only(bottom=ft.BorderSide(1, ft.Colors.GREY_200)),
        padding=ft.Padding.only(left=40, right=40, top=18, bottom=14),
    )

    content_area = ft.Container(
        content=ft.Column(panels, spacing=0, expand=True),
        expand=True,
        padding=ft.Padding.only(left=24, right=24, top=16, bottom=8),
    )

    nav_bar = ft.Container(
        content=ft.Row([back_btn, ft.Container(width=8), quit_btn,
                        ft.Container(expand=True), step_hint,
                        ft.Container(width=12), next_btn]),
        bgcolor=ft.Colors.WHITE,
        border=ft.Border.only(top=ft.BorderSide(1, ft.Colors.GREY_200)),
        padding=ft.Padding.symmetric(horizontal=24, vertical=12),
    )

    log_panel = ft.Container(
        content=ft.Column(
            [
                ft.Row(
                    [
                        ft.Icon(ft.Icons.TERMINAL_ROUNDED, size=16,
                                color=ft.Colors.GREEN_400),
                        ft.Text("Log Aktivitas", size=13,
                                weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                        ft.Container(expand=True),
                        ft.IconButton(
                            icon=ft.Icons.DELETE_SWEEP_OUTLINED,
                            icon_size=16, icon_color=ft.Colors.GREY_400,
                            tooltip="Bersihkan log",
                            on_click=clear_logs,
                        ),
                    ],
                    spacing=6,
                ),
                log_list,
            ],
            spacing=4,
            expand=True,
        ),
        width=360,
        bgcolor=ft.Colors.BLUE_GREY_900,
        border_radius=10,
        margin=ft.Margin.only(top=12, right=12, bottom=12),
        padding=ft.Padding.only(left=14, right=8, top=12, bottom=10),
    )

    body = ft.Row(
        [content_area, log_panel],
        spacing=0,
        expand=True,
        vertical_alignment=ft.CrossAxisAlignment.STRETCH,
    )

    page.add(header, tracker_bar, body, nav_bar)

    # Pesan awal
    log(f"Selamat datang di {APP_FULL_NAME} {DISPLAY_VERSION}.", "STEP")
    log("Langkah 1: pilih jenis dokumen yang ingin dibuat.", "INFO")
    if not PDF_AVAILABLE:
        log("Konversi PDF nonaktif (LibreOffice bundel tidak ditemukan).", "WARN")
    update_nav()
    render_tracker()
    if record_launch_and_should_prompt():
        show_feedback(resume_update_check=True)
    else:
        page.run_task(on_check_updates)


if __name__ == "__main__":
    ft.run(main)











