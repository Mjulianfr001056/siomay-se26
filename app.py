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

import flet as ft

from src import lampiran_spk
from src import bapp_pml
from src import bapp_ppl
from src import bapp_pml_t2
from src import bapp_ppl_t2
from src import spp
from src.validator import validate_excel_file, analyze_nulls
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
    fill_row,
    slugify,
)
from utils import (
    MERGE_AVAILABLE,
    PDF_AVAILABLE,
    close_window,
    convert_docx_to_pdf,
    ensure_extension,
    file_order_key,
    make_activity_log,
    make_snackbar,
    merge_pdfs,
    open_in_explorer,
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
    "SPP": ft.Icons.ASSIGNMENT_RETURNED_ROUNDED,
    "BAPP Termin 2": ft.Icons.FACT_CHECK_OUTLINED,
    "BAST": ft.Icons.HANDSHAKE_OUTLINED,
}


def main(page: ft.Page):
    page.title = "Generator Dokumen Administrasi SE2026"
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
        "saved_path": None,          # hasil step 5
        "saved": False,              # True setelah output berhasil disimpan
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

    # ------------------------------------------------------------------ #
    # Step tracker (atas, ala balenaEtcher)                               #
    # ------------------------------------------------------------------ #
    CIRCLE = 50

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
            content=ft.Text(STEP_DEFS[index]["label"], size=12, color=color,
                            weight=weight, text_align=ft.TextAlign.CENTER),
            width=CIRCLE + 40,
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
        doc = current_doc()
        builtin = doc.builtin_template_path
        restyle_doc_cards()
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
        if doc and doc.group == "Lampiran SPK":
            ok, errors, dfs = lampiran_spk.validate_input(path)
        elif doc and doc.group == "BAPP Termin 1":
            bapp_mod = bapp_ppl if doc.kind == "ppl" else bapp_pml
            ok, errors, dfs = bapp_mod.validate_input(path)
        elif doc and doc.group == "BAPP Termin 2":
            bapp_mod = bapp_ppl_t2 if doc.kind == "ppl" else bapp_pml_t2
            ok, errors, dfs = bapp_mod.validate_input(path)
        elif doc and doc.group == "SPP":
            ok, errors, dfs = spp.validate_input(path)
        else:
            ok, errors, dfs = validate_excel_file(path)
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
        if doc and doc.group == "Lampiran SPK":
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

        if doc and doc.group == "SPP":
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
            data_file_chip.value = f"{name} \u2014 SPP {role_label}"
            data_file_chip.color = ft.Colors.GREY_900
            log(f"Data diverifikasi: {name} \u2192 VALID "
                f"(PPL: {n_ppl}, PML: {n_pml})", "OK")
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
        state["template_path"] = files[0].path
        state["template_source"] = "unggahan"
        name = os.path.basename(files[0].path)
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
    bapp_t1_link_warning = ft.Container(
        visible=False,
        content=ft.Row(
            [
                ft.Icon(ft.Icons.WARNING_AMBER_ROUNDED,
                        color=ft.Colors.AMBER_800, size=20),
                ft.Text(
                    "PENTING: Pastikan tautan Google Drive screenshot "
                    "bukti dukung memiliki akses "
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
    btn_generate = None

    def refresh_step4():
        doc = current_doc()
        if doc is None or not state["data_ok"] or not state["template_path"]:
            return
        # Tampilkan peringatan akses link hanya untuk BAPP T1/T2 (PML/PPL)
        bapp_t1_link_warning.visible = (doc.group in ("BAPP Termin 1", "BAPP Termin 2"))
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
        elif doc.group == "SPP":
            df_mitra = state["dfs"].get(spp.SHEET_DATA_MITRA)
            n_total = len(df_mitra) if df_mitra is not None else 0
            extra = f" ({n_total} petugas total)"
        else:
            n_rows = len(state["dfs"].get("data_petugas", []))
            extra = f" ({n_rows} baris)"
        gen_summary_text.value = (
            f"Dokumen : {doc.label}\n"
            f"Sumber data : {os.path.basename(state['file_path'])}{extra}\n"
            f"Template : {os.path.basename(state['template_path'])}"
        )

    def _gen_ui_start():
        """Persiapan UI/state yang sama untuk semua jalur generate."""
        state["busy"] = True
        state["generated_files"] = []
        state["generation_done"] = False
        btn_generate.disabled = True
        gen_progress.visible = True
        gen_progress.value = 0
        gen_status.visible = True

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
            state["busy"] = False
            btn_generate.disabled = False
            update_nav()
            page.update()

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
            state["busy"] = False
            btn_generate.disabled = False
            update_nav()
            page.update()

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
            state["busy"] = False
            btn_generate.disabled = False
            update_nav()
            page.update()

    async def generate_bapp_pml_t2():
        """Populasi BAPP T2 PML — screenshot grid bukti dukung."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        out_dir = tempfile.mkdtemp(prefix="gen_bapp_pml_t2_")
        log("=" * 46, "STEP")
        log("MEMULAI GENERATE \u2014 BAPP T2 PML", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")

        try:
            for ev in bapp_pml_t2.iter_generate(
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
            state["busy"] = False
            btn_generate.disabled = False
            update_nav()
            page.update()

    async def generate_bapp_ppl_t2():
        """Populasi BAPP T2 PPL — screenshot grid bukti dukung."""
        _gen_ui_start()
        gen_status.value = "Menyiapkan populasi dokumen\u2026"
        page.update()

        out_dir = tempfile.mkdtemp(prefix="gen_bapp_ppl_t2_")
        log("=" * 46, "STEP")
        log("MEMULAI GENERATE \u2014 BAPP T2 PPL", "STEP")
        log(f"Template : {os.path.basename(state['template_path'])}", "INFO")
        log(f"Data     : {os.path.basename(state['file_path'])}", "INFO")

        try:
            for ev in bapp_ppl_t2.iter_generate(
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
            state["busy"] = False
            btn_generate.disabled = False
            update_nav()
            page.update()

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
            state["busy"] = False
            btn_generate.disabled = False
            update_nav()
            page.update()

    async def run_generation(e=None):
        if state["busy"]:
            return
        doc = current_doc()

        # ── Jalur khusus: grup Lampiran SPK (PPL / PML) ────────────────
        if doc and doc.group == "Lampiran SPK":
            await generate_lampiran_spk(
                doc, tempfile.mkdtemp(prefix="gen_spk_"))
            return

        # ── Jalur khusus: grup BAPP Termin 1 ────────────────
        if doc and doc.group == "BAPP Termin 1":
            if doc.kind == "ppl":
                await generate_bapp_ppl()
            else:
                await generate_bapp_pml()
            return

        # ── Jalur khusus: grup BAPP Termin 2 ────────────────
        if doc and doc.group == "BAPP Termin 2":
            if doc.kind == "ppl":
                await generate_bapp_ppl_t2()
            else:
                await generate_bapp_pml_t2()
            return

        # ── Jalur khusus: grup SPP ────────────────
        if doc and doc.group == "SPP":
            await generate_spp()
            return

        df = state["dfs"].get("data_petugas")
        if df is None or df.empty:
            log("Tidak ada baris data pada sheet 'data_petugas'.", "ERROR")
            show_snackbar("Data kosong!", ft.Colors.RED_700)
            return

        state["busy"] = True
        state["generated_files"] = []
        state["generation_done"] = False
        btn_generate.disabled = True
        gen_progress.visible = True
        gen_progress.value = 0
        gen_status.visible = True
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
        state["busy"] = False
        btn_generate.disabled = False
        gen_status.value = (
            f"Selesai: {len(state['generated_files'])}/{total} dokumen berhasil."
        )
        log(f"GENERATE SELESAI — {len(state['generated_files'])} dokumen, "
            f"{error_count} gagal.", "OK" if state["generation_done"] else "ERROR")
        update_nav()
        page.update()

    btn_generate = ft.Button(
        "Mulai Generate Dokumen",
        icon=ft.Icons.PLAY_ARROW_ROUNDED,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE, bgcolor=ft.Colors.GREEN_700,
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
            btn_generate,
            ft.Container(height=6),
            gen_progress,
            gen_status,
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

    # Dialog modal: ditampilkan saat konversi DOCX→PDF berlangsung
    _save_dialog_msg = ft.Text("Menyiapkan data…", size=14,
                               color=ft.Colors.BLUE_900,
                               text_align=ft.TextAlign.CENTER)
    save_preparing_dialog = ft.AlertDialog(
        modal=True,
        content=ft.Column(
            [
                ft.ProgressRing(width=48, height=48, stroke_width=4,
                                color=ft.Colors.BLUE_700),
                _save_dialog_msg,
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
        if PDF_AVAILABLE:
            # ZIP berisi hasil konversi PDF — satu berkas PDF per petugas
            zip_label = (f"Arsip ZIP ({n} berkas PDF)" if n > 1
                         else "Arsip ZIP (1 berkas PDF)")
        else:
            zip_label = (f"Arsip ZIP ({n} dokumen DOCX)"
                         " — konversi PDF tidak tersedia")
        options = [ft.Radio(value="zip", label=zip_label)]
        if PDF_AVAILABLE and MERGE_AVAILABLE:
            if n > 1:
                merged_label = (f"PDF gabungan ({n} dokumen menjadi 1 berkas, "
                                "urut no_urut_spk — siap cetak)")
            else:
                merged_label = "PDF gabungan (siap cetak)"
            options.append(ft.Radio(value="merged", label=merged_label))
        fmt_options_col.controls = options
        fmt_group.value = "zip"
        if PDF_AVAILABLE and not MERGE_AVAILABLE:
            hint_zip = ("ZIP akan berisi satu berkas PDF untuk setiap petugas "
                        "(opsi PDF gabungan butuh paket 'pypdf').")
        elif PDF_AVAILABLE:
            hint_zip = ("ZIP berisi satu PDF per petugas; pilih 'PDF gabungan' "
                        "untuk mencetak semuanya dari satu berkas.")
        else:
            hint_zip = ("ZIP berisi DOCX (pasang MS Word + pywin32 untuk PDF).")
        save_info_text.value = (
            f"{n} dokumen '{doc.label}' siap disimpan. "
            f"{hint_zip} Pilih format keluaran lalu klik Simpan."
        )
        page.update()

    def open_output_folder(e=None):
        if state["saved_path"]:
            open_in_explorer(state["saved_path"])

    def restart_workflow(e=None):
        state.update({
            "step": 0, "max_step": 0, "doc_id": None, "file_path": None,
            "dfs": {}, "data_ok": False, "errors": [],
            "template_path": None, "template_source": None,
            "generated_files": [], "generation_done": False,
            "saved_path": None, "saved": False, "busy": False,
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
        bapp_t1_link_warning.visible = False
        gen_summary_text.value = "Lengkapi langkah sebelumnya."
        save_result_area.visible = False
        save_progress.visible = False
        save_status.visible = False
        fmt_options_col.disabled = False
        try:
            page.pop_dialog()
        except Exception:
            pass
        log("── Sesi baru dimulai ──", "STEP")
        goto(0)

    def save_ui_start(message: str = "Menyiapkan penyimpanan…"):
        """Tampilkan indikator loading dan kunci tombol selama menyimpan."""
        state["busy"] = True
        btn_save.disabled = True
        btn_save.text = "Menyimpan…"
        fmt_options_col.disabled = True
        save_status.value = message
        save_progress.visible = True
        save_status.visible = True
        save_result_area.visible = False

    def save_ui_done():
        """Kembalikan tombol ke normal dan sembunyikan indikator proses."""
        state["busy"] = False
        btn_save.disabled = False
        btn_save.text = "Simpan Hasil"
        fmt_options_col.disabled = False
        save_progress.visible = False
        save_status.visible = False

    async def on_save_output(e):
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
        btn_save.disabled = True
        btn_save.text = "Menyimpan…"
        fmt_options_col.disabled = True
        _save_dialog_msg.value = "Menyiapkan data…"
        page.show_dialog(save_preparing_dialog)
        page.update()
        try:
            if fmt == "merged" and not PDF_AVAILABLE:
                raise RuntimeError(
                    "Konversi PDF tidak tersedia "
                    "(butuh MS Word terpasang + paket pywin32).")
            elif fmt == "merged" and not MERGE_AVAILABLE:
                raise RuntimeError(
                    "Penggabungan PDF tidak tersedia — "
                    "install dulu: pip install pypdf")
            elif fmt == "zip" and not PDF_AVAILABLE:
                # Tanpa MS Word/pywin32: fallback — isi ZIP dengan DOCX
                page.pop_dialog()
                save_ui_start(
                    f"Mengemas {n_total} dokumen DOCX ke dalam ZIP…")
                page.update()
                log("MS Word/pywin32 tidak tersedia — konversi PDF dilewati; "
                    "ZIP diisi dokumen DOCX.", "WARN")
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

                # ── Konversi setiap DOCX → satu berkas PDF per petugas ──
                tmp_pdf_dir = tempfile.mkdtemp(prefix="gen_pdf_")
                pdfs, gagal_konversi = [], []
                for idx, f in enumerate(files_sorted, start=1):
                    base = os.path.splitext(os.path.basename(f))[0]
                    pdf_out = os.path.join(tmp_pdf_dir, base + ".pdf")
                    _save_dialog_msg.value = (
                        f"Mengonversi ke PDF: dokumen {idx} dari "
                        f"{n_total}…")
                    log(f"Konversi PDF: {os.path.basename(f)}", "INFO")
                    page.update()
                    # Jalankan di worker thread agar UI tetap responsif;
                    # convert_docx_to_pdf menangani inisialisasi COM-nya.
                    try:
                        await asyncio.to_thread(
                            convert_docx_to_pdf, f, pdf_out)
                        pdfs.append(pdf_out)
                    except Exception as ex:
                        gagal_konversi.append(os.path.basename(f))
                        log(f"    GAGAL dikonversi: {ex}", "ERROR")
                if gagal_konversi:
                    log(f"{len(gagal_konversi)} dokumen gagal dikonversi ke "
                        "PDF dan tidak ikut dalam hasil.", "WARN")
                if not pdfs:
                    raise RuntimeError(
                        "Tidak ada dokumen yang berhasil dikonversi ke PDF.")

                # ── Tutup dialog, lanjutkan dengan progress bar ─────────
                page.pop_dialog()

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
            state["saved"] = True      # penanda sesi langkah 5 tuntas
            update_nav()
            fmt_label = {"zip": "Pengemasan ZIP",
                         "merged": "Penggabungan PDF"}.get(fmt, "Penyimpanan")
            show_snackbar(f"{fmt_label} selesai.", ft.Colors.GREEN_700)
        except Exception as ex:
            log(f"Gagal menyimpan hasil: {ex}", "ERROR")
            show_snackbar(f"Gagal menyimpan: {ex}", ft.Colors.RED_700)
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
        back_btn.visible = st > 0
        next_btn.disabled = False
        if st == 0:
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
        st = state["step"]
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
                        ft.Text("Generator Dokumen Administrasi SE2026",
                                size=16, weight=ft.FontWeight.BOLD,
                                color=ft.Colors.BLUE_900),
                        ft.Text("Panduan langkah demi langkah pembuatan dokumen",
                                size=11, color=ft.Colors.GREY_600),
                    ],
                    spacing=0,
                ),
                ft.Container(expand=True),
                ft.Container(
                    content=ft.Text("Workflow Wizard", size=11,
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
    log("Selamat datang di Generator Dokumen Administrasi SE2026.", "STEP")
    log("Langkah 1: pilih jenis dokumen yang ingin dibuat.", "INFO")
    if not PDF_AVAILABLE:
        log("Konversi PDF nonaktif (butuh MS Word terpasang + "
            "'pip install pywin32').", "WARN")
    update_nav()
    render_tracker()


if __name__ == "__main__":
    ft.run(main)











