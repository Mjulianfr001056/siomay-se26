import flet as ft
import pandas as pd
import os
import tempfile

from src.validator import validate_excel_file, analyze_nulls, EXPECTED_SCHEMA
from src.template_generator import generate_template

def main(page: ft.Page):
    page.title = "Excel Data Validator & Inspector"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window.width = 1100
    page.window.height = 850
    page.window.min_width = 800
    page.window.min_height = 650
    page.padding = 20
    page.scroll = ft.ScrollMode.AUTO

    # State
    current_dfs = {}
    current_sheet = "data_petugas"
    current_file_path = None

    # UI Components references
    selected_file_text = ft.Text("Belum ada file yang dipilih", italic=True, color=ft.Colors.GREY_600)
    warning_card = ft.Container(visible=False)
    validation_content = ft.Column(visible=False, spacing=15)
    
    # 5-Row bottom table container
    table_container = ft.Container(
        content=ft.Text("Upload file valid untuk melihat preview 5 baris data.", color=ft.Colors.GREY_500),
        border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
        border_radius=8,
        padding=10,
        bgcolor=ft.Colors.WHITE
    )

    # Health Check summary components
    stats_row = ft.Row(wrap=True, spacing=15)
    cols_check_list = ft.ListView(expand=1, spacing=8, height=180, padding=5)
    rows_check_list = ft.ListView(expand=1, spacing=8, height=180, padding=5)

    def show_snackbar(message: str, color: str = ft.Colors.BLUE_700):
        page.show_dialog(
            ft.SnackBar(
                content=ft.Text(message, color=ft.Colors.WHITE, weight=ft.FontWeight.W_500),
                bgcolor=color,
                duration=4000,
            )
        )

    def update_sheet_view(sheet_name: str):
        nonlocal current_sheet
        current_sheet = sheet_name
        
        if sheet_name not in current_dfs:
            return
            
        df = current_dfs[sheet_name]
        analysis = analyze_nulls(df)
        
        # 1. Update Stats Cards
        stats_row.controls.clear()
        stats_row.controls.extend([
            ft.Container(
                content=ft.Column([
                    ft.Text("Total Baris", size=12, color=ft.Colors.GREY_700),
                    ft.Text(str(analysis['total_rows']), size=22, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.BLUE_50,
                border_radius=8,
                padding=12,
                width=160,
                border=ft.Border.all(1, ft.Colors.BLUE_200)
            ),
            ft.Container(
                content=ft.Column([
                    ft.Text("Total Kolom", size=12, color=ft.Colors.GREY_700),
                    ft.Text(str(analysis['total_cols']), size=22, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.BLUE_50,
                border_radius=8,
                padding=12,
                width=160,
                border=ft.Border.all(1, ft.Colors.BLUE_200)
            ),
            ft.Container(
                content=ft.Column([
                    ft.Text("Baris Mengandung Null", size=12, color=ft.Colors.GREY_700),
                    ft.Text(
                        f"{analysis['rows_with_null']} / {analysis['total_rows']}",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.AMBER_900 if analysis['rows_with_null'] > 0 else ft.Colors.GREEN_700
                    ),
                ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.AMBER_50 if analysis['rows_with_null'] > 0 else ft.Colors.GREEN_50,
                border_radius=8,
                padding=12,
                width=200,
                border=ft.Border.all(1, ft.Colors.AMBER_200 if analysis['rows_with_null'] > 0 else ft.Colors.GREEN_200)
            ),
            ft.Container(
                content=ft.Column([
                    ft.Text("Kolom Mengandung Null", size=12, color=ft.Colors.GREY_700),
                    ft.Text(
                        f"{len(analysis['cols_with_null'])} / {analysis['total_cols']}",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.AMBER_900 if len(analysis['cols_with_null']) > 0 else ft.Colors.GREEN_700
                    ),
                ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.AMBER_50 if len(analysis['cols_with_null']) > 0 else ft.Colors.GREEN_50,
                border_radius=8,
                padding=12,
                width=200,
                border=ft.Border.all(1, ft.Colors.AMBER_200 if len(analysis['cols_with_null']) > 0 else ft.Colors.GREEN_200)
            ),
        ])

        # 2. Update Column Null Check List
        cols_check_list.controls.clear()
        for col_name, null_count in analysis['col_null_counts'].items():
            has_null = null_count > 0
            cols_check_list.controls.append(
                ft.Container(
                    content=ft.Row([
                        ft.Icon(
                            ft.Icons.WARNING_AMBER_ROUNDED if has_null else ft.Icons.CHECK_CIRCLE_ROUNDED,
                            color=ft.Colors.AMBER_700 if has_null else ft.Colors.GREEN_600,
                            size=18
                        ),
                        ft.Text(col_name, weight=ft.FontWeight.W_500, size=13, expand=1),
                        ft.Container(
                            content=ft.Text(
                                f"{null_count} null" if has_null else "Lengkap (0 null)",
                                size=11,
                                weight=ft.FontWeight.BOLD,
                                color=ft.Colors.AMBER_900 if has_null else ft.Colors.GREEN_800
                            ),
                            bgcolor=ft.Colors.AMBER_100 if has_null else ft.Colors.GREEN_100,
                            border_radius=12,
                            padding=ft.Padding.symmetric(horizontal=8, vertical=3)
                        )
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    bgcolor=ft.Colors.GREY_50,
                    border_radius=6,
                    padding=ft.Padding.symmetric(horizontal=10, vertical=6),
                    border=ft.Border.all(1, ft.Colors.GREY_200)
                )
            )

        # 3. Update Row Null Check List
        rows_check_list.controls.clear()
        if not analysis['row_null_status']:
            rows_check_list.controls.append(ft.Text("Tidak ada baris data.", italic=True, size=12))
        else:
            for row_info in analysis['row_null_status']:
                has_null = row_info['has_null']
                row_label = f"Baris #{row_info['excel_row']}"
                desc = f"Null di: {', '.join(row_info['null_cols'])}" if has_null else "Semua kolom terisi lengkap"
                
                rows_check_list.controls.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Icon(
                                ft.Icons.WARNING_AMBER_ROUNDED if has_null else ft.Icons.CHECK_CIRCLE_ROUNDED,
                                color=ft.Colors.AMBER_700 if has_null else ft.Colors.GREEN_600,
                                size=18
                            ),
                            ft.Text(row_label, weight=ft.FontWeight.BOLD, size=13, width=80),
                            ft.Text(desc, size=12, color=ft.Colors.GREY_700, expand=1, no_wrap=False),
                        ], alignment=ft.MainAxisAlignment.START),
                        bgcolor=ft.Colors.GREY_50,
                        border_radius=6,
                        padding=ft.Padding.symmetric(horizontal=10, vertical=6),
                        border=ft.Border.all(1, ft.Colors.GREY_200)
                    )
                )

        # 4. Update Bottom 5-Row Data Preview Table
        preview_df = df.tail(5) if len(df) >= 5 else df
        
        # Build DataColumns
        data_columns = [
            ft.DataColumn(
                ft.Text("No", weight=ft.FontWeight.BOLD, size=12)
            )
        ]
        for col in df.columns:
            data_columns.append(
                ft.DataColumn(
                    ft.Text(col, weight=ft.FontWeight.BOLD, size=12)
                )
            )
            
        # Build DataRows
        data_rows = []
        for idx, row in preview_df.iterrows():
            cells = [
                ft.DataCell(ft.Text(str(idx + 1), weight=ft.FontWeight.BOLD, size=11, color=ft.Colors.BLUE_900))
            ]
            for col in df.columns:
                val = str(row[col]) if pd.notna(row[col]) and str(row[col]).strip() != '' and str(row[col]).strip().lower() != 'nan' else '-'
                is_null = val == '-'
                cells.append(
                    ft.DataCell(
                        ft.Text(
                            val,
                            size=12,
                            color=ft.Colors.RED_700 if is_null else ft.Colors.BLACK87,
                            italic=is_null
                        )
                    )
                )
            data_rows.append(ft.DataRow(cells=cells))

        dt = ft.DataTable(
            columns=data_columns,
            rows=data_rows,
            heading_row_color=ft.Colors.BLUE_50,
            border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
            border_radius=6,
            horizontal_margin=10,
            column_spacing=20,
            heading_row_height=36,
            data_row_min_height=32,
            data_row_max_height=42
        )

        table_container.content = ft.Column([
            ft.Row([
                ft.Icon(ft.Icons.TABLE_CHART_OUTLINED, color=ft.Colors.BLUE_800, size=18),
                ft.Text(f"Preview Data 5 Baris Terakhir ({sheet_name}) - Menampilkan {len(preview_df)} baris", weight=ft.FontWeight.BOLD, size=13, color=ft.Colors.BLUE_900),
            ]),
            ft.Divider(height=1, color=ft.Colors.GREY_300),
            ft.Row(
                controls=[dt],
                scroll=ft.ScrollMode.ADAPTIVE
            )
        ], spacing=8)
        
        page.update()

    def handle_file_process(file_path: str):
        nonlocal current_dfs, current_file_path
        current_file_path = file_path
        selected_file_text.value = os.path.basename(file_path)
        selected_file_text.italic = False
        selected_file_text.color = ft.Colors.BLACK87
        
        is_valid, errors, dfs = validate_excel_file(file_path)
        
        if not is_valid:
            # Show warning message
            current_dfs = {}
            validation_content.visible = False
            
            error_items = [ft.Text(f"• {err}", size=13, color=ft.Colors.RED_900) for err in errors]
            warning_card.content = ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.ERROR_OUTLINE_ROUNDED, color=ft.Colors.RED_700, size=26),
                        ft.Text("Format File Tidak Sesuai!", weight=ft.FontWeight.BOLD, size=16, color=ft.Colors.RED_900)
                    ]),
                    ft.Text(
                        "File yang diunggah tidak memiliki struktur lembar kerja (sheet) atau nama kolom yang sesuai dengan format acuan.",
                        size=13,
                        color=ft.Colors.RED_800
                    ),
                    ft.Container(
                        content=ft.Column(error_items, spacing=4),
                        bgcolor=ft.Colors.WHITE,
                        border_radius=6,
                        padding=10,
                        border=ft.Border.all(1, ft.Colors.RED_200)
                    ),
                    ft.Text("Silakan download template resmi menggunakan tombol 'Download Template' di atas.", italic=True, size=12, color=ft.Colors.RED_700)
                ], spacing=8),
                bgcolor=ft.Colors.RED_50,
                border=ft.Border.all(1.5, ft.Colors.RED_300),
                border_radius=8,
                padding=15
            )
            warning_card.visible = True
            table_container.content = ft.Container(
                content=ft.Text("File tidak valid. Silakan upload file yang sesuai format.", color=ft.Colors.RED_400, italic=True),
                padding=15
            )
            show_snackbar("Format file Excel tidak sesuai!", ft.Colors.RED_700)
            page.update()
            return

        # Success - Load data
        warning_card.visible = False
        validation_content.visible = True
        current_dfs = dfs
        
        # Reset selector ke sheet pertama
        sheet_selector.value = "data_petugas"
        update_sheet_view("data_petugas")
        show_snackbar("File Excel berhasil divalidasi dan dimuat!", ft.Colors.GREEN_700)

    # File Pickers (Flet >= 0.80: awaitable API, registered via page.services)
    async def on_pick_file(e):
        files = await file_picker.pick_files(
            dialog_title="Pilih File Excel Database Administrasi",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xls"],
        )
        if files:
            handle_file_process(files[0].path)

    async def on_save_template(e):
        save_path = await template_save_picker.save_file(
            dialog_title="Simpan Template Excel",
            file_name="template_database_administrasi.xlsx",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
        )
        if save_path:
            if not save_path.endswith('.xlsx'):
                save_path += '.xlsx'
            try:
                generate_template(save_path)
                show_snackbar(f"Template berhasil disimpan ke: {os.path.basename(save_path)}", ft.Colors.GREEN_700)
            except Exception as ex:
                show_snackbar(f"Gagal menyimpan template: {str(ex)}", ft.Colors.RED_700)

    file_picker = ft.FilePicker()
    page.services.append(file_picker)
    template_save_picker = ft.FilePicker()
    page.services.append(template_save_picker)

    # Sheet selector dropdown (Tabs API berubah di Flet >= 0.80)
    sheet_selector = ft.Dropdown(
        label="Sheet Aktif",
        value="data_petugas",
        width=300,
        options=[
            ft.dropdown.Option(key="data_petugas", text="data_petugas"),
            ft.dropdown.Option(key="data_organik", text="data_organik"),
            ft.dropdown.Option(key="merged", text="merged"),
        ],
        on_select=lambda e: update_sheet_view(e.control.value),
    )

    # Construct the Validation Content layout
    validation_content.controls = [
        sheet_selector,
        stats_row,
        ft.Row([
            # Column Health Check Box
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.VIEW_COLUMN_ROUNDED, color=ft.Colors.BLUE_800, size=18),
                        ft.Text("Pemeriksaan Kelengkapan Kolom (Null Check)", weight=ft.FontWeight.BOLD, size=13, color=ft.Colors.BLUE_900),
                    ]),
                    ft.Divider(height=1, color=ft.Colors.GREY_300),
                    cols_check_list
                ], spacing=6),
                bgcolor=ft.Colors.WHITE,
                border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
                border_radius=8,
                padding=12,
                expand=1
            ),
            # Row Health Check Box
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.TABLE_ROWS_ROUNDED, color=ft.Colors.BLUE_800, size=18),
                        ft.Text("Pemeriksaan Kelengkapan Baris (Null Check)", weight=ft.FontWeight.BOLD, size=13, color=ft.Colors.BLUE_900),
                    ]),
                    ft.Divider(height=1, color=ft.Colors.GREY_300),
                    rows_check_list
                ], spacing=6),
                bgcolor=ft.Colors.WHITE,
                border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
                border_radius=8,
                padding=12,
                expand=1
            ),
        ], spacing=15),
    ]

    # Header section
    header = ft.Container(
        content=ft.Row([
            ft.Row([
                ft.Container(
                    content=ft.Icon(ft.Icons.ANALYTICS_ROUNDED, color=ft.Colors.WHITE, size=30),
                    bgcolor=ft.Colors.BLUE_800,
                    border_radius=10,
                    padding=10
                ),
                ft.Column([
                    ft.Text("Excel Data Inspector & Validator", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                    ft.Text("Validasi format & cek kelengkapan nilai data sensus/administrasi", size=12, color=ft.Colors.GREY_600),
                ], spacing=2)
            ]),
            ft.Button(
                "Download Template Excel",
                icon=ft.Icons.DOWNLOAD_ROUNDED,
                style=ft.ButtonStyle(
                    color=ft.Colors.BLUE_800,
                    bgcolor=ft.Colors.BLUE_50,
                    side=ft.BorderSide(width=1, color=ft.Colors.BLUE_300),
                    shape=ft.RoundedRectangleBorder(radius=8),
                    padding=ft.Padding.symmetric(horizontal=15, vertical=12)
                ),
                on_click=on_save_template
            )
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        padding=ft.Padding.only(bottom=10)
    )

    # File input section
    file_upload_card = ft.Container(
        content=ft.Row([
            ft.Button(
                "Pilih File Excel (.xlsx)",
                icon=ft.Icons.FOLDER_OPEN_ROUNDED,
                style=ft.ButtonStyle(
                    color=ft.Colors.WHITE,
                    bgcolor=ft.Colors.BLUE_800,
                    shape=ft.RoundedRectangleBorder(radius=8),
                    padding=ft.Padding.symmetric(horizontal=18, vertical=12)
                ),
                on_click=on_pick_file
            ),
            ft.VerticalDivider(width=1, color=ft.Colors.GREY_300),
            ft.Row([
                ft.Icon(ft.Icons.ATTACH_FILE_ROUNDED, color=ft.Colors.GREY_600, size=18),
                selected_file_text
            ], expand=1)
        ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.CENTER),
        bgcolor=ft.Colors.GREY_50,
        border=ft.Border.all(1, ft.Colors.BLUE_GREY_100),
        border_radius=8,
        padding=12
    )

    # Main page assembly
    page.add(
        header,
        file_upload_card,
        warning_card,
        validation_content,
        ft.Container(height=10),
        # Bottom preview table
        table_container
    )

if __name__ == "__main__":
    ft.run(main)
