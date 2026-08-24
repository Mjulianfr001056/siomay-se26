import pandas as pd

EXPECTED_SCHEMA = {
    'data_petugas': [
        'nama_bps', 'nik', 'nama_lengkap', 'jabatan', 'wilayah_tugas', 
        'custNoRef', 'no_rekening', 'is_bri', 'id_bank', 'no_spk', 
        'no_bapp', 'no_sppl', 'bukti_bapp', 'tgl_penyelesaian_t1', 
        'min_jml_sls', 'jml_sls', 'listed_usaha_t1', 'target_usaha_t1'
    ],
    'data_organik': [
        'Nama', 'custNoRef', 'no_rekening', 'is_bri', 'id_bank'
    ],
    'merged': [
        'nama_bps', 'custNoRef', 'no_rekening', 'is_bri', 'id_bank', 'jabatan'
    ]
}

def validate_excel_file(file_path: str):
    """
    Validates if the provided excel file matches the pre-designated format.
    Returns (is_valid: bool, errors: list[str], dfs: dict[str, pd.DataFrame])
    """
    errors = []
    dfs = {}
    
    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        
        # Check required sheets
        missing_sheets = [s for s in EXPECTED_SCHEMA.keys() if s not in sheet_names]
        if missing_sheets:
            errors.append(f"Lembar kerja (Sheet) yang hilang: {', '.join(missing_sheets)}")
            return False, errors, {}
            
        for sheet, expected_cols in EXPECTED_SCHEMA.items():
            # Read sheet as string dtype to preserve formatting and prevent float conversion of IDs
            df = xl.parse(sheet, dtype=str)
            dfs[sheet] = df
            actual_cols = list(df.columns)
            
            missing_cols = [c for c in expected_cols if c not in actual_cols]
            if missing_cols:
                errors.append(f"Sheet '{sheet}' kehilangan kolom: {', '.join(missing_cols)}")
                
        if errors:
            return False, errors, dfs
            
        return True, [], dfs
        
    except Exception as e:
        return False, [f"Gagal membaca file Excel: {str(e)}"], {}

def is_empty_cell(x):
    """
    Checks if a single cell value is considered empty/null.
    """
    if pd.isna(x):
        return True
    s = str(x).strip().lower()
    return s in ['', 'nan', 'none', '<na>', 'nat']

def get_null_mask(df: pd.DataFrame):
    """
    Element-wise null mask check that is robust to empty strings, 'nan', etc.
    """
    if hasattr(df, 'map'):
        return df.map(is_empty_cell)
    else:
        return df.applymap(is_empty_cell)

def analyze_nulls(df: pd.DataFrame):
    """
    Analyzes null/missing values in a DataFrame.
    Returns detailed summary of null values.
    """
    if df is None or df.empty:
        return {
            'total_rows': 0,
            'total_cols': 0,
            'rows_with_null': 0,
            'col_null_counts': {},
            'cols_with_null': [],
            'row_null_status': []
        }
        
    total_rows = len(df)
    total_cols = len(df.columns)
    
    null_mask = get_null_mask(df)
    
    # Calculate nulls per column
    col_null_counts = null_mask.sum().to_dict()
    cols_with_null = [col for col, count in col_null_counts.items() if count > 0]
    
    # Calculate nulls per row
    row_has_null = null_mask.any(axis=1)
    rows_with_null = row_has_null.sum()
    
    # List of null fields per row (for detailed inspection)
    row_null_details = []
    for idx, row in null_mask.iterrows():
        null_cols = [col for col in df.columns if row[col]]
        # 1-based index for row numbering (Excel sheet row number is idx + 2 since headers are row 1)
        row_null_details.append({
            'row_idx': idx,
            'excel_row': idx + 2,
            'has_null': len(null_cols) > 0,
            'null_cols': null_cols
        })
        
    return {
        'total_rows': total_rows,
        'total_cols': total_cols,
        'rows_with_null': int(rows_with_null),
        'col_null_counts': col_null_counts,
        'cols_with_null': cols_with_null,
        'row_null_status': row_null_details
    }