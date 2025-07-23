import pandas as pd
def load_excel(file_path,sheet_name=0,header=0):
    """Load Excel file into DataFrame
    sheet_name: Sheet name or index (0 for first sheet)
    header: Row to use as column names (0 for first row, None for no header)
    """
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header)

def get_single_cell(df, row, col):
    """Get a single cell value from DataFrame
    row: Row index (0-based)
    col: Column index (0-based)
    """
    try:
        return df.iloc[row, col]
    except IndexError:
        return None
    
def get_single_cell_by_name(df, row, column_name):
    """Get cell value using row index and column name
    Example: get_single_cell_by_name(df, 0, 'Name')
    """
    try:
        return df.loc[row, column_name]
    except (IndexError, KeyError):
        return None
    
def get_multiple_cells(df, cell_positions):
    """Get multiple cells by position
    cell_positions: List of (row, col) tuples
    Example: get_multiple_cells(df, [(0,0), (1,1), (2,2)])
    """
    results = {}
    for i, (row, col) in enumerate(cell_positions):
        try:
            results[f"R{row}C{col}"] = df.iloc[row, col]
        except IndexError:
            results[f"R{row}C{col}"] = None
    
    return results

def get_range(df, start_row, end_row, start_col, end_col):
    """Get a range of cells
    Example: get_range(df, 0, 2, 0, 2) gets A1:C3
    """
    try:
        return df.iloc[start_row:end_row+1, start_col:end_col+1]
    except IndexError:
        return pd.DataFrame()

def get_row(df, row_index):
    """Get entire row as Series
    Example: get_row(df, 0) gets first row
    """
    try:
        return df.iloc[row_index]
    except IndexError:
        return None

def get_column(df, column_name_or_index):
    """Get entire column
    Example: get_column(df, 'Name') or get_column(df, 0)
    """
    try:
        if isinstance(column_name_or_index, str):
            return df[column_name_or_index]
        else:
            return df.iloc[:, column_name_or_index]
    except (KeyError, IndexError):
        return None
    
def find_cells_with_text(df, search_text):
    """Find all cells containing specific text
    Returns list of (row, col, value) tuples
    """
    found_cells = []
    
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            cell_value = df.iloc[row_idx, col_idx]
            if pd.notna(cell_value) and str(search_text).lower() in str(cell_value).lower():
                found_cells.append((row_idx, col_idx, cell_value))
    
    return found_cells

def find_cells_by_condition(df, condition_func):
    """Find cells matching a condition
    condition_func: Function that returns True/False for each cell
    Example: find_cells_by_condition(df, lambda x: isinstance(x, (int, float)) and x > 100)
    """
    found_cells = []
    
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            cell_value = df.iloc[row_idx, col_idx]
            if condition_func(cell_value):
                found_cells.append((row_idx, col_idx, cell_value))
    
    return found_cells

def get_non_empty_cells(df):
    """Get all non-empty cells with their positions"""
    non_empty = []
    
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            cell_value = df.iloc[row_idx, col_idx]
            if pd.notna(cell_value) and str(cell_value).strip() != '':
                non_empty.append((row_idx, col_idx, cell_value))
    
    return non_empty

def filter_data(df, column_name, condition_value):
    """Filter rows based on column value
    Example: filter_data(df, 'Status', 'Active')
    """
    try:
        return df[df[column_name] == condition_value]
    except KeyError:
        return pd.DataFrame()

def get_summary_stats(df, numeric_only=True):
    """Get summary statistics for numeric columns"""
    if numeric_only:
        return df.describe()
    else:
        return df.describe(include='all')

def convert_cell_reference_to_indices(cell_ref):
    """Convert Excel cell reference (like 'A1') to (row, col) indices
    Example: 'A1' -> (0, 0), 'B2' -> (1, 1)
    """
    col_letters = ''
    row_numbers = ''
    
    for char in cell_ref:
        if char.isalpha():
            col_letters += char
        else:
            row_numbers += char
    
    # Convert column letters to number
    col_num = 0
    for char in col_letters:
        col_num = col_num * 26 + (ord(char.upper()) - ord('A') + 1)
    col_num -= 1  # Convert to 0-based
    
    row_num = int(row_numbers) - 1  # Convert to 0-based
    
    return (row_num, col_num)

def get_cell_by_excel_reference(df, cell_ref):
    """Get cell value using Excel reference like 'A1', 'B2'
    Example: get_cell_by_excel_reference(df, 'A1')
    """
    row, col = convert_cell_reference_to_indices(cell_ref)
    return get_single_cell(df, row, col)

# Simple usage examples
if __name__ == "__main__":
    # Load Excel file
    file_path = "sample.xlsx"
    
    # Example 1: Load entire sheet
    df = load_excel(file_path, sheet_name='Sheet1')
    print("Data shape:", df.shape)
    print("First 5 rows:")
    print(df.head())
    
    # Example 2: Get single cell (row 0, column 0 = A1)
    cell_value = get_single_cell(df, 0, 0)
    print(f"\nCell A1 (row 0, col 0): {cell_value}")
    
    # Example 3: Get cell by Excel reference
    cell_a1 = get_cell_by_excel_reference(df, 'A1')
    cell_b2 = get_cell_by_excel_reference(df, 'B2')
    print(f"Cell A1: {cell_a1}")
    print(f"Cell B2: {cell_b2}")
    
    # Example 4: Get multiple cells
    positions = [(0, 0), (1, 1), (2, 2)]  # A1, B2, C3
    multiple_values = get_multiple_cells(df, positions)
    print(f"\nMultiple cells: {multiple_values}")
    
    # Example 5: Get a range (A1:C3)
    range_data = get_range(df, 0, 2, 0, 2)
    print(f"\nRange A1:C3:")
    print(range_data)
    
    # Example 6: Get entire row
    first_row = get_row(df, 0)
    print(f"\nFirst row: {first_row.tolist()}")
    
    # Example 7: Get entire column
    if len(df.columns) > 0:
        first_col = get_column(df, 0)
        print(f"\nFirst column (first 5 values): {first_col.head().tolist()}")
    
    # Example 8: Find cells with specific text
    found_cells = find_cells_with_text(df, 'Total')
    print(f"\nCells containing 'Total': {found_cells}")
    
    # Example 9: Find numeric cells greater than 100
    numeric_cells = find_cells_by_condition(df, 
        lambda x: pd.notna(x) and isinstance(x, (int, float)) and x > 100)
    print(f"\nNumeric cells > 100: {numeric_cells}")
    
    # Example 10: Get all non-empty cells
    non_empty = get_non_empty_cells(df)
    print(f"\nTotal non-empty cells: {len(non_empty)}")
    
    # Example 11: Load without headers (raw data)
    df_raw = load_excel(file_path, sheet_name='Sheet1', header=None)
    print(f"\nRaw data (no headers) shape: {df_raw.shape}")
    
    # Example 12: Load specific sheets
    all_sheets = pd.read_excel(file_path, sheet_name=None)  # Load all sheets
    print(f"\nAvailable sheets: {list(all_sheets.keys())}")