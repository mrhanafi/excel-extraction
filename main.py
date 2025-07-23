import pandas as pd

def load_excel(file_path,sheet_name=0,header=0):
    """Load Excel file into DataFrame
    sheet_name: Sheet name or index (0 for first sheet)
    header: Row to use as column names (0 for first row, None for no header)
    """
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header)

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

def get_single_cell(df, row, col):
    """Get a single cell value from DataFrame
    row: Row index (0-based)
    col: Column index (0-based)
    """
    try:
        return df.iloc[row, col]
    except IndexError:
        return None

def get_cell_by_excel_reference(df, cell_ref):
    """Get cell value using Excel reference like 'A1', 'B2'
    Example: get_cell_by_excel_reference(df, 'A1')
    """
    row, col = convert_cell_reference_to_indices(cell_ref)
    return get_single_cell(df, row, col)


if __name__ == "__main__":
    # Load Excel file
    file_path = "sample.xlsx"

    df = load_excel(file_path, sheet_name='Sheet1')