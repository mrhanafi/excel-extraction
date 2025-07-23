import pandas as pd
import json

def convert_cell_reference_to_indices(cell_ref):
    """Convert Excel cell reference (like 'A1') to (row, col) indices"""
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
    """Get cell value using Excel reference like 'A1', 'B2'"""
    try:
        row, col = convert_cell_reference_to_indices(cell_ref)
        return df.iloc[row, col]
    except (IndexError, ValueError):
        return None

def extract_to_json_simple(file_path, sheet_name, cell_mapping):
    """Extract Excel data to JSON using cell references
    
    Args:
        file_path: Path to Excel file
        sheet_name: Name of the sheet
        cell_mapping: Dictionary with key names and cell references
                     Example: {'name': 'A1', 'age': 'B1', 'salary': 'C1'}
    
    Returns:
        JSON string
    """
    # Load Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Extract values
    result = {}
    for key, cell_ref in cell_mapping.items():
        value = get_cell_by_excel_reference(df, cell_ref)
        result[key] = value
    
    # Convert to JSON
    return json.dumps(result, indent=2, default=str)

def extract_multiple_records(file_path, sheet_name, key_column, value_column, start_row='A1', end_row=None):
    """Extract key-value pairs from two columns
    
    Args:
        file_path: Path to Excel file
        sheet_name: Name of the sheet
        key_column: Column letter for keys (e.g., 'A')
        value_column: Column letter for values (e.g., 'B')
        start_row: Starting row number (default: 1)
        end_row: Ending row number (None for auto-detect)
    
    Returns:
        JSON string
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    result = {}
    start_row_num = int(start_row[1:]) if start_row else 1
    
    # Auto-detect end row if not specified
    if end_row is None:
        end_row_num = len(df)
    else:
        end_row_num = int(end_row[1:])
    
    # Extract key-value pairs
    for row_num in range(start_row_num, end_row_num + 1):
        key_cell = f"{key_column}{row_num}"
        value_cell = f"{value_column}{row_num}"
        
        key = get_cell_by_excel_reference(df, key_cell)
        value = get_cell_by_excel_reference(df, value_cell)
        
        if key is not None and pd.notna(key):
            result[str(key)] = value
    
    return json.dumps(result, indent=2, default=str)

def extract_form_data(file_path, sheet_name, field_definitions):
    """Extract form-like data where labels and values are in specific cells
    
    Args:
        field_definitions: List of dictionaries with 'key', 'label_cell', 'value_cell'
                          Example: [
                              {'key': 'customer_name', 'label_cell': 'A1', 'value_cell': 'B1'},
                              {'key': 'order_date', 'label_cell': 'A2', 'value_cell': 'B2'}
                          ]
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    result = {}
    for field in field_definitions:
        key = field['key']
        label_cell = field.get('label_cell')
        value_cell = field['value_cell']
        
        value = get_cell_by_excel_reference(df, value_cell)
        
        # Optionally include the label
        if label_cell:
            label = get_cell_by_excel_reference(df, label_cell)
            result[key] = {
                'label': label,
                'value': value
            }
        else:
            result[key] = value
    
    return json.dumps(result, indent=2, default=str)

def extract_table_to_json(file_path, sheet_name, start_cell, end_cell):
    """Extract a table range and convert to JSON array
    
    Args:
        start_cell: Top-left cell (e.g., 'A1')
        end_cell: Bottom-right cell (e.g., 'C5')
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    start_row, start_col = convert_cell_reference_to_indices(start_cell)
    end_row, end_col = convert_cell_reference_to_indices(end_cell)
    
    # Extract the range
    table_data = []
    for row in range(start_row, end_row + 1):
        row_data = []
        for col in range(start_col, end_col + 1):
            try:
                value = df.iloc[row, col]
                row_data.append(value)
            except IndexError:
                row_data.append(None)
        table_data.append(row_data)
    
    return json.dumps(table_data, indent=2, default=str)

# Usage examples
if __name__ == "__main__":
    file_path = "sample.xlsx"
    
    # Example 1: Simple key-value extraction
    print("=== Example 1: Simple extraction ===")
    cell_mapping = {
        'customer_name': 'A1',
        'order_id': 'B1', 
        'total_amount': 'A2',
        'order_date': 'B2'
    }
    
    json_result = extract_to_json_simple(file_path, 'Sheet1', cell_mapping)
    print(json_result)
    
    # Example 2: Extract from two columns (A=keys, B=values)
    print("\n=== Example 2: Column-based extraction ===")
    json_result2 = extract_multiple_records(file_path, 'Sheet1', 'A', 'B', 'A1')
    print(json_result2)
    
    # Example 3: Form-style extraction with labels
    print("\n=== Example 3: Form data extraction ===")
    field_definitions = [
        {'key': 'customer_name', 'label_cell': 'A1', 'value_cell': 'B1'},
        {'key': 'phone', 'label_cell': 'A2', 'value_cell': 'B2'},
        {'key': 'email', 'label_cell': 'A3', 'value_cell': 'B3'},
        {'key': 'address', 'label_cell': 'A4', 'value_cell': 'B4'}
    ]
    
    json_result3 = extract_form_data(file_path, 'Sheet1', field_definitions)
    print(json_result3)
    
    # Example 4: Extract table range
    print("\n=== Example 4: Table extraction ===")
    json_result4 = extract_table_to_json(file_path, 'Sheet1', 'A1', 'C5')
    print(json_result4)
    
    # Example 5: Custom extraction with specific cells
    print("\n=== Example 5: Custom mixed extraction ===")
    df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
    
    custom_data = {
        'report_title': get_cell_by_excel_reference(df, 'A1'),
        'report_date': get_cell_by_excel_reference(df, 'A2'),
        'total_revenue': get_cell_by_excel_reference(df, 'B10'),
        'total_orders': get_cell_by_excel_reference(df, 'C10'),
        'notes': get_cell_by_excel_reference(df, 'A20')
    }
    
    print(json.dumps(custom_data, indent=2, default=str))