import pandas as pd
import json
import csv

def convert_cell_reference_to_indices(cell_ref):
    """Convert Excel cell reference (like 'A1') to (row, col) indices"""
    col_letters = ''
    row_numbers = ''
    
    for char in cell_ref:
        if char.isalpha():
            col_letters += char
        else:
            row_numbers += char
    
    col_num = 0
    for char in col_letters:
        col_num = col_num * 26 + (ord(char.upper()) - ord('A') + 1)
    col_num -= 1
    
    row_num = int(row_numbers) - 1
    return (row_num, col_num)

def get_cell_by_excel_reference(df, cell_ref):
    """Get cell value using Excel reference like 'A1', 'B2'"""
    try:
        row, col = convert_cell_reference_to_indices(cell_ref)
        return df.iloc[row, col]
    except (IndexError, ValueError):
        return None

def extract_to_json_simple(file_path, sheet_name, cell_mapping):
    """Extract Excel data to JSON using cell references"""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    result = {}
    for key, cell_ref in cell_mapping.items():
        value = get_cell_by_excel_reference(df, cell_ref)
        result[key] = value
    
    return result

def json_to_csv_simple(json_data, csv_file_path):
    """Convert simple JSON key-value pairs to CSV
    
    Args:
        json_data: Dictionary or JSON string
        csv_file_path: Output CSV file path
    """
    if isinstance(json_data, str):
        data = json.loads(json_data)
    else:
        data = json_data
    
    # Create DataFrame with one row
    df = pd.DataFrame([data])
    df.to_csv(csv_file_path, index=False)
    print(f"CSV saved to: {csv_file_path}")

def json_to_csv_multiple_records(json_data_list, csv_file_path):
    """Convert multiple JSON records to CSV table
    
    Args:
        json_data_list: List of dictionaries
        csv_file_path: Output CSV file path
    """
    df = pd.DataFrame(json_data_list)
    df.to_csv(csv_file_path, index=False)
    print(f"CSV saved to: {csv_file_path}")

def json_to_csv_transposed(json_data, csv_file_path):
    """Convert JSON to transposed CSV (keys as rows, values as columns)
    
    Args:
        json_data: Dictionary or JSON string
        csv_file_path: Output CSV file path
    """
    if isinstance(json_data, str):
        data = json.loads(json_data)
    else:
        data = json_data
    
    # Create DataFrame with keys as index
    df = pd.DataFrame(list(data.items()), columns=['Field', 'Value'])
    df.to_csv(csv_file_path, index=False)
    print(f"Transposed CSV saved to: {csv_file_path}")

def process_multiple_excel_files_to_csv(file_configs, output_csv):
    """Process multiple Excel files and combine results into one CSV
    
    Args:
        file_configs: List of dictionaries with file info
                     Example: [
                         {'file': 'file1.xlsx', 'sheet': 'Sheet1', 'mapping': {...}},
                         {'file': 'file2.xlsx', 'sheet': 'Sheet1', 'mapping': {...}}
                     ]
        output_csv: Output CSV file path
    """
    all_records = []
    
    for config in file_configs:
        try:
            # Extract data from Excel
            json_data = extract_to_json_simple(
                config['file'], 
                config['sheet'], 
                config['mapping']
            )
            
            # Add source file info
            json_data['source_file'] = config['file']
            all_records.append(json_data)
            
        except Exception as e:
            print(f"Error processing {config['file']}: {e}")
    
    # Save all records to CSV
    if all_records:
        json_to_csv_multiple_records(all_records, output_csv)
    else:
        print("No records to save")

def excel_to_csv_pipeline(excel_file, sheet_name, cell_mapping, csv_output):
    """Complete pipeline: Excel -> JSON -> CSV
    
    Args:
        excel_file: Input Excel file path
        sheet_name: Sheet name
        cell_mapping: Dictionary mapping keys to cell references
        csv_output: Output CSV file path
    """
    print(f"Processing: {excel_file}")
    
    # Step 1: Extract from Excel to JSON
    json_data = extract_to_json_simple(excel_file, sheet_name, cell_mapping)
    print("Extracted JSON:", json.dumps(json_data, indent=2, default=str))
    
    # Step 2: Convert JSON to CSV
    json_to_csv_simple(json_data, csv_output)
    
    return json_data

def create_csv_with_headers(json_data_list, csv_file_path, custom_headers=None):
    """Create CSV with custom headers
    
    Args:
        json_data_list: List of dictionaries
        csv_file_path: Output CSV file path
        custom_headers: Custom column names (optional)
    """
    df = pd.DataFrame(json_data_list)
    
    if custom_headers:
        df.columns = custom_headers
    
    df.to_csv(csv_file_path, index=False)
    print(f"CSV with custom headers saved to: {csv_file_path}")

# Usage examples
if __name__ == "__main__":
    
    # Example 1: Single Excel file to CSV
    print("=== Example 1: Single file extraction ===")
    cell_mapping = {
        'customer_name': 'A1',
        'order_id': 'B1',
        'amount': 'C1',
        'date': 'D1'
    }
    
    json_result = excel_to_csv_pipeline(
        'order_data.xlsx', 
        'Sheet1', 
        cell_mapping, 
        'output_single.csv'
    )
    
    # Example 2: Multiple records from different Excel files
    print("\n=== Example 2: Multiple files to one CSV ===")
    file_configs = [
        {
            'file': 'order1.xlsx',
            'sheet': 'Sheet1',
            'mapping': {'customer': 'A1', 'amount': 'B1', 'date': 'C1'}
        },
        {
            'file': 'order2.xlsx', 
            'sheet': 'Sheet1',
            'mapping': {'customer': 'A1', 'amount': 'B1', 'date': 'C1'}
        },
        {
            'file': 'order3.xlsx',
            'sheet': 'Sheet1', 
            'mapping': {'customer': 'A1', 'amount': 'B1', 'date': 'C1'}
        }
    ]
    
    process_multiple_excel_files_to_csv(file_configs, 'combined_orders.csv')
    
    # Example 3: Create transposed CSV (fields as rows)
    print("\n=== Example 3: Transposed CSV ===")
    sample_json = {
        'customer_name': 'John Doe',
        'order_id': '12345', 
        'amount': 99.99,
        'status': 'Completed'
    }
    
    json_to_csv_transposed(sample_json, 'transposed_output.csv')
    
    # Example 4: Custom CSV with specific headers
    print("\n=== Example 4: Custom headers ===")
    multiple_records = [
        {'customer_name': 'John', 'amount': 100, 'date': '2024-01-01'},
        {'customer_name': 'Jane', 'amount': 200, 'date': '2024-01-02'},
        {'customer_name': 'Bob', 'amount': 150, 'date': '2024-01-03'}
    ]
    
    custom_headers = ['Customer', 'Order Amount', 'Order Date']
    create_csv_with_headers(multiple_records, 'custom_headers.csv', custom_headers)
    
    # Example 5: Manual JSON to CSV conversion
    print("\n=== Example 5: Manual conversion ===")
    
    # Simulate extracted JSON data
    extracted_data = [
        {'name': 'Product A', 'price': 25.99, 'stock': 100},
        {'name': 'Product B', 'price': 15.50, 'stock': 50},
        {'name': 'Product C', 'price': 35.00, 'stock': 75}
    ]
    
    json_to_csv_multiple_records(extracted_data, 'products.csv')
    
    # Example 6: Read back and verify CSV
    print("\n=== Example 6: Verify CSV output ===")
    try:
        df = pd.read_csv('output_single.csv')
        print("CSV contents:")
        print(df)
        print(f"Rows: {len(df)}, Columns: {len(df.columns)}")
    except FileNotFoundError:
        print("CSV file not found. Run the extraction first.")