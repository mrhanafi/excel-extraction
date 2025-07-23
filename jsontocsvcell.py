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

def json_to_csv_with_headers(json_data, csv_file_path):
    """Convert single JSON record to CSV with headers
    
    Args:
        json_data: Dictionary with key-value pairs
        csv_file_path: Output CSV file path
    
    Example:
        Input: {'name': 'John', 'age': 30, 'city': 'NYC'}
        Output CSV:
        name,age,city
        John,30,NYC
    """
    if isinstance(json_data, str):
        data = json.loads(json_data)
    else:
        data = json_data
    
    # Create DataFrame with one row
    df = pd.DataFrame([data])
    df.to_csv(csv_file_path, index=False)
    print(f"CSV with headers saved to: {csv_file_path}")
    return df

def multiple_json_to_csv_with_headers(json_data_list, csv_file_path):
    """Convert multiple JSON records to CSV with headers
    
    Args:
        json_data_list: List of dictionaries
        csv_file_path: Output CSV file path
    
    Example:
        Input: [
            {'name': 'John', 'age': 30, 'city': 'NYC'},
            {'name': 'Jane', 'age': 25, 'city': 'LA'}
        ]
        Output CSV:
        name,age,city
        John,30,NYC
        Jane,25,LA
    """
    if not json_data_list:
        print("No data to save")
        return pd.DataFrame()
    
    df = pd.DataFrame(json_data_list)
    df.to_csv(csv_file_path, index=False)
    print(f"CSV with headers saved to: {csv_file_path}")
    return df

def append_json_to_csv(json_data, csv_file_path):
    """Append JSON record to existing CSV file
    
    Args:
        json_data: Dictionary with key-value pairs
        csv_file_path: CSV file path (will be created if doesn't exist)
    """
    if isinstance(json_data, str):
        data = json.loads(json_data)
    else:
        data = json_data
    
    try:
        # Try to read existing CSV
        existing_df = pd.read_csv(csv_file_path)
        
        # Append new row
        new_df = pd.DataFrame([data])
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        
    except FileNotFoundError:
        # Create new CSV if file doesn't exist
        combined_df = pd.DataFrame([data])
    
    combined_df.to_csv(csv_file_path, index=False)
    print(f"Data appended to: {csv_file_path}")
    return combined_df

def process_multiple_excel_files_to_csv_with_headers(file_configs, output_csv):
    """Process multiple Excel files and create CSV with headers
    
    Args:
        file_configs: List of dictionaries with file configurations
        output_csv: Output CSV file path
    
    Example:
        file_configs = [
            {'file': 'order1.xlsx', 'sheet': 'Sheet1', 'mapping': {'customer': 'A1', 'amount': 'B1'}},
            {'file': 'order2.xlsx', 'sheet': 'Sheet1', 'mapping': {'customer': 'A1', 'amount': 'B1'}}
        ]
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
            
            # Add source file info (optional)
            if config.get('add_source', False):
                json_data['source_file'] = config['file']
            
            all_records.append(json_data)
            print(f"Processed: {config['file']}")
            
        except Exception as e:
            print(f"Error processing {config['file']}: {e}")
    
    # Save all records to CSV with headers
    if all_records:
        return multiple_json_to_csv_with_headers(all_records, output_csv)
    else:
        print("No records to save")
        return pd.DataFrame()

def excel_to_csv_with_headers_pipeline(excel_file, sheet_name, cell_mapping, csv_output):
    """Complete pipeline: Excel -> JSON -> CSV with headers
    
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
    
    # Step 2: Convert JSON to CSV with headers
    df = json_to_csv_with_headers(json_data, csv_output)
    
    return json_data, df

def custom_headers_csv(json_data_list, csv_file_path, header_mapping=None):
    """Create CSV with custom header names
    
    Args:
        json_data_list: List of dictionaries
        csv_file_path: Output CSV file path
        header_mapping: Dictionary to rename headers
                       Example: {'customer_name': 'Customer', 'order_date': 'Date'}
    """
    df = pd.DataFrame(json_data_list)
    
    if header_mapping:
        # Rename columns based on mapping
        df = df.rename(columns=header_mapping)
    
    df.to_csv(csv_file_path, index=False)
    print(f"CSV with custom headers saved to: {csv_file_path}")
    return df

def json_to_csv_with_column_order(json_data_list, csv_file_path, column_order=None):
    """Create CSV with specific column order
    
    Args:
        json_data_list: List of dictionaries
        csv_file_path: Output CSV file path
        column_order: List of column names in desired order
    """
    df = pd.DataFrame(json_data_list)
    
    if column_order:
        # Reorder columns, include any missing columns at the end
        available_cols = [col for col in column_order if col in df.columns]
        remaining_cols = [col for col in df.columns if col not in column_order]
        df = df[available_cols + remaining_cols]
    
    df.to_csv(csv_file_path, index=False)
    print(f"CSV with ordered columns saved to: {csv_file_path}")
    return df

# Usage examples
if __name__ == "__main__":
    
    # Example 1: Single JSON record to CSV with headers
    print("=== Example 1: Single record ===")
    single_record = {
        'customer_name': 'John Doe',
        'order_id': '12345',
        'amount': 99.99,
        'order_date': '2024-01-15',
        'status': 'Completed'
    }
    
    df1 = json_to_csv_with_headers(single_record, 'single_record.csv')
    print("Single record CSV:")
    print(df1)
    
    # Example 2: Multiple JSON records to CSV
    print("\n=== Example 2: Multiple records ===")
    multiple_records = [
        {'customer_name': 'John Doe', 'order_id': '12345', 'amount': 99.99, 'status': 'Completed'},
        {'customer_name': 'Jane Smith', 'order_id': '12346', 'amount': 149.50, 'status': 'Pending'},
        {'customer_name': 'Bob Johnson', 'order_id': '12347', 'amount': 75.25, 'status': 'Shipped'}
    ]
    
    df2 = multiple_json_to_csv_with_headers(multiple_records, 'multiple_records.csv')
    print("Multiple records CSV:")
    print(df2)
    
    # Example 3: Process multiple Excel files
    print("\n=== Example 3: Multiple Excel files ===")
    file_configs = [
        {
            'file': 'order1.xlsx',
            'sheet': 'Sheet1',
            'mapping': {'customer_name': 'A1', 'amount': 'B1', 'date': 'C1'},
            'add_source': True
        },
        {
            'file': 'order2.xlsx', 
            'sheet': 'Sheet1',
            'mapping': {'customer_name': 'A1', 'amount': 'B1', 'date': 'C1'},
            'add_source': True
        }
    ]
    
    # This would process multiple files
    # df3 = process_multiple_excel_files_to_csv_with_headers(file_configs, 'combined_orders.csv')
    
    # Example 4: Custom headers
    print("\n=== Example 4: Custom headers ===")
    header_mapping = {
        'customer_name': 'Customer',
        'order_id': 'Order Number',
        'amount': 'Total Amount',
        'status': 'Order Status'
    }
    
    df4 = custom_headers_csv(multiple_records, 'custom_headers.csv', header_mapping)
    print("Custom headers CSV:")
    print(df4)
    
    # Example 5: Specific column order
    print("\n=== Example 5: Ordered columns ===")
    column_order = ['order_id', 'customer_name', 'amount', 'status']
    
    df5 = json_to_csv_with_column_order(multiple_records, 'ordered_columns.csv', column_order)
    print("Ordered columns CSV:")
    print(df5)
    
    # Example 6: Append to existing CSV
    print("\n=== Example 6: Append new record ===")
    new_record = {
        'customer_name': 'Alice Brown',
        'order_id': '12348', 
        'amount': 200.00,
        'status': 'Processing'
    }
    
    df6 = append_json_to_csv(new_record, 'multiple_records.csv')
    print("After appending:")
    print(df6.tail())  # Show last few rows
    
    # Example 7: Complete pipeline
    print("\n=== Example 7: Complete pipeline ===")
    
    # Define extraction mapping
    extraction_mapping = {
        'customer': 'A1',
        'product': 'B1',
        'quantity': 'C1',
        'price': 'D1'
    }
    
    # This would run the complete pipeline
    # json_data, csv_df = excel_to_csv_with_headers_pipeline(
    #     'sales_data.xlsx', 
    #     'Sheet1', 
    #     extraction_mapping, 
    #     'sales_output.csv'
    # )
    
    print("\nAll examples completed!")
    print("\nGenerated CSV files:")
    print("- single_record.csv")
    print("- multiple_records.csv") 
    print("- custom_headers.csv")
    print("- ordered_columns.csv")