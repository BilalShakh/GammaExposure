import pandas as pd
import os
from typing import Dict, List, Tuple

# User-defined constant for input file
INPUT_FILE = "NQH6 - 2026-01-05.xls"

def find_table_boundaries(df: pd.DataFrame) -> List[Tuple[int, int]]:
    """
    Find the boundaries of all tables in the dataframe.
    Tables end with "TOTALS" or "No month data for this option type" in the first column.
    
    Returns:
        List of tuples containing (start_row, end_row) for each table
    """
    table_boundaries = []
    table_start = 0
    
    for idx, row in df.iterrows():
        first_col_value = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        
        if first_col_value == "TOTALS" or "No month data" in first_col_value:
            table_boundaries.append((table_start, idx))
            table_start = idx + 1
    
    # Handle case where last table doesn't end with TOTALS
    if table_start < len(df):
        table_boundaries.append((table_start, len(df) - 1))
    
    return table_boundaries

def extract_table(df: pd.DataFrame, start_row: int, end_row: int) -> pd.DataFrame:
    """Extract a single table from the dataframe using row indices."""
    return df.iloc[start_row:end_row + 1].reset_index(drop=True)

def find_header_row(df: pd.DataFrame, start_row: int, end_row: int) -> int:
    """
    Find the row containing actual column headers by looking for 'Month' column.
    Returns the row index of the header row.
    """
    for idx in range(start_row, end_row + 1):
        row_values = [str(val).strip().lower() if pd.notna(val) else "" for val in df.iloc[idx]]
        if any("month" == val for val in row_values):
            return idx
    return start_row  # Fallback to start row if no header found

def get_months_from_first_table(first_table: pd.DataFrame) -> List[str]:
    """
    Extract available months from the first table.
    Looks for the "month" column and returns unique values, excluding TOTALS.
    """
    # Find month column (case-insensitive)
    month_col = None
    for col in first_table.columns:
        if str(col).lower().strip() == "month":
            month_col = col
            break
    
    if month_col is None:
        raise ValueError(f"Could not find 'month' column in first table. Available columns: {first_table.columns.tolist()}")
    
    months = first_table[month_col].dropna().unique().tolist()
    # Filter out TOTALS rows
    months = [str(m).strip() for m in months if str(m).strip().upper() != "TOTALS"]
    return months

def parse_options_file(filepath: str) -> Tuple[List[str], pd.DataFrame]:
    """
    Parse the options data file and extract months and raw data.
    
    Returns:
        Tuple of (available_months, full_dataframe)
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Read the Excel file
    df = pd.read_excel(filepath, sheet_name=0, header=None)
    
    # Find table boundaries
    table_boundaries = find_table_boundaries(df)
    
    if not table_boundaries:
        raise ValueError("No tables found in the file")
    
    # Extract first table to get months
    first_table_start, first_table_end = table_boundaries[0]
    
    # Find the actual header row (the one containing "Month") within the first table
    header_row_idx = find_header_row(df, first_table_start, first_table_end)
    
    # Extract table data starting from the row AFTER the header to the end
    first_table = df.iloc[header_row_idx + 1:first_table_end + 1].copy()
    
    # Set column names from header row
    first_table.columns = df.iloc[header_row_idx]
    first_table = first_table.reset_index(drop=True)
    
    # Remove rows where Month column is NaN or contains non-data values
    month_col = None
    for col in first_table.columns:
        if str(col).lower().strip() == "month":
            month_col = col
            break
    
    if month_col:
        first_table = first_table.dropna(subset=[month_col])
    
    months = get_months_from_first_table(first_table)
    
    return months, df

def find_table_for_month(df: pd.DataFrame, month: str, table_type: str) -> Tuple[pd.DataFrame, bool]:
    """
    Find and extract a specific table for a month (Calls or Puts).
    
    Args:
        df: Full dataframe
        month: Month to search for (e.g., "MAR 26")
        table_type: Either "Calls" or "Puts"
    
    Returns:
        Tuple of (table_dataframe, found)
    """
    search_text = f"{month} {table_type}"
    
    for idx, row in df.iterrows():
        first_col_value = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        
        if first_col_value == search_text:
            # Found table title row (e.g., "MAR 26 Calls")
            # The next row should be headers, then data starts after that
            header_row_idx = idx + 1
            data_start_idx = idx + 2
            
            # Find end of table (TOTALS or no data)
            table_end = len(df) - 1
            for end_idx in range(data_start_idx, len(df)):
                end_col_value = str(df.iloc[end_idx, 0]).strip() if pd.notna(df.iloc[end_idx, 0]) else ""
                if end_col_value == "TOTALS" or "No month data" in end_col_value:
                    table_end = end_idx - 1
                    break
            
            # Extract data rows (excluding title and header)
            table = df.iloc[data_start_idx:table_end + 1].copy()
            
            # Set column names from the header row
            table.columns = df.iloc[header_row_idx].values
            table = table.reset_index(drop=True)
            
            # Remove rows that are all NaN or empty
            table = table.dropna(how='all')
            
            return table, True
    
    return pd.DataFrame(), False

def filter_table_columns(table: pd.DataFrame) -> pd.DataFrame:
    """
    Filter table to only keep Strike and At Close columns.
    """
    columns_to_keep = []
    
    # Find Strike column
    for col in table.columns:
        if str(col).lower().strip() == "strike":
            columns_to_keep.append(col)
            break
    
    # Find At Close column
    for col in table.columns:
        col_lower = str(col).lower().strip()
        if col_lower == "at close":
            columns_to_keep.append(col)
            break
    
    if columns_to_keep:
        return table[columns_to_keep].copy()
    else:
        return table  # Return original if columns not found

def main():
    """Main function to run the options parser."""
    # Check if input file exists
    if not os.path.exists(INPUT_FILE):
        print(f"Error: File '{INPUT_FILE}' not found.")
        print(f"Please ensure the file is in the current directory: {os.getcwd()}")
        return
    
    try:
        # Parse the file
        print(f"Reading file: {INPUT_FILE}\n")
        months, full_df = parse_options_file(INPUT_FILE)
        
        # Display available months
        print("Available Months:")
        for i, month in enumerate(months, 1):
            print(f"  {i}. {month}")
        
        # Get user selection
        while True:
            try:
                choice = int(input(f"\nSelect month (1-{len(months)}): "))
                if 1 <= choice <= len(months):
                    selected_month = months[choice - 1]
                    break
                else:
                    print(f"Please enter a number between 1 and {len(months)}")
            except ValueError:
                print("Invalid input. Please enter a number.")
        
        # Extract Calls table
        calls_table, calls_found = find_table_for_month(full_df, selected_month, "Calls")
        
        # Extract Puts table
        puts_table, puts_found = find_table_for_month(full_df, selected_month, "Puts")
        
        # Prepare data for CSV
        csv_data = []
        
        if calls_found:
            calls_filtered = filter_table_columns(calls_table)
            for idx, row in calls_filtered.iterrows():
                csv_data.append({
                    'OptionType': 'Call',
                    'Strike': row.iloc[0],
                    'OI': row.iloc[1]
                })
        
        if puts_found:
            puts_filtered = filter_table_columns(puts_table)
            for idx, row in puts_filtered.iterrows():
                csv_data.append({
                    'OptionType': 'Put',
                    'Strike': row.iloc[0],
                    'OI': row.iloc[1]
                })
        
        # Write to CSV
        if csv_data:
            csv_df = pd.DataFrame(csv_data)
            csv_df.to_csv('data.csv', index=False)
            print(f"Data saved to data.csv ({len(csv_data)} rows)")
            print(f"Selected month: {selected_month}")
        else:
            print(f"No data found for {selected_month}")
    
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
