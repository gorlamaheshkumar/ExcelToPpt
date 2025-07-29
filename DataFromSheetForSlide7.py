import pandas as pd
import os

# =================================================================================
# HELPER FUNCTIONS FOR DATA EXTRACTION
# =================================================================================

def find_header_row(df: pd.DataFrame, title_text: str) -> int:
    """
    Finds the row index of a table's header by searching for a title
    that appears in the row below the header.
    """
    for index, row in df.iterrows():
        # Search the entire row for the title text
        for cell_value in row:
            if title_text.lower() in str(cell_value).lower():
                # The header row is assumed to be one row above the title
                return index - 1
    return -1

def find_data_row(df: pd.DataFrame, label_text: str) -> int:
    """Finds the row index for a specific data label in the first column."""
    for index, row in df.iterrows():
        if label_text.lower() in str(row.iloc[0]).lower():
            return index
    return -1

def extract_data_block(df: pd.DataFrame, start_label: str, num_rows: int):
    """
    Extracts a block of data starting from a specific label in the first column.
    This is used for tables like 'INC Created by' and 'INC Resolved by'.
    """
    start_row_idx = find_data_row(df, start_label)
    if start_row_idx == -1:
        print(f"\nWarning: Could not find data block starting with '{start_label}'")
        return None, None

    # The header is assumed to be the row directly above the starting data label
    header_row_idx = start_row_idx - 1
    header_row = df.iloc[header_row_idx]
    
    # FIX: Find all column identifiers whose header contains "Week"
    all_week_cols = [col for col, h in header_row.items() if "Week" in str(h)]
    # FIX: Get the last four of these column identifiers
    last_four_week_cols = all_week_cols[-4:]
    # FIX: Get the corresponding header text for those columns
    last_four_week_headers = [str(header_row[col]) for col in last_four_week_cols]

    data_end_row = start_row_idx + num_rows
    
    categories = df.iloc[start_row_idx : data_end_row, 0].values.tolist()
    
    data_dict = {
        header: df.loc[start_row_idx : data_end_row - 1, col].tolist()
        for header, col in zip(last_four_week_headers, last_four_week_cols)
    }
        
    return categories, data_dict


def extract_stats_chart_data(df: pd.DataFrame):
    """Extracts data for the 'Incidents weekly Stats' chart by finding row labels."""
    # The week headers are in the first row of the DataFrame
    inc_header_row = df.iloc[0]

    # FIX: Find all column identifiers whose header contains "Week"
    all_week_cols = [col for col, h in inc_header_row.items() if "Week" in str(h)]
    # FIX: Get the last four of these column identifiers
    last_four_week_cols = all_week_cols[-4:]
    # FIX: Get the corresponding header text for those columns
    last_four_week_headers = [str(inc_header_row[col]) for col in last_four_week_cols]

    # Find the specific rows for 'created' and 'resolved'
    created_row_idx = find_data_row(df, 'INCs created')
    resolved_row_idx = find_data_row(df, 'INCs resolved')

    if created_row_idx == -1 or resolved_row_idx == -1:
        print("Warning: Could not find 'INCs created' or 'INCs resolved' rows for Stats chart.")
        return None, None

    created_data = df.loc[created_row_idx, last_four_week_cols].tolist()
    resolved_data = df.loc[resolved_row_idx, last_four_week_cols].tolist()

    categories = ['INCs created', 'INCs resolved']
    data_dict = {
        week: [created, resolved] 
        for week, created, resolved in zip(last_four_week_headers, created_data, resolved_data)
    }
    
    return categories, data_dict

def print_chart_data(title: str, categories, data_dict):
    """Prints the extracted chart data to the console for verification."""
    if data_dict is None or categories is None:
        return
    
    print(f"\n--- Verifying Data for: {title} ---")
    
    series_labels = list(data_dict.keys())
    
    header = f"{'Category':<25}" + "".join([f"{label:<15}" for label in series_labels])
    print(header)
    print("-" * len(header))
    
    for i, category in enumerate(categories):
        row_str = f"{category:<25}"
        for week in series_labels:
            # Ensure index is within bounds
            if i < len(data_dict[week]):
                row_str += f"{str(data_dict[week][i]):<15}"
        print(row_str)

# =================================================================================
# SCRIPT EXECUTION
# =================================================================================

def main():
    """
    This script reads an Excel file and prints specific data tables to the terminal.
    """
    # --- SCRIPT CONFIGURATION ---
    EXCEL_FILE_PATH = 'C:/Users/Mahesh/VScode/ExcelToPpt/Files/ExcelData27.xlsx'
    SHEET_NAME = 'Created'

    print(f"--- Reading data from {os.path.basename(EXCEL_FILE_PATH)} ---")

    try:
        # Using default engine for .xlsx files
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
    except ImportError:
        print("🛑 ERROR: The 'openpyxl' library is required for .xlsx files. Please run: pip install openpyxl")
        return
    except FileNotFoundError:
        print(f"🛑 ERROR: The file was not found at the specified path:\n{EXCEL_FILE_PATH}")
        return
    except Exception as e:
        print(f"🛑 ERROR: Failed to read Excel file. {e}")
        return

    # --- Extract and print data for all three charts ---
    # The first data block starts with "Tools Created" and has 4 rows
    creation_cats, creation_data = extract_data_block(df, start_label="Tools Created", num_rows=4)
    print_chart_data("INC Created by", creation_cats, creation_data)

    # The second block starts with "Auto closed by Tools" and has 2 rows
    closed_by_cats, closed_by_data = extract_data_block(df, start_label="Auto closed by Tools", num_rows=2)
    print_chart_data("INC Resolved by", closed_by_cats, closed_by_data)

    # The stats data is extracted by its specific function
    stats_cats, stats_data = extract_stats_chart_data(df)
    print_chart_data("INC Stats", stats_cats, stats_data)
    
    print("\n--- Data extraction complete ---")


if __name__ == "__main__":
    main()
