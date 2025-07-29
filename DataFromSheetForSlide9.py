import pandas as pd
import os
from typing import List, Dict, Tuple

# =================================================================================
# HELPER FUNCTIONS FOR DATA EXTRACTION
# =================================================================================

def find_data_row(df: pd.DataFrame, label_text: str) -> int:
    """Finds the row index for a specific data label in the first column."""
    for index, row in df.iterrows():
        # Check if the label text is present in the first cell of the row
        if label_text.lower() in str(row.iloc[0]).lower():
            return index
    return -1

def extract_change_data_block(df: pd.DataFrame, start_label: str, num_rows: int) -> Tuple[List[str], Dict[str, List]]:
    """
    Extracts a block of data for Slide 9, focusing on the last four data columns.
    """
    start_row_idx = find_data_row(df, start_label)
    if start_row_idx == -1:
        print(f"\nWarning: Could not find data block starting with '{start_label}'")
        return None, None

    # The header is the same row as the start label for this sheet's format
    header_row = df.iloc[start_row_idx]
    
    # Identify all potential data columns (all columns except the first one)
    all_data_cols = df.columns[1:]
    if len(all_data_cols) < 4:
        print(f"Warning: Fewer than 4 data columns found for '{start_label}'. Using all available.")
        last_four_data_cols = all_data_cols
    else:
        # Select the last four data columns
        last_four_data_cols = all_data_cols[-4:]
    
    # Get the headers for these last four columns from the header row
    last_four_week_headers = [str(header_row[col]) for col in last_four_data_cols]

    # The data starts from the row *after* the header/start_label row
    data_start_row = start_row_idx + 1
    data_end_row = data_start_row + num_rows
    
    # Get the category labels (e.g., "Automated Change", "Successful")
    categories = df.iloc[data_start_row : data_end_row, 0].values.tolist()
    
    # Extract the data for the last four weeks
    data_dict = {
        header: df.loc[data_start_row : data_end_row - 1, col].tolist()
        for header, col in zip(last_four_week_headers, last_four_data_cols)
    }
        
    return categories, data_dict

def print_chart_data(title: str, categories: List[str], data_dict: Dict[str, List]):
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
            if i < len(data_dict[week]):
                row_str += f"{str(data_dict[week][i]):<15}"
        print(row_str)

# =================================================================================
# SCRIPT EXECUTION
# =================================================================================

def main():
    """
    This script reads the 'Volumetric Change details' sheet and prints the data for Slide 9.
    """
    # --- SCRIPT CONFIGURATION ---
    EXCEL_FILE_PATH = 'C:/Users/2399586/VScode/ExcelToPpt/Files/ExcelData27.xlsx'
    SHEET_NAME = 'Volumetric Change details'

    print(f"--- Reading data from sheet '{SHEET_NAME}' in {os.path.basename(EXCEL_FILE_PATH)} ---")

    try:
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
    except ImportError:
        print("🛑 ERROR: The 'openpyxl' library is required for .xlsx files. Please run: pip install openpyxl")
        return
    except FileNotFoundError:
        print(f"🛑 ERROR: The file was not found at the specified path:\n{EXCEL_FILE_PATH}")
        return
    except ValueError as e:
         print(f"🛑 ERROR: Sheet '{SHEET_NAME}' not found in the Excel file. {e}")
         return
    except Exception as e:
        print(f"🛑 ERROR: Failed to read Excel file. {e}")
        return

    # --- Extract and print data for both tables ---
    type_cats, type_data = extract_change_data_block(df, start_label="Type", num_rows=6)
    print_chart_data("Change Request Type", type_cats, type_data)

    closure_cats, closure_data = extract_change_data_block(df, start_label="Closure Type", num_rows=4)
    print_chart_data("Change Closure Type", closure_cats, closure_data)
    
    print("\n--- Data extraction for Slide 9 complete ---")


if __name__ == "__main__":
    main()
