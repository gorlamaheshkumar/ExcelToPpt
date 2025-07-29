import pandas as pd
import os
from typing import List, Dict, Tuple

# =================================================================================
# HELPER FUNCTIONS FOR DATA EXTRACTION (NEW LOGIC)
# =================================================================================

def extract_pending_data_from_rows(df: pd.DataFrame, series_labels: List[str], week_column_label: str = 'Week #') -> Tuple[List[str], Dict[str, List]]:
    """
    Extracts the last four rows of data for a given set of series labels (column headers).
    """
    # Ensure all requested series labels exist as columns in the DataFrame
    valid_series_labels = [label for label in series_labels if label in df.columns]
    missing_labels = set(series_labels) - set(valid_series_labels)
    if missing_labels:
        print(f"Warning: The following columns were not found and will be skipped: {list(missing_labels)}")

    if not valid_series_labels:
        print("Warning: None of the requested data columns were found.")
        return None, None

    # Get the last four rows of data
    if len(df) < 4:
        print("Warning: Fewer than 4 rows of data available. Extracting all available rows.")
        last_four_rows_df = df.tail(len(df))
    else:
        last_four_rows_df = df.tail(4)

    # The categories for the chart are the week numbers/labels
    categories = last_four_rows_df[week_column_label].astype(str).tolist()

    # Create the data dictionary: {'Series Name': [val1, val2, val3, val4]}
    data_dict = {
        label: last_four_rows_df[label].tolist()
        for label in valid_series_labels
    }
    
    return categories, data_dict

def print_chart_data(title: str, categories: List[str], data_dict: Dict[str, List]):
    """
    Prints the extracted chart data to the console for verification.
    This version is adapted for a row-based data structure.
    """
    if data_dict is None or categories is None:
        return
    
    print(f"\n--- Verifying Data for: {title} ---")
    
    # Header row: 'Category' followed by the week numbers
    header = f"{'Category':<25}" + "".join([f"{cat:<15}" for cat in categories])
    print(header)
    print("-" * len(header))
    
    # Data rows: Each series name followed by its data for each week
    for series_name, series_data in data_dict.items():
        row_str = f"{series_name:<25}"
        row_str += "".join([f"{str(value):<15}" for value in series_data])
        print(row_str)

# =================================================================================
# SCRIPT EXECUTION
# =================================================================================

def main():
    """
    This script reads the 'Pending Counts' sheet and prints the data for Slide 8.
    """
    # --- SCRIPT CONFIGURATION ---
    EXCEL_FILE_PATH = 'C:/Users/Mahesh/VScode/ExcelToPpt/Files/ExcelData27.xlsx'
    SHEET_NAME = 'Pending Counts'

    print(f"--- Reading data from sheet '{SHEET_NAME}' in {os.path.basename(EXCEL_FILE_PATH)} ---")

    try:
        # Using default engine for .xlsx files
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
        # Clean up column names by stripping leading/trailing whitespace
        df.columns = df.columns.str.strip()
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

    # --- Define the series labels (column headers) for each chart ---
    inc_series_labels = ["Pending INCs", "INCs Resolved", "Total INCs Created"]
    ritm_series_labels = ["Pending RITMs", "RITMs Fulfilled", "Total RITMs Created"]

    # --- Extract and print data for both charts ---
    inc_cats, inc_data = extract_pending_data_from_rows(df, inc_series_labels)
    print_chart_data("INC Pending/Resolved/Created", inc_cats, inc_data)

    ritm_cats, ritm_data = extract_pending_data_from_rows(df, ritm_series_labels)
    print_chart_data("RITM Pending/Fulfilled/Created", ritm_cats, ritm_data)
    
    print("\n--- Data extraction for Slide 8 complete ---")


if __name__ == "__main__":
    main()
