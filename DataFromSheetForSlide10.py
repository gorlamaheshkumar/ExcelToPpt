import pandas as pd
import os

def print_data_from_sheet_pandas(excel_file_path, sheet_name):
    """
    This function reads an Excel sheet with multiple tables, finds specific tables
    by their title, and for each table, prints all its rows and the data from 
    the last four columns in a neatly formatted table.

    Args:
        excel_file_path (str): The full path to the Excel file.
        sheet_name (str): The name of the sheet to read from.
    """
    try:
        # --- Initial Setup ---
        # Read the Excel file and immediately fill any empty cells (NaN) with 0.
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None).fillna(0)
        df[0] = df[0].astype(str).str.strip()
        
        file_name = os.path.basename(excel_file_path)
        print(f"--- Reading data from sheet '{sheet_name}' in {file_name} ---")

        target_tables = [
            "Business Service - INC",
            "Business Service - RITM",
            "Business Service - Change Requests"
        ]

        # --- Process Each Table ---
        for table_title in target_tables:
            try:
                # Find the start and end rows for the current table block
                title_rows = df[df[0] == table_title]
                if title_rows.empty:
                    print(f"\n--- Could not find table title: '{table_title}' ---")
                    continue
                start_index = title_rows.index[0]

                search_area = df.iloc[start_index:]
                grand_total_rows = search_area[search_area[0] == 'Grand Total']
                if grand_total_rows.empty:
                    print(f"\n--- Could not find 'Grand Total' for table: '{table_title}' ---")
                    continue
                end_index = grand_total_rows.index[0]

                # --- Collect Data for Formatting ---
                # Find the actual header row within the table block by searching for 'Week'
                header_row_index = -1
                # We search from the title row down to the grand total row
                for idx in range(start_index, end_index):
                    row_data = df.iloc[idx]
                    # Check if any of the last 4 cells in the row contain the string 'Week'
                    if any('Week' in str(cell) for cell in row_data.iloc[-4:]):
                        header_row_index = idx
                        break
                
                if header_row_index == -1:
                    print(f"\n--- Could not find a header row (e.g., 'Week 27') for table: '{table_title}' ---")
                    continue

                header_row = df.iloc[header_row_index]
                week_labels = [str(label).strip() for label in header_row.iloc[-4:]]
                
                # Prepare data for printing
                header_for_print = ['Category'] + week_labels
                data_for_print = []

                # The actual data rows start on the line AFTER the header and go to the end of the block
                for i in range(header_row_index + 1, end_index + 1):
                    current_row = df.iloc[i]
                    row_title = str(current_row.iloc[0]).strip()
                    if not row_title:
                        continue
                    
                    # Convert data to integers to remove decimals from the '0.0' that fillna might create
                    data = [int(val) for val in current_row.iloc[-4:]]
                    data_for_print.append([row_title] + data)
                
                # --- Format and Print the Table ---
                print(f"\n--- Verifying Data for: {table_title} ---")
                
                # Calculate column widths for alignment
                all_rows = [header_for_print] + data_for_print
                col_widths = [max(len(str(item)) for item in col) for col in zip(*all_rows)]

                # Print Header
                header_line = "  ".join(header_for_print[i].ljust(col_widths[i]) for i in range(len(header_for_print)))
                print(header_line)
                print('-' * len(header_line))

                # Print Data Rows
                for row in data_for_print:
                    data_line = "  ".join(str(row[i]).ljust(col_widths[i]) for i in range(len(row)))
                    print(data_line)

            except Exception as find_error:
                print(f"\n--- An error occurred while processing table '{table_title}': {find_error} ---")

        print(f"\n--- Data extraction for sheet '{sheet_name}' complete ---")

    except FileNotFoundError:
        print(f"Error: The file was not found at '{excel_file_path}'")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    # --- Configuration ---
    EXCEL_FILE_PATH = 'C:/Users/2399586/VScode/ExcelToPpt/Files/ExcelData29.xlsb'
    SHEET_NAME = 'Business Services'

    # --- Execute the function ---
    print_data_from_sheet_pandas(EXCEL_FILE_PATH, SHEET_NAME)
