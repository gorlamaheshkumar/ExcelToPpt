import pandas as pd

# --- SCRIPT CONFIGURATION ---
EXCEL_FILE_PATH = 'C:/Users/Mahesh/VScode/ExcelToPpt/Files/ExcelData29.xlsb'
SHEET_NAME = 'Volumetric trends INC & RITM'

def print_weekly_table(title, row_labels, week_labels, data):
    print(f"\n{title}")
    print(f"{'':<8}" + "".join([f"{w:<10}" for w in week_labels]))
    for i, label in enumerate(row_labels):
        row = [str(data[w][i]) for w in week_labels]
        print(f"{label:<8}" + "".join([f"{v:<10}" for v in row]))

def main():
    # Load the sheet into a DataFrame
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
    # INC block
    inc_header_row = df.iloc[0]
    # Get all week labels in order
    all_weeks = [inc_header_row[col] for col in df.columns if "Week" in str(inc_header_row[col])]
    last_four_weeks = all_weeks[-4:]
    inc_week_cols = [col for col in df.columns if inc_header_row[col] in last_four_weeks]
    inc_rows = df.iloc[1:6].set_index(df.columns[0])
    inc_labels = inc_rows.index.tolist()
    inc_data = {inc_header_row[col]: inc_rows[col].tolist() for col in inc_week_cols}
    print_weekly_table("INC", inc_labels, last_four_weeks, inc_data)
    # RITM block
    ritm_header_row = df.iloc[8]
    all_weeks_ritm = [ritm_header_row[col] for col in df.columns if "Week" in str(ritm_header_row[col])]
    last_four_weeks_ritm = all_weeks_ritm[-4:]
    ritm_week_cols = [col for col in df.columns if ritm_header_row[col] in last_four_weeks_ritm]
    ritm_rows = df.iloc[9:14].set_index(df.columns[0])
    ritm_labels = ritm_rows.index.tolist()
    ritm_data = {ritm_header_row[col]: ritm_rows[col].tolist() for col in ritm_week_cols}
    print_weekly_table("RITM", ritm_labels, last_four_weeks_ritm, ritm_data)

if __name__ == "__main__":
    main()