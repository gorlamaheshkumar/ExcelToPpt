# PowerPoint Report Automator

This script automates the creation of a detailed PowerPoint presentation from a structured Excel workbook. It reads data from multiple sheets, generates tables and charts, and populates a predefined PowerPoint template, finishing by applying specific visual styles to the charts.

## 🚀 Key Features

  * **Data Extraction**: Parses data from specific sheets in an Excel file using `pandas`.
  * **Dynamic Content Generation**: Creates and populates slides with tables, line charts, and bar charts using `python-pptx`.
  * **Template-Based**: Uses a source `.pptx` file as a template for consistent branding and layout.
  * **Text Updates**: Automatically updates titles and conclusion text boxes with values derived from the data (e.g., current week number, totals).
  * **Advanced Styling**: Leverages `pywin32` to interact with the PowerPoint application directly, applying specific chart styles not available in `python-pptx`.

## ⚙️ Prerequisites

Before you begin, ensure you have the following installed:

  * Python 3.8+
  * Microsoft PowerPoint
  * **A Windows operating system** (required for the `pywin32` library to control PowerPoint).

## 🛠️ Installation

1.  **Clone the repository (or download the files)**

    ```bash
    git clone https://your-repository-url.com/
    cd your-project-directory
    ```

2.  **Create and activate a virtual environment (recommended)**

    ```bash
    # Create the virtual environment
    python -m venv venv

    # Activate it
    .\venv\Scripts\activate
    ```

3.  **Install the required packages**
    Create a file named `requirements.txt` in the project directory with the following content:

    ```
    pandas
    python-pptx
    pywin32
    openpyxl
    ```

    Then, install the packages using pip:

    ```bash
    pip install -r requirements.txt
    ```

-----

## 📄 Configuration

Before running the script, you must configure the file paths and sheet names inside the main script file (e.g., `main.py`).

1.  Place your source Excel file and template PowerPoint file in a known location.

2.  Open the script and modify the variables in the `if __name__ == "__main__":` block to match your file paths and sheet names.

    ```python
    if __name__ == "__main__":
        # --- MODIFY THESE PATHS ---
        EXCEL_FILE_PATH = 'C:/Path/To/Your/ExcelData.xlsx'
        TEMPLATE_PPTX_PATH = 'C:/Path/To/Your/Template.pptx'
        FINAL_OUTPUT_PPTX_PATH = 'C:/Path/To/Your/Output/Final_Report.pptx'

        # --- VERIFY SHEET NAMES (if different) ---
        SHEET_NAME_SLIDE_6 = 'Volumetric trends INC & RITM'
        SHEET_NAME_SLIDE_7 = 'Created'
        # ... and so on for other sheets
    ```

-----

## ▶️ Usage

Once the prerequisites are installed and the script is configured, run it from your terminal:

```bash
python main.py
```

The script will print its progress to the console and, upon completion, you will find the final presentation at the `FINAL_OUTPUT_PPTX_PATH` you specified.

## 📜 License

This project is licensed under the MIT License.