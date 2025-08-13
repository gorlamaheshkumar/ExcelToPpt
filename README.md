# PowerPoint Report Automator Web App

This project is a **Flask web application** that automates the creation of detailed PowerPoint presentations from Excel data. It provides a simple web interface where users can drag and drop an Excel file, and the application generates a polished, data-driven report from a predefined template. This eliminates the manual effort of creating recurring reports, saving time and preventing errors.

The application has two deployment options:

1.  **Linux/Cloud (Recommended):** A portable version that runs on any server (like Azure App Service) and handles all data processing and chart creation.
2.  **Windows-Only:** A version for local development that includes an extra feature to apply advanced chart styling by directly controlling the PowerPoint application.

## 🚀 Key Features

  * **Web-Based Interface**: A user-friendly front end built with Flask, allowing users to generate reports from any web browser.
  * **Drag & Drop Upload**: Easily upload your Excel data file and an optional PowerPoint template.
  * **Real-Time Feedback**: The UI shows the live status of the report generation process, from "running" to "finished" or "failed".
  * **Dynamic Content Generation**: Creates and populates slides with tables, line charts, and bar charts using `python-pptx`.
  * **Template-Based**: Uses a source `.pptx` file as a template for consistent branding and layout.
  * **Direct Download**: Download the finished presentation directly from the browser once it's ready.
  * **(Windows Only) Advanced Styling**: Leverages `pywin32` to apply specific chart styles not available in `python-pptx`.

-----

## ⚙️ Prerequisites

  * Python 3.8+
  * **For the advanced styling feature:** A Windows operating system with Microsoft PowerPoint installed.

-----

## 🛠️ Installation & Setup

1.  **Clone the repository:**

    ```bash
    git clone https://your-repository-url.com/
    cd your-project-directory
    ```

2.  **Create and activate a virtual environment (recommended):**

    ```bash
    # Create the virtual environment
    python -m venv venv

    # Activate it (Windows)
    .\venv\Scripts\activate

    # Activate it (Linux/macOS)
    source venv/bin/activate
    ```

3.  **Install the required packages:**
    Create a file named `requirements.txt` with the content below. The packages are compatible with both Windows and Linux.

    ```
    flask
    pandas
    python-pptx
    openpyxl
    gunicorn
    pywin32; sys_platform == 'win32'
    ```

    Then, install them using pip:

    ```bash
    pip install -r requirements.txt
    ```

    *Note: `pywin32` will only be installed on Windows systems.*

-----

## ▶️ Usage

This application is designed to be run as a web server.

1.  **Ensure your files are in the correct folders:**

      * Your main Flask script (e.g., `App_Linux.py`) and your processing script (`Main_Linux.py`) should be in the `Linux/` subfolder.
      * Your `index.html` file must be in the `Templates/` folder.
      * Your default PowerPoint template should be in the `Files/` folder.

2.  **Run the Flask web server:**
    Navigate to the `Linux/` directory and run the application:

    ```bash
    cd Linux
    python App_Linux.py
    ```

3.  **Access the application:**
    Open your web browser and go to the URL provided in the terminal (usually `http://127.0.0.1:5000`).

4.  **Generate your report:**

      * Drag and drop your Excel file onto the upload area.
      * (Optional) Upload a custom PowerPoint template.
      * The process will start automatically. You can monitor the progress in the log viewer on the page.
      * When the status is "finished," click the "Download PPTX" button.

-----

## ☁️ Deployment to the Cloud (Azure)

This application is designed to be deployed to a **Linux App Service**.

1.  **Prepare your code:** Make sure you are using the Linux-compatible versions of your scripts (`App_Linux.py`, `Main_Linux.py`) that do not use `pywin32`.
2.  **Set the Startup Command:** In your Azure App Service configuration, set the startup command. Since your script is in the `Linux/` subfolder, you must also create a blank `__init__.py` file inside it and use this command:
    ```bash
    gunicorn --bind=0.0.0.0 --timeout 600 Linux.App_Linux:app
    ```

-----

## 📜 License

This project is licensed under the MIT License.
