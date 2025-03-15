# Document Generator

## Overview
Document Generator is a streamlined tool designed to create personalized documents from a Word template and spreadsheet data. With a user-friendly web interface built using Streamlit, the application extracts variables formatted as `{variable_name}` from your Word document, maps them to corresponding columns in your data file, and generates customized documents for each row. It supports multiple spreadsheet formats—including Excel, CSV, TSV, ODS, and more—making it ideal for bulk document creation.

## Features
- **Template Variables:** Use `{variable_name}` in your Word document to define placeholders.
- **Multiple Data Sources:** Supports Excel, CSV, TSV, ODS, and other common spreadsheet formats.
- **Smart Mapping:** Automatically matches template variables with spreadsheet column headers (with manual adjustments available).
- **Fast Processing:** Optimized for large datasets using caching and session state management.
- **User-Friendly Interface:** Intuitive web application built with Streamlit.
- **Flexible Output:** Choose any column for naming output files and specify a custom output directory.
- **Cross-Platform:** Run on any system with Python 3.7+; a Windows executable option may be provided separately.

## Requirements
### For Python Users
- **Python:** Version 3.7 or higher
- **Required Packages:**
  - streamlit>=1.22.0
  - pandas>=1.3.0
  - python-docx>=0.8.11
  - openpyxl>=3.0.9
  - odfpy>=1.4.1
  - xlrd>=2.0.1
  - pyxlsb>=1.0.9

### For Windows Users
- **Standalone Executable:** A Windows executable may be available on the Releases page (no Python installation required).

## Installation & Usage
### Option 1: Run with Python
1. **Clone the Repository:**
   ```bash
   git clone https://github.com/yourusername/document-generator.git
   cd document-generator
   ```
2. **Run the Setup Script:**
   The setup script will check your Python version, install any missing dependencies, create a `requirements.txt` if needed, and launch the application.
   ```bash
   python setup-script.py
   ```
   Alternatively, you can install dependencies manually:
   ```bash
   pip install -r requirements.txt
   streamlit run generate.py
   ```

### Option 2: Windows Executable
If available, download `DocumentGenerator.exe` from the Releases page and double-click to run the application without needing a Python installation.

## How to Use
1. **Prepare Your Template:**
   - Create a Word document (`.docx`).
   - Insert variables using the format `{variable_name}`.
   - *Example:*  
     `Dear {customer_name}, your order #{order_id} will arrive on {delivery_date}.`

2. **Prepare Your Data:**
   - Create a spreadsheet (Excel, CSV, TSV, ODS, etc.).
   - Ensure the column headers match your template variables.
   - Each row should contain data for one document.

3. **Generate Documents:**
   - Launch the application.
   - Upload your Word template and data file.
   - Map detected template variables to the corresponding spreadsheet columns (default matching is automatic, but you can adjust manually).
   - Select the column to use for output file names.
   - Specify the output directory path.
   - Click **"Generate Documents"** to create the files.

4. **Access Your Files:**
   - Generated documents are saved in the specified output directory.
   - Files are named based on the selected filename column.

## Project Structure
- **generate.py:** Main Streamlit application that handles file uploads, variable mapping, and document generation.
- **requirements.txt:** Lists all required Python packages.
- **setup-script.py:** Script that verifies dependencies, creates a requirements file if missing, and launches the application.

## Customization
- **Filename Column:** Choose any column from your data to determine output file names.
- **Output Directory:** Specify a custom path for saving generated documents.
- **Manual Variable Mapping:** Adjust the automatic mapping if your template variables and spreadsheet headers differ.

## Troubleshooting
- **No Variables Detected:** Ensure your template uses the correct `{variable_name}` format.
- **File Format Issues:** If you encounter errors, try converting your data file to `.xlsx` or `.csv` format.
- **Executable Issues (Windows):** If using a standalone executable, confirm that your antivirus software is not blocking the application.
- **General Errors:** The app will display error messages for issues like file conversion failures or missing dependencies.

## Contributing
Contributions are welcome! To contribute:
1. Fork the repository.
2. Create a feature branch:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. Commit your changes:
   ```bash
   git commit -m "Add some amazing feature"
   ```
4. Push to your branch:
   ```bash
   git push origin feature/your-feature-name
   ```
5. Open a Pull Request with a detailed description of your changes.

## License
This project is **not licensed**. No permissions are granted to reuse, modify, or distribute this project without explicit permission from the author.

## Contact
For questions or feedback, please contact:
- **Abdelfateh Elhadjnouna** – [abdelfateh.elhadjnouna@gmail.com](mailto:abdelfateh.elhadjnouna@gmail.com)
- **Project Repository:** [https://github.com/yourusername/document-generator](https://github.com/yourusername/document-generator)
