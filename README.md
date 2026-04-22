
# Mini‑Rupter Switch Resistance Test Log

A PyQt5 desktop application for recording Mini‑Rupter switch resistance test results and logging them automatically into a structured Excel workbook using openpyxl. Designed for production‑floor use with a clean UI, operator‑friendly workflows, and automated monthly Excel organization.

## Features
- Modern PyQt5 user interface
- Automatic date and time stamping
- Writes data directly to an Excel table
- Automatically creates monthly worksheets (e.g. April_2026)
- Excel tables with formatting, merged headers, and centered data
- Smart field memory (auto‑fills last entered values)
- Required field validation
- Branded header with logo support
- Preserves Shift, Pad, and Operator ID between entries

## Excel Output Format
Each month gets its own worksheet named `<Month>_<Year>` (example: `April_2026`).

Columns:
- Date and Time – Auto‑generated timestamp
- Cat# – Product Catalog Number
- JO# – Job Order Number
- Operator ID# – Operator Identifier
- Shift – AM / PM
- Type of Pad – Cu / Al
- A∅ – Phase A resistance (µΩ)
- B∅ – Phase B resistance (µΩ)
- C∅ – Phase C resistance (µΩ)

Sheets include fixed column widths, centered alignment, highlighted resistance headers (“Value Recorded / µΩ”), and a styled Excel table (TableStyleMedium2).

## Requirements
- Python 3.8+
- Windows OS
- Microsoft Excel (for viewing output)

Python dependencies:
PyQt5  
openpyxl  

Install with:
`pip install PyQt5 openpyxl`

## Running the Application
`python mini_rupter_app.py`

## Usage
1. Launch the application
2. Enter Cat#, JO#, and Operator ID#
3. Select Shift (AM / PM) and Pad type (Cu / Al)
4. Enter resistance values for A∅, B∅, and C∅
5. Click Submit

After submission, the date/time refreshes automatically, measurement fields are cleared, and repeating values are remembered.

## Configuration
Paths are configured directly in the source code:
`self.logo_path = r"...\logo.png"`  
`self.save_dir = r"...\Resistance Test Table"`

The Excel file `Resistance_Test_Table.xlsx` must already exist. The application does not create the file; it only creates and manages monthly worksheets within it.

## Build Executable (Optional)
Using PyInstaller:
`python -m PyInstaller -w -F mini_rupter_app.py --name MiniRupter_Resistance_Log --icon AppData\exe_logo.ico`

This produces a single executable suitable for production‑floor deployment.

## Error Handling
The application warns if required fields are missing, stops execution if the Excel file is unavailable, and displays clear user‑friendly error messages.

## Screenshots
Place screenshots in a `screenshots` folder if desired:
`screenshots/main_ui.png`  
`screenshots/excel_output.png`

## Author
Jonathan Huang, QA