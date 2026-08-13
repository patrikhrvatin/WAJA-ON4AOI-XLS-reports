# WAJA-ON4AOI-XLS-reports

---

# JARL WAJA (Worked All Japan) Log Processor

This Python tool processes ADIF log files (typically exported from LoTW) to track your progress toward the **JARL WAJA (Worked All Japan) award**.

## Features

* **Automatic Sorting**: Tracks the first confirmed station for every band (160m to 6m) and prefecture (01-47).
* **Portable Handling**: Automatically identifies `/P` stations, marking them with an asterisk (`*`) in your reports and isolating them in a specific "Checklist" report for JARL verification.
* **Multiple Outputs**:
* **Band-Specific Reports**: Individual Excel files for each band (e.g., `20m_report.xlsx`).
* **Unified Summary**: `all_JA_per_bands.xlsx` containing all data.
* **WAJA Matrix**: `WAJA_Summary.xlsx` providing a clean grid view of your progress.


* **Visual Status**: Automatically color-codes your Excel sheets (Green for WORKED, Red for MISSING).

## Prerequisites

You will need Python installed on your computer along with the following libraries:

```bash
pip install pandas openpyxl

```

## How to use

1. **Prepare your Log**: Export your log from LoTW (or your logging software) as an `.adi` file and save it in the script folder as `lotwreport.adi`.
2. **Run the Script**:
```bash
python your_script_name.py

```


3. **Check Results**: The script will automatically generate the Excel reports in the same directory.

## File Descriptions

* **`WAJA_Summary.xlsx`**: Your master grid. Use this to see at a glance which band/prefecture combinations you are missing (marked by `.`).
* **`portable_stations_report.xlsx`**: A specialized list of all `/P` contacts. Use the **Note** column to track manual JARL verification status.
* **`all_JA_per_bands.xlsx`**: An aggregated file containing all your confirmed QSOs in a single list.

## Troubleshooting

* **Missing Data?**: Ensure your ADIF file contains the `STATE` or `CNTY` field, as this is how the script identifies the JARL prefecture code.
* **Encoding Errors**: The script uses `latin-1` to read ADIF files to ensure compatibility with most logging software exports. If you encounter encoding issues, ensure your exported ADIF file is saved correctly.

