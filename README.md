# PowerShellGUI_to_SSIS_Export

# Contract Report Export Tool

This PowerShell tool provides a simple Windows Forms interface for generating contract-specific Excel reports from SQL Server. It is designed for automation + interactivity in clinical reporting workflows.

## 🔧 Features

- GUI for entering:
  - Contract filter
  - Health Plan (HP) filter for client segments with various coverage types
  - Contract exclusions
  - Export folder path
- Executes dynamic SQL query
- Transposes results and exports to `.xlsx`
- Leverages the [ImportExcel](https://github.com/dfinke/ImportExcel) module
- Optionally used to pass parameters to SSIS packages (see future release)

## 🗂️ File Structure
├── ContractExport_GUI.ps1 # Main PowerShell script with GUI

├── README.md # Project documentation

└── sql/ # Optional folder for raw queries

## 🚀 Requirements

- Windows with PowerShell 5.1+
- SQL Server access (`Integrated Security`)
- [ImportExcel module](https://github.com/dfinke/ImportExcel)

To install:
```powershell
Install-Module ImportExcel -Scope CurrentUser
