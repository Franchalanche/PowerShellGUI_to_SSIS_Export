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

