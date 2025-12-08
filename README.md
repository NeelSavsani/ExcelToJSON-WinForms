# ExcelToJSON-WinForms
A lightweight Windows Forms utility that converts Excel files (.xls/.xlsx) into clean JSON format. Supports header/footer skipping, column mapping, and shows real-time progress with a modern UI.
Excel to JSON Converter – Windows Forms Application

A modern and easy-to-use WinForms desktop application for converting Excel files (.xlsx / .xls) into structured JSON.

This tool is designed for businesses and developers who work with product master data, price lists, stock files, and catalog sheets.
It includes intelligent Excel parsing, automatic header/footer skipping, and stable column mapping based on your layout.

✨ Features

Convert Excel data directly to JSON with one click

Supports .xlsx and .xls formats

Skips header rows & footer totals automatically

Custom column-to-JSON mapping

Shows a progress dialog during long operations

Clean, centered UI with modern styling

Error handling for invalid files

Auto-generated JSON preview

Produces UTF-8 formatted JSON with proper indentation

🛠️ Technologies Used

C# (.NET Framework / WinForms)

ExcelDataReader

Newtonsoft.Json

Visual Studio

Output Example:
[
  {
    "Code": "ABC123",
    "Name": "Paracetamol 500mg",
    "ShortName": "Paracetamol 500mg",
    "MRP": 45.5,
    "SalesRate": 42.0,
    "Stock": 50,
    "Scheme": "10+1",
    "ExpDate": "2026-10-10"
  }
]


🚀 Ideal For

Pharma distributors

Inventory data digitization

ERP import/export tools

Anyone converting Excel master lists into JSON
