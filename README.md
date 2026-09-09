# Spreadsheet Price Processor

A script that bulk-edits an Excel sheet and generates a chart from the data.

## What it does
Loads `sales.xlsx`, decreases the price column by 10% for every row, writes the adjusted price to a new column, and adds a pie chart summarizing the data.

## Tech
- Python
- openpyxl

## Run
```bash
pip install openpyxl
python main.py
```
