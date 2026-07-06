# Rotina 860 - TOTVS Winthor Automation

Automated SQL query execution and data export from TOTVS Winthor using UI automation (pyautogui) combined with Winthor's built-in Rotina 860 SQL query tool.

## Overview

This automation:
1. Opens TOTVS Winthor cloud application
2. Runs SQL queries through Winthor's Rotina 860 (SQL Console)
3. Exports results to Excel files
4. Merges data into a master spreadsheet

## Available Queries

- **bp**: Base Produtos (all products)
- **0**: Pedido Compras (purchase orders with pending quantity > 0)
- **f1, f2, f5**: Individual filial data (pending quantities by branch)
- **f**: Combined filial data (all three branches in one query)

## Files

- `Folder completa/860 auto.py` - Main automation script
- `Folder completa/merge_filiais.py` - Data merging script
- `Folder completa/*.png` - Screen detection images for pyautogui

## Requirements

See `Folder completa/requirements.txt`

- pyautogui
- pyperclip
- pandas
- openpyxl
- pygetwindow
- pywin32

## Usage

```bash
python "Folder completa/860 auto.py"
```

Then select which query to run:
- bp: Base Products
- 0: Purchase Orders
- f5, f2, f1: Individual branch data
- f: All branches combined (optimized single query)
