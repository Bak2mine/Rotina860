import pandas as pd
import openpyxl
import os
import glob
import sys
import datetime
from openpyxl.styles import Border, Side

# Define standard thin border
thin = Side(style='thin')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def get_master_file(SAVE_FOLDER):
    path_only = os.path.splitdrive(SAVE_FOLDER)[1]
    for drive in ["V:", "C:"]:
        folder = drive + path_only
        matches = glob.glob(os.path.join(folder, "RELA*.xlsx"))
        if matches:
            print(f"  ✓ Master file found: {matches[0]}")
            return matches[0]
    print("  ✗ Master file not found on V: or C:")
    sys.exit(1)

def get_latest_file(SAVE_FOLDER, prefix):
    today = datetime.datetime.now().strftime("%Y%m%d")
    path_only = os.path.splitdrive(SAVE_FOLDER)[1]
    for drive in ["V:", "C:"]:
        folder = drive + path_only
        pattern = os.path.join(folder, f"{prefix}_{today}_*.xls")
        files = glob.glob(pattern)
        if files:
            latest = max(files, key=os.path.getmtime)
            print(f"  ✓ Found: {os.path.basename(latest)}")
            return latest
    print(f"  ✗ No file found for {prefix}")
    sys.exit(1)

# ============================================================
# RUN F — Filiais 1+2+5, Qtd. Pedida+Reservada
# ============================================================
def run(SAVE_FOLDER):
    MASTER_FILE = get_master_file(SAVE_FOLDER)
    wb = openpyxl.load_workbook(MASTER_FILE)

    # --- Filial 1+2+5 sheet ---
    SHEET_NAME = "Filial 1+2+5"
    if SHEET_NAME not in wb.sheetnames:
        print(f"  ✗ Sheet '{SHEET_NAME}' not found")
        print(f"  Available sheets: {wb.sheetnames}")
        sys.exit(1)

    print("Finding exported files...")
    f5_file = get_latest_file(SAVE_FOLDER, "QTD PEDIDA_RESERVADA filial5")
    f2_file = get_latest_file(SAVE_FOLDER, "QTD PEDIDA_RESERVADA filial2")
    f1_file = get_latest_file(SAVE_FOLDER, "QTD PEDIDA_RESERVADA filial1")

    print("Reading exported files...")
    df_f5 = pd.read_excel(f5_file)
    df_f2 = pd.read_excel(f2_file)
    df_f1 = pd.read_excel(f1_file)

    df_combined = pd.concat([df_f5, df_f2, df_f1], ignore_index=True)
    print(f"  ✓ Combined {len(df_combined)} rows total")

    ws = wb[SHEET_NAME]
    print("Clearing previous data from Filial 1+2+5...")
    ws.delete_rows(4, ws.max_row)

    print("Writing Filial 1+2+5 data...")
    for row_idx, row in enumerate(df_combined.itertuples(index=False), start=4):
        for col_idx, value in enumerate(row, start=2):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border

    # --- Qtd. Pedida+Reservada sheet ---
    QTD_SHEET = "Qtd. Pedida+Reservada"
    if QTD_SHEET not in wb.sheetnames:
        print(f"  ✗ Sheet '{QTD_SHEET}' not found")
        print(f"  Available sheets: {wb.sheetnames}")
        sys.exit(1)

    ws2 = wb[QTD_SHEET]
    print("Clearing previous data from Qtd. Pedida+Reservada...")
    ws2.delete_rows(5, ws2.max_row)

    df_pivot = df_combined.groupby('CODPROD', as_index=False)['QTPENDENTE'].sum()
    df_pivot = df_pivot.sort_values('CODPROD')

    print("Writing Qtd. Pedida+Reservada data...")
    for row_idx, row in enumerate(df_pivot.itertuples(index=False), start=5):
        ws2.cell(row=row_idx, column=3, value=row.CODPROD).border = border
        ws2.cell(row=row_idx, column=6, value=row.QTPENDENTE).border = border

    wb.save(MASTER_FILE)
    print(f"  ✓ Filial 1+2+5 updated with {len(df_combined)} rows")
    print(f"  ✓ Qtd. Pedida+Reservada updated with {len(df_pivot)} rows")
    print(f"\n✅ Done! Master file updated: {MASTER_FILE}")

# ============================================================
# RUN BP — Base Produtos
# ============================================================
def run_bp(SAVE_FOLDER):
    MASTER_FILE = get_master_file(SAVE_FOLDER)
    wb = openpyxl.load_workbook(MASTER_FILE)

    BASE_SHEET = "Base Produtos"
    if BASE_SHEET not in wb.sheetnames:
        print(f"  ✗ Sheet '{BASE_SHEET}' not found")
        print(f"  Available sheets: {wb.sheetnames}")
        sys.exit(1)

    print("Finding Base Produtos file...")
    base_file = get_latest_file(SAVE_FOLDER, "BASE PRODUTOS")
    df_base = pd.read_excel(base_file)

    ws4 = wb[BASE_SHEET]
   # Clear only data columns B-E from row 11 onwards, keep F/G/H formulas intact
    print("Clearing previous data from Base Produtos...")
    for row in ws4.iter_rows(min_row=11, max_row=ws4.max_row, min_col=2, max_col=5):
        for cell in row:
            cell.value = None

    print("Writing Base Produtos data...")
    for row_idx, row in enumerate(df_base.itertuples(index=False), start=11):
        ws4.cell(row=row_idx, column=2, value=row.CODPROD).border = border
        ws4.cell(row=row_idx, column=3, value=row.DESCRICAO).border = border
        ws4.cell(row=row_idx, column=4, value=row.CUSTOREP).border = border
        ws4.cell(row=row_idx, column=5, value=row.DTULTALTCUSTOREP).border = border

    ws4.column_dimensions['E'].width = 18

    wb.save(MASTER_FILE)
    print(f"  ✓ Base Produtos updated with {len(df_base)} rows")
    print(f"\n✅ Done! Master file updated: {MASTER_FILE}")

# ============================================================
# RUN PEDIDOS — Pedidos Compra
# ============================================================
def run_pedidos(SAVE_FOLDER):
    MASTER_FILE = get_master_file(SAVE_FOLDER)
    wb = openpyxl.load_workbook(MASTER_FILE)

    PEDIDOS_SHEET = "Pedidos Compra"
    if PEDIDOS_SHEET not in wb.sheetnames:
        print(f"  ✗ Sheet '{PEDIDOS_SHEET}' not found")
        print(f"  Available sheets: {wb.sheetnames}")
        sys.exit(1)

    print("Finding Pedidos Compra file...")
    pedidos_file = get_latest_file(SAVE_FOLDER, "PEDIDO COMPRAS")
    df_pedidos = pd.read_excel(pedidos_file)

    ws3 = wb[PEDIDOS_SHEET]
    print("Clearing previous data from Pedidos Compra...")
    ws3.delete_rows(5, ws3.max_row)

    print("Writing Pedidos Compra data...")
    for row_idx, row in enumerate(df_pedidos.itertuples(index=False), start=5):
        ws3.cell(row=row_idx, column=2, value=row.CODPROD).border = border
        ws3.cell(row=row_idx, column=3, value=row.DESCRICAO).border = border
        ws3.cell(row=row_idx, column=4, value=row.QTPEDIDA).border = border
        ws3.cell(row=row_idx, column=5, value=row.CODFILIAL).border = border
        ws3.cell(row=row_idx, column=6, value=row.DTULTPEDCOMPRA).border = border
        ws3.cell(row=row_idx, column=7, value=row.DTULTALTCOM).border = border

    wb.save(MASTER_FILE)
    print(f"  ✓ Pedidos Compra updated with {len(df_pedidos)} rows")
    print(f"\n✅ Done! Master file updated: {MASTER_FILE}")