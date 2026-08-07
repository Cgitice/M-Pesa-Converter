#!/usr/bin/env python3
"""
M-Pesa Statement Converter

Converts Safaricom M-Pesa PDF statements into a structured Excel workbook,
including transaction analysis and automatically generated pivot tables.

Features
--------
- Supports password-protected PDFs
- Extracts transaction history
- Cleans and formats transaction data
- Generates Excel reports with pivot tables
- Automatically removes temporary unlocked PDFs

Author: Caroline Mwende Gitice
Version: 1.0.0
"""

# Standard Library
import os
import sys
import re
import locale
from datetime import datetime

# Third-party Libraries
import numpy as np
import pandas as pd
import pdfplumber
import pikepdf
from tkinter import Tk, filedialog, simpledialog, messagebox

# ─────────────────────────────────────────────
# Locale (non-fatal if unsupported on platform)
# ─────────────────────────────────────────────
try:
    locale.setlocale(locale.LC_TIME, '')
except locale.Error:
    pass

# Application Constants
# =====================================================
APP_NAME = "M-Pesa Statement Converter"
VERSION = "1.0.0"
OUTPUT_FILENAME = "mpesa_statement.xlsx"
TEMP_UNLOCKED_PDF = "unlocked_mpesa.pdf"

# ─────────────────────────────────────────────
# PDF UNLOCK
# ─────────────────────────────────────────────
def unlock_pdf(input_path: str, output_path: str, password: str) -> str:
    """
    Unlocks a password-protected M-Pesa PDF and saves a temporary unlocked copy.
    """
    try:
        with pikepdf.open(input_path, password=password) as pdf:
            pdf.save(output_path)
        return output_path
    except Exception as e:
        messagebox.showerror("Error", f"Failed to unlock PDF: {e}")
        sys.exit(1)


# ─────────────────────────────────────────────
# TABLE EXTRACTION
# ─────────────────────────────────────────────
def extract_tables_with_pdfplumber(pdf_path: str) -> pd.DataFrame:
    dataframes = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table and len(table) > 1:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    dataframes.append(df)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract tables: {e}")
        sys.exit(1)

    if not dataframes:
        messagebox.showinfo("No Tables", "No tables found in the PDF.")
        sys.exit(0)

    combined = pd.concat(dataframes, ignore_index=True)

    # Strip whitespace from all string cells and normalize column names
    combined = combined.apply(lambda col: col.str.strip() if col.dtype == object else col)
    combined.columns = [c.strip() if isinstance(c, str) else c for c in combined.columns]
    return combined


# ─────────────────────────────────────────────
# DATE FORMATTING  (robust multi-format parser)
# ─────────────────────────────────────────────
_DATE_FORMATS = [
    "%d/%m/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%d-%m-%Y %H:%M:%S",
    "%d/%m/%Y",
    "%Y-%m-%d",
]


def _parse_date_series(series: pd.Series) -> pd.Series:
    for fmt in _DATE_FORMATS:
        parsed = pd.to_datetime(series, format=fmt, errors='coerce')
        if parsed.notna().sum() > 0:
            return parsed
    return pd.to_datetime(series, infer_datetime_format=True, errors='coerce')


def format_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Parse transaction dates and create a Month column
    used for pivot table aggregation.
    """
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    if 'Completion Time' in df.columns:
        parsed = _parse_date_series(df['Completion Time'])
        null_count = parsed.isna().sum()
        if null_count > 0:
            messagebox.showwarning(
                "Date Warning",
                f"{null_count} row(s) have unrecognised date formats and will be excluded from the pivot."
            )
        df['Month'] = parsed.dt.strftime('%Y-%m')
    else:
        messagebox.showwarning("Date Warning", "'Completion Time' column not found.")
    return df


# ─────────────────────────────────────────────
# NUMERIC NORMALISER
# ─────────────────────────────────────────────
def _normalise_numeric(series: pd.Series) -> pd.Series:
    """Strip commas, remove currency symbols, convert accounting parentheses to negatives."""
    return (
        series.astype(str)
        .str.replace(r'[^0-9\.\-\(\)]', '', regex=True)   # remove currency symbols and stray text
        .str.replace(',', '', regex=False)
        .str.replace(r'^\((.+)\)$', r'-\1', regex=True)   # (1234) → -1234
        .pipe(pd.to_numeric, errors='coerce')
    )


def _find_column_case_insensitive(df: pd.DataFrame, name: str):
    """Return actual column name in df that matches name case-insensitively, or None."""
    lname = name.strip().lower()
    for c in df.columns:
        if isinstance(c, str) and c.strip().lower() == lname:
            return c
    return None


def coerce_money_columns(df: pd.DataFrame, min_fraction: float = 0.5):
    """
    Detect columns that look like monetary values and convert them to numeric.
    Returns (df_converted, money_columns_list).
    """
    df = df.copy()
    money_cols = []

    # Try to detect money-like columns automatically
    for col in df.columns:
        if df[col].dtype == object:
            non_null_count = df[col].notna().sum()
            if non_null_count == 0:
                continue
            converted = _normalise_numeric(df[col])
            converted_non_null = converted.notna().sum()
            if (converted_non_null / non_null_count) >= min_fraction:
                df[col] = converted
                money_cols.append(col)

    # Force-convert common expected money columns if present (case-insensitive)
    for expected in ['Paid In', 'Withdrawn', 'Balance', 'Amount', 'Credit', 'Debit']:
        actual = _find_column_case_insensitive(df, expected)
        if actual:
            df[actual] = _normalise_numeric(df[actual])
            if actual not in money_cols and df[actual].notna().sum() > 0:
                money_cols.append(actual)

    return df, money_cols


# ─────────────────────────────────────────────
# PIVOT TABLE
# ─────────────────────────────────────────────
def create_pivot_table(df: pd.DataFrame, excel_writer: pd.ExcelWriter) -> None:
    """
    Generate a six-month financial summary pivot table
    and write it to the Excel workbook with formatting.
    """
    try:
        df = df.copy()

        # Defensive: ensure pivot columns are numeric (case-insensitive lookups)
        for expected in ['Paid In', 'Withdrawn', 'Balance']:
            actual = _find_column_case_insensitive(df, expected)
            if actual and df[actual].dtype == object:
                df[actual] = _normalise_numeric(df[actual])

        paid_col = _find_column_case_insensitive(df, 'Paid In')
        withdrawn_col = _find_column_case_insensitive(df, 'Withdrawn')
        balance_col = _find_column_case_insensitive(df, 'Balance')

        if 'Month' not in df.columns or (paid_col is None and withdrawn_col is None and balance_col is None):
            messagebox.showwarning("Pivot Warning", "Not enough columns to create pivot (need 'Month' and at least one of Paid In/Withdrawn/Balance).")
            return

        values = [c for c in [paid_col, withdrawn_col, balance_col] if c is not None]
        aggfunc = {}
        if paid_col:      aggfunc[paid_col] = 'sum'
        if withdrawn_col: aggfunc[withdrawn_col] = 'sum'
        if balance_col:   aggfunc[balance_col] = 'mean'

        pivot = pd.pivot_table(df, index='Month', values=values, aggfunc=aggfunc)

        # Standardize names: balance -> 'Average of Balance', rename paid/withdrawn if needed
        if balance_col and balance_col in pivot.columns:
            pivot = pivot.rename(columns={balance_col: 'Average of Balance'})
        rename_map = {}
        if paid_col and paid_col != 'Paid In' and paid_col in pivot.columns:
            rename_map[paid_col] = 'Paid In'
        if withdrawn_col and withdrawn_col != 'Withdrawn' and withdrawn_col in pivot.columns:
            rename_map[withdrawn_col] = 'Withdrawn'
        pivot = pivot.rename(columns=rename_map)

        # Reorder to Paid In, Withdrawn, Average of Balance if present
        ordered = [c for c in ['Paid In', 'Withdrawn', 'Average of Balance'] if c in pivot.columns]
        pivot = pivot[ordered]

        # Sort newest → oldest using Month index
        pivot.index = pd.to_datetime(pivot.index, format='%Y-%m', errors='coerce')
        pivot = pivot[pivot.index.notna()].sort_index(ascending=False)

        # 6-month window logic
        def safe_val(df_slice, row_idx, col):
            val = df_slice.iloc[row_idx][col]
            return 0.0 if pd.isna(val) else float(val)

        if len(pivot) <= 6:
            window = pivot.copy()
        else:
            top7 = pivot.iloc[:7]
            current_paid_in = safe_val(top7, 0, 'Paid In') if 'Paid In' in top7.columns else 0
            month7_paid_in  = safe_val(top7, 6, 'Paid In') if 'Paid In' in top7.columns else 0
            window = top7.iloc[0:6] if current_paid_in >= month7_paid_in else top7.iloc[1:7]

        # Ensure floats and round to 2 decimals for display
        window = window.astype(float).round(2)
        window.index = window.index.strftime('%Y-%m')
        pivot_main = window.copy()

        # Grand Total (keeps values for all columns)
        grand_total = pd.DataFrame([{
            'Paid In':            pivot_main['Paid In'].sum() if 'Paid In' in pivot_main.columns else np.nan,
            'Withdrawn':          pivot_main['Withdrawn'].sum() if 'Withdrawn' in pivot_main.columns else np.nan,
            'Average of Balance': pivot_main['Average of Balance'].mean() if 'Average of Balance' in pivot_main.columns else np.nan,
        }], index=['Grand Total']).round(2)

        # Average: Paid In mean
        avg_paid = pivot_main['Paid In'].mean() if 'Paid In' in pivot_main.columns else np.nan
        avg_row = pd.DataFrame([{
            'Paid In':            round(avg_paid, 2) if not pd.isna(avg_paid) else np.nan,
            'Withdrawn':          np.nan,
            'Average of Balance': np.nan,
        }], index=['Average'])

        # Discounting chain (as requested):
        # 70% Discounting = 70% of Average
        # Profitability @ 20% = 20% of 70% Discounting
        # Disposable Income @ 25% = 25% of Profitability @ 20%
        paid_mean = float(avg_row.at['Average', 'Paid In']) if 'Paid In' in avg_row.columns and not pd.isna(avg_row.at['Average', 'Paid In']) else 0.0

        discount70_val = round(paid_mean * 0.70, 2)
        profitability20_val = round(discount70_val * 0.20, 2)
        disposable25_val = round(profitability20_val * 0.25, 2)

        discount_70 = pd.DataFrame([{'Paid In': discount70_val, 'Withdrawn': np.nan, 'Average of Balance': np.nan}], index=['70% Discounting'])
        discount_20 = pd.DataFrame([{'Paid In': profitability20_val, 'Withdrawn': np.nan, 'Average of Balance': np.nan}], index=['Profitability @ 20%'])
        discount_25 = pd.DataFrame([{'Paid In': disposable25_val, 'Withdrawn': np.nan, 'Average of Balance': np.nan}], index=['Disposable Income @ 25%'])

        # Concat in final order
        pivot_final = pd.concat([pivot_main, grand_total, avg_row, discount_70, discount_20, discount_25])

        # Set index name for presentation
        pivot_final.index.name = 'Months'

        # Write the full pivot to Excel (pandas writes headers and values)
        pivot_final.to_excel(excel_writer, sheet_name='Pivot Table')

        # ----- Formatting: only format the highlighted rectangle area ----- #
        workbook = excel_writer.book
        worksheet = excel_writer.sheets['Pivot Table']

        # Define formats (borders applied ONLY to cells inside the highlighted area)
        header_fmt   = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D9E1F2'})
        idx_fmt      = workbook.add_format({'align': 'center', 'border': 1})           # Months now CENTER-aligned
        idx_bold_fmt = workbook.add_format({'align': 'center', 'border': 1, 'bold': True})
        center_fmt   = workbook.add_format({'align': 'center', 'border': 1, 'num_format': '#,##0.00;(#,##0.00)'})
        center_bold  = workbook.add_format({'align': 'center', 'border': 1, 'bold': True, 'num_format': '#,##0.00;(#,##0.00)'})
        blank_fmt    = workbook.add_format({'border': 1})

        # Determine excel coordinates
        start_row = 1
        n_rows = len(pivot_final)
        end_row = start_row + n_rows - 1
        data_cols = pivot_final.shape[1]
        last_col = data_cols  # index is col 0, data occupy 1..data_cols

        # Apply header format to header cells inside highlighted area (row 0)
        worksheet.write(0, 0, pivot_final.index.name if pivot_final.index.name else '', header_fmt)
        for j, col_name in enumerate(pivot_final.columns):
            worksheet.write(0, j+1, col_name, header_fmt)

        # Rewrite every cell inside the highlighted area to guarantee uniform formatting
        index_list = list(pivot_final.index)
        # summary labels to treat specially when formatting rows
        summary_labels = ('Grand Total', 'Average', '70% Discounting', 'Profitability @ 20%', 'Disposable Income @ 25%')

        for i, idx_label in enumerate(index_list):
            excel_row = start_row + i
            is_summary = idx_label in summary_labels

            # index cell (Months) — center aligned now
            if is_summary:
                worksheet.write(excel_row, 0, idx_label, idx_bold_fmt)
            else:
                worksheet.write(excel_row, 0, idx_label, idx_fmt)

            # data columns
            for j, col_name in enumerate(pivot_final.columns):
                excel_col = j + 1
                val = pivot_final.iloc[i, j]
                if pd.isna(val):
                    worksheet.write(excel_row, excel_col, '', blank_fmt)
                else:
                    # For Average and discount rows: only Paid In has value
                    if idx_label in ('Average', '70% Discounting', 'Profitability @ 20%', 'Disposable Income @ 25%'):
                        if col_name == 'Paid In':
                            # Average row bold; discounts non-bold (keep center alignment)
                            if idx_label == 'Average':
                                worksheet.write_number(excel_row, excel_col, float(val), center_bold)
                            else:
                                worksheet.write_number(excel_row, excel_col, float(val), center_fmt)
                        else:
                            worksheet.write(excel_row, excel_col, '', blank_fmt)
                    else:
                        # Normal months and Grand Total: center-aligned numeric; Grand Total bold
                        if idx_label == 'Grand Total':
                            worksheet.write_number(excel_row, excel_col, float(val), center_bold)
                        else:
                            worksheet.write_number(excel_row, excel_col, float(val), center_fmt)

        # Done — cells outside the rectangle are untouched (no borders added)

    except Exception as e:
        messagebox.showwarning("Pivot Warning", f"Could not create pivot table: {e}")
        raise
# ─────────────────────────────────────────────
# LOAN FILTER
# ─────────────────────────────────────────────
LOAN_COMPANIES = [
    "Abito Limited", "Absolute Credit Kenya Ltd", "Acquire Credit Limited",
    "Adjacent Possible Finance Limited", "Adroit Credit Limited", "Ajax Credit Kenya Limited",
    "Aleza Limited", "Ambush Capital Limited", "FlashPesa", "Anjoy Credit Limited",
    "ASA International Kenya Limited", "Asante FS East Africa Limited", "Asap Credit Limited",
    "Aspire Lending Ltd", "Autochek Limited", "Auxiliary Credit Ltd", "Avenews Ke Ltd",
    "Aventus Technology Limited", "Lendplus", "AVL Capital Ltd", "Azura Credit Limited",
    "BCF Kenya Limited", "Bidii Credit Limited", "Bimas Kenya Limited",
    "Bingwa Micro Capital Limited", "Blesmark Credit Limited", "Boostline Capital Limited",
    "Bossrich Credit Limited", "BRAC Kenya Company Limited", "Brisk Credit Limited",
    "Bytech Credit Limited", "Cashmart Capital Limited", "Ceres Tech Limited",
    "Chapeo Capital Limited", "Chelete Credit Limited", "Chime Capital Limited", "Credit",
    "Creditarea Capital Limited", "Kashbean", "Decimal Capital Limited",
    "Dexintec Kenya Limited", "Dime Credit Limited", "Dotcash Credit Limited",
    "East Africa Futures Company Limited", "Easy Asset Management Limited",
    "Easyways Credit Limited", "ED Partners Africa Limited", "Edenbridge Capital Ltd",
    "EDOMX Limited", "Elevate Credit Limited", "Ellegant Credit Limited",
    "Extend Money Services Limited", "Fabilo Credit Ltd", "Factorhouse Limited",
    "Fahari Point Capital Limited", "Fantom Capital Limited", "Fezotech Kenya Limited",
    "Finberry Capital Ltd", "Finboom Credit Kenya Limited", "Fincorp Credit Limited",
    "Fincredit Limited", "Finseil Limited", "Fortune Credit Limited",
    "Fourth Generation Capital Limited", "4G Capital", "Frictionless Enterprises Limited",
    "M-Power", "Futureinno Digital Tech Limited", "Geoland Credit Limited",
    "Getcash Capital Limited", "Giando Africa Limited", "Flash Credit Africa",
    "Girls First Kenya Limited", "Granary Capital Limited", "Guava Capital Limited",
    "Hanis Capital Limited", "Hela Capital Limited", "Helium Credit Limited",
    "Inspire Credit Limited", "Inventure Mobile Limited", "Tala", "Ismuk Credit Limited",
    "Jackfruit Associates Limited", "Jafari Credit Limited", "Jambofin Credit Limited",
    "Jijenge Credit Limited", "Juhudi Kilimo Company Limited", "Jumo Kenya Limited",
    "Keep Vision and Growth Credit Ltd", "KCB M-PESA", "Kifedha Ltd", "Kikwetu Credit Ltd",
    "Kweli Smart Solutions Limited", "Lasiri Capital Limited", "Leaf Credit Limited",
    "Leja Ltd", "Lenana Innovative Solutions Limited", "Letshego Kenya Ltd",
    "Liberty Afrika Technologies Ltd", "Lipa Later Limited", "Little Limited", "SpotIt",
    "Little Pesa Limited", "Loan Plus Digital Credit Provider Ltd", "Lobelitec Credit Limited",
    "LockBx Limited", "Longitude Capital Limited", "Lucason Capital Limited",
    "Maison Capital Limited", "Malicash Investment Limited", "Maralal Ledger Limited",
    "Marble Capital Solutions Limited", "Maxxton Enterprises Ltd", "Mayflower Capital Limited",
    "MCF 2 Kenya Limited", "Medical Credit Fund", "Mednow Capital Limited",
    "MFS Technologies Limited", "Milhan Access Capital Limited", "Mimi Credit Limited",
    "Mint Credit Limited", "MKash Solutions Limited", "MKM Capital Limited", "M-Shwari",
    "M-Kopa", "Mkulimapay Credit Ltd", "Modesty Credit Ltd", "Mogo",
    "Momentum Credit Limited", "Moneza Ltd", "Moto Hope Capital Limited",
    "Mular Credit Limited", "Musoni Capital Limited", "Mwananchi Credit Ltd",
    "Mwanzo Credit Limited", "Mycredit Limited", "MyWagepay Limited", "Natal Tech Limited",
    "Nawiri African Sprouts Ltd", "Fast Credit", "Newark Frontiers Limited",
    "Ngao Credit Limited", "Numida Technologies Kenya Limited", "ODI Credit Limited",
    "Okolea International Limited", "Onwards Swift Company Limited", "Opal Quick Limited",
    "Otas Credit Limited", "Pato Capital Limited", "Payablu Credit Limited",
    "Pembeni Cash Ltd", "Pesaglow Capital Limited", "Pesakuu Credit Limited", "iZiLoan",
    "Peshee Capital Limited", "Pezesha Africa Limited", "Phoenix Capital Limited",
    "Pi Capital Limited", "Bayes", "Platinum Credit Limited", "Premier Credit Limited",
    "Progressive Credit Limited", "Puphik Credit Limited", "Radi Credit Limited",
    "Real People Kenya Limited", "Reazilla DCP Limited", "Rewot Ciro Limited",
    "Risine Credit Limited", "Rosatap Credit Limited", "Seanala Credit Limited",
    "Select Management Services Limited", "Senti Capital Limited", "Sevi Innovation Limited",
    "Simbageld Ltd", "Simplepay Capital Limited", "Sipranda Capital Ltd",
    "Siti Mobility Technologies Limited", "Snowflex Capital Limited", "Sokohela Limited",
    "Spectrum Credit Ltd", "Swift Capital", "Spread Capital Ltd", "Steadfast Credit Ltd",
    "Stride Credit Limited", "Suffice Ltd", "Sure Cred Capital Limited",
    "Tanir Credit & Accounting Services Ltd", "Tazu Credit Limited",
    "Tenakata Enterprises Limited", "Tentacorp Holdings Limited",
    "Tinycost Credit Kenya Limited", "Tip-Point Capital Limited", "Transsnet Credit Limited",
    "Treasure Store Limited", "TrustGro SCA Limited", "UbaPesa Limited",
    "Umoja Fanisi Limited", "Unidirect Ltd", "Unifi Credit Limited",
    "Vision Edge Credit Partners Limited", "Fast Growth", "Wabema Credit Ltd",
    "Wakanda Credit Limited", "Watu Credit Ltd", "Westlip Credit Limited",
    "Zaidi Pato Limited", "Zamaradi Capital & Credit Group Ltd", "Zanifu Limited",
    "Zenka Digital Limited", "Promotion", "Salary", "Zash Loan", "Loan",
    "MICROMART AFRICA LIMITED", "DOWNTOWN AFRICA", "PAGE CAPITAL LIMITED",
]

_EXCLUDE_PATTERN = re.compile(
    r"overdraft of credit party|m-pesa overdraw|mpesa overdraw|od loan repayment",
    re.IGNORECASE,
)
_INCLUDE_PATTERN = re.compile(
    "|".join(re.escape(c) for c in LOAN_COMPANIES),
    re.IGNORECASE,
)


def filter_loans(df: pd.DataFrame, excel_writer: pd.ExcelWriter, money_cols: list = None) -> None:
    """
    Filter loan entries and write a 'Loans' sheet. If money_cols is provided,
    use it (case-insensitively) to format the relevant columns; otherwise
    fall back to pandas.is_numeric_dtype to detect numeric columns safely.

    This function ALWAYS writes a 'Loans' sheet:
      - If matching rows exist, those rows are written.
      - If no matches, an empty sheet with headers is written plus a note row.
      - A companion sheet 'Loan Companies' is written listing all LOAN_COMPANIES.
    """
    try:
        work = df.copy()
        detail_col = None
        for c in work.columns:
            if isinstance(c, str) and 'detail' in c.strip().lower():
                detail_col = c
                break

        if not detail_col:
            messagebox.showwarning("Loan Filter Warning", "No 'Details' column found.")
            # Still create Loans sheet with headers if possible
            empty = pd.DataFrame(columns=work.columns)
            empty.to_excel(excel_writer, sheet_name="Loans", index=False)
            return

        work[detail_col] = work[detail_col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

        loan_mask = work[detail_col].str.contains(_INCLUDE_PATTERN, na=False)
        exclude_mask = work[detail_col].str.contains(_EXCLUDE_PATTERN, na=False)
        loan_df = work[loan_mask & ~exclude_mask]

        workbook = excel_writer.book

        if not loan_df.empty:
            loan_df.to_excel(excel_writer, sheet_name="Loans", index=False)
            worksheet = excel_writer.sheets['Loans']

            money_fmt = workbook.add_format({'num_format': '#,##0.00'})

            # Preferred: format columns based on supplied money_cols (case-insensitive)
            formatted = set()
            if money_cols:
                col_map = {c.lower(): c for c in loan_df.columns if isinstance(c, str)}
                for m in money_cols:
                    key = m.lower()
                    if key in col_map:
                        try:
                            col_idx = loan_df.columns.get_loc(col_map[key])
                            worksheet.set_column(col_idx, col_idx, 18, money_fmt)
                            formatted.add(col_idx)
                        except Exception:
                            pass

            # Fallback: use pandas' is_numeric_dtype which understands extension dtypes
            if not formatted:
                from pandas.api import types as ptypes
                for i, col in enumerate(loan_df.columns):
                    try:
                        if ptypes.is_numeric_dtype(loan_df[col]):
                            worksheet.set_column(i, i, 18, money_fmt)
                    except Exception:
                        continue

            print(f"[Loans] {len(loan_df)} loan entries written.")
        else:
            # No matches: write Loans sheet with headers and a note row
            empty = pd.DataFrame(columns=work.columns)
            empty.to_excel(excel_writer, sheet_name="Loans", index=False)
            worksheet = excel_writer.sheets['Loans']
            # Write a note in A2 (row 1, col 0) to explain why the sheet is empty
            try:
                worksheet.write(1, 0, "No matching loan entries found for the LOAN_COMPANIES list.")
            except Exception:
                pass



    except Exception as e:
        messagebox.showwarning("Loan Filter Warning", f"Could not filter loans: {e}")


# ─────────────────────────────────────────────
# MAIN CONVERSION
# ─────────────────────────────────────────────
def convert_pdf_to_excel(pdf_path: str, excel_path: str) -> None:
    df = extract_tables_with_pdfplumber(pdf_path)
    df = format_dates(df)

    # Detect and coerce money-like columns (and force known columns)
    df_numeric, money_cols = coerce_money_columns(df)

    # Ensure pivot-required columns are numeric (defensive)
    for expected in ['Paid In', 'Withdrawn', 'Balance']:
        actual = _find_column_case_insensitive(df_numeric, expected)
        if actual and df_numeric[actual].dtype == object:
            df_numeric[actual] = _normalise_numeric(df_numeric[actual])
            if actual not in money_cols:
                money_cols.append(actual)

    # Write everything using xlsxwriter engine so we can format cells
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        # Full Data sheet — write numeric-converted DF
        df_numeric.to_excel(writer, sheet_name='Full Data', index=False)

        workbook = writer.book
        worksheet = writer.sheets['Full Data']
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})

        # Apply number formatting to detected money columns
        for col in money_cols:
            try:
                col_idx = df_numeric.columns.get_loc(col)
                worksheet.set_column(col_idx, col_idx, 18, money_fmt)
            except Exception:
                continue

        # Create pivot and loans using numeric dataframe; pass money_cols into filter_loans
        create_pivot_table(df_numeric, writer)
        filter_loans(df_numeric, writer, money_cols)


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
def main() -> None:
    root = Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Select M-Pesa PDF",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not pdf_path:
        sys.exit()

    folder = os.path.dirname(pdf_path)

# Use the PDF filename for the Excel output
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    excel_path = os.path.join(folder, f"{base_name}.xlsx")

    unlocked_path = None
    try:
        answer = messagebox.askyesno("PDF Password", "Is the PDF password-protected?")
        if answer:
            password = simpledialog.askstring("PDF Password", "Enter the M-Pesa PDF password:", show='*')
            if not password:
                sys.exit()
            unlocked_path = os.path.join(folder, TEMP_UNLOCKED_PDF)
            pdf_to_use = unlock_pdf(pdf_path, unlocked_path, password)
        else:
            pdf_to_use = pdf_path

        convert_pdf_to_excel(pdf_to_use, excel_path)
        messagebox.showinfo("Success", f"Excel file saved:\n{excel_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

    finally:
        if unlocked_path and os.path.exists(unlocked_path):
            try:
                os.remove(unlocked_path)
                print(f"[Cleanup] Removed temp file: {unlocked_path}")
            except OSError:
                pass


if __name__ == "__main__":
    main()