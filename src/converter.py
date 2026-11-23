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
APP_NAME = "M-Pesa Statement Converter"
VERSION = "1.0.0"
OUTPUT_FILENAME = "mpesa_statement.xlsx"
TEMP_UNLOCKED_PDF = "unlocked_mpesa.pdf"


# ─────────────────────────────────────────────
# PDF UNLOCK
# ─────────────────────────────────────────────
def unlock_pdf(input_path: str, output_path: str, password: str) -> str:
    """Unlock a password-protected M-Pesa PDF and save a temporary copy."""
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
    """Extract tables from all pages of a PDF into one DataFrame."""
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

    # Strip whitespace from string cells and column names.
    combined = combined.apply(
        lambda col: col.str.strip() if col.dtype == object else col
    )
    combined.columns = [
        col.strip() if isinstance(col, str) else col
        for col in combined.columns
    ]

    return combined


# ─────────────────────────────────────────────
# DATE FORMATTING
# ─────────────────────────────────────────────
_DATE_FORMATS = [
    "%d/%m/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%d-%m-%Y %H:%M:%S",
    "%d/%m/%Y",
    "%Y-%m-%d",
]


def _parse_date_series(series: pd.Series) -> pd.Series:
    """Try known date formats, then fall back to pandas date parsing."""
    for date_format in _DATE_FORMATS:
        parsed = pd.to_datetime(
            series,
            format=date_format,
            errors='coerce',
        )
        if parsed.notna().any():
            return parsed

    return pd.to_datetime(series, errors='coerce')


def format_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Parse transaction dates and create a Month column for the pivot."""
    df = df.copy()
    df.columns = [
        col.strip() if isinstance(col, str) else col
        for col in df.columns
    ]

    if 'Completion Time' in df.columns:
        parsed = _parse_date_series(df['Completion Time'])
        null_count = parsed.isna().sum()

        if null_count > 0:
            messagebox.showwarning(
                "Date Warning",
                f"{null_count} row(s) have unrecognised date formats "
                "and will be excluded from the pivot."
            )

        df['Month'] = parsed.dt.strftime('%Y-%m')
    else:
        messagebox.showwarning(
            "Date Warning",
            "'Completion Time' column not found."
        )

    return df


# ─────────────────────────────────────────────
# NUMERIC NORMALISER AND COLUMN LOOKUP
# ─────────────────────────────────────────────
def _normalise_numeric(series: pd.Series) -> pd.Series:
    """Remove currency/text characters and convert accounting negatives."""
    cleaned = (
        series.astype(str)
        .str.replace(r'[^0-9.\-()]', '', regex=True)
        .str.replace(r'^\((.+)\)$', r'-\1', regex=True)
    )
    return pd.to_numeric(cleaned, errors='coerce')


def _find_column_case_insensitive(
    df: pd.DataFrame,
    name: str,
):
    """Find a column by name, ignoring case and surrounding whitespace."""
    expected = name.strip().lower()

    for col in df.columns:
        if isinstance(col, str) and col.strip().lower() == expected:
            return col

    return None


def coerce_money_columns(
    df: pd.DataFrame,
    min_fraction: float = 0.5,
):
    """
    Convert columns that look like monetary values to numeric.

    Returns:
        A tuple of (converted DataFrame, list of monetary column names).
    """
    df = df.copy()
    money_cols = []

    # Detect money-like object columns.
    for col in df.columns:
        if df[col].dtype == object:
            non_null_count = df[col].notna().sum()
            if non_null_count == 0:
                continue

            converted = _normalise_numeric(df[col])
            converted_count = converted.notna().sum()

            if converted_count / non_null_count >= min_fraction:
                df[col] = converted
                money_cols.append(col)

    # Also convert common monetary columns, if present.
    for expected in [
        'Paid In', 'Withdrawn', 'Balance', 'Amount', 'Credit', 'Debit'
    ]:
        actual = _find_column_case_insensitive(df, expected)
        if actual:
            df[actual] = _normalise_numeric(df[actual])
            if actual not in money_cols and df[actual].notna().any():
                money_cols.append(actual)

    return df, money_cols


# ─────────────────────────────────────────────
# PIVOT TABLE
# ─────────────────────────────────────────────
def create_pivot_table(
    df: pd.DataFrame,
    excel_writer: pd.ExcelWriter,
) -> None:
    """Create and format the six-month financial summary pivot table."""
    try:
        df = df.copy()

        # Ensure pivot data columns are numeric.
        for expected in ['Paid In', 'Withdrawn', 'Balance']:
            actual = _find_column_case_insensitive(df, expected)
            if actual and df[actual].dtype == object:
                df[actual] = _normalise_numeric(df[actual])

        paid_col = _find_column_case_insensitive(df, 'Paid In')
        withdrawn_col = _find_column_case_insensitive(df, 'Withdrawn')
        balance_col = _find_column_case_insensitive(df, 'Balance')

        if 'Month' not in df.columns or all(
            col is None for col in [paid_col, withdrawn_col, balance_col]
        ):
            messagebox.showwarning(
                "Pivot Warning",
                "Not enough columns to create pivot "
                "(need 'Month' and at least one of "
                "Paid In/Withdrawn/Balance)."
            )
            return

        values = [
            col for col in [paid_col, withdrawn_col, balance_col]
            if col is not None
        ]

        aggfunc = {}
        if paid_col:
            aggfunc[paid_col] = 'sum'
        if withdrawn_col:
            aggfunc[withdrawn_col] = 'sum'
        if balance_col:
            aggfunc[balance_col] = 'mean'

        pivot = pd.pivot_table(
            df,
            index='Month',
            values=values,
            aggfunc=aggfunc,
        )

        # Standardize pivot column names.
        rename_map = {}
        if balance_col and balance_col in pivot.columns:
            rename_map[balance_col] = 'Average of Balance'
        if paid_col and paid_col != 'Paid In' and paid_col in pivot.columns:
            rename_map[paid_col] = 'Paid In'
        if (
            withdrawn_col
            and withdrawn_col != 'Withdrawn'
            and withdrawn_col in pivot.columns
        ):
            rename_map[withdrawn_col] = 'Withdrawn'

        pivot = pivot.rename(columns=rename_map)

        ordered = [
            col for col in ['Paid In', 'Withdrawn', 'Average of Balance']
            if col in pivot.columns
        ]
        pivot = pivot[ordered]

        # Sort from newest to oldest.
        pivot.index = pd.to_datetime(
            pivot.index,
            format='%Y-%m',
            errors='coerce',
        )
        pivot = pivot[pivot.index.notna()].sort_index(ascending=False)

        def safe_val(df_slice, row_idx, col):
            value = df_slice.iloc[row_idx][col]
            return 0.0 if pd.isna(value) else float(value)

        # Select the six-month window.
        if len(pivot) <= 6:
            window = pivot.copy()
        else:
            top7 = pivot.iloc[:7]
            current_paid_in = (
                safe_val(top7, 0, 'Paid In')
                if 'Paid In' in top7.columns
                else 0
            )
            month7_paid_in = (
                safe_val(top7, 6, 'Paid In')
                if 'Paid In' in top7.columns
                else 0
            )
            window = (
                top7.iloc[0:6]
                if current_paid_in >= month7_paid_in
                else top7.iloc[1:7]
            )

        window = window.astype(float).round(2)
        window.index = window.index.strftime('%Y-%m')
        pivot_main = window.copy()

        # Grand Total row.
        grand_total = pd.DataFrame([{
            'Paid In': (
                pivot_main['Paid In'].sum()
                if 'Paid In' in pivot_main.columns else np.nan
            ),
            'Withdrawn': (
                pivot_main['Withdrawn'].sum()
                if 'Withdrawn' in pivot_main.columns else np.nan
            ),
            'Average of Balance': (
                pivot_main['Average of Balance'].mean()
                if 'Average of Balance' in pivot_main.columns else np.nan
            ),
        }], index=['Grand Total']).round(2)

        # Average Paid In row.
        avg_paid = (
            pivot_main['Paid In'].mean()
            if 'Paid In' in pivot_main.columns
            else np.nan
        )
        avg_row = pd.DataFrame([{
            'Paid In': round(avg_paid, 2) if not pd.isna(avg_paid) else np.nan,
            'Withdrawn': np.nan,
            'Average of Balance': np.nan,
        }], index=['Average'])

        # Discounting chain.
        paid_mean = (
            float(avg_row.at['Average', 'Paid In'])
            if (
                'Paid In' in avg_row.columns
                and not pd.isna(avg_row.at['Average', 'Paid In'])
            )
            else 0.0
        )

        discount70_val = round(paid_mean * 0.70, 2)
        profitability20_val = round(discount70_val * 0.20, 2)
        disposable25_val = round(profitability20_val * 0.25, 2)

        discount_70 = pd.DataFrame([{
            'Paid In': discount70_val,
            'Withdrawn': np.nan,
            'Average of Balance': np.nan,
        }], index=['70% Discounting'])

        discount_20 = pd.DataFrame([{
            'Paid In': profitability20_val,
            'Withdrawn': np.nan,
            'Average of Balance': np.nan,
        }], index=['Profitability @ 20%'])

        discount_25 = pd.DataFrame([{
            'Paid In': disposable25_val,
            'Withdrawn': np.nan,
            'Average of Balance': np.nan,
        }], index=['Disposable Income @ 25%'])

        pivot_final = pd.concat([
            pivot_main,
            grand_total,
            avg_row,
            discount_70,
            discount_20,
            discount_25,
        ])
        pivot_final.index.name = 'Months'

        # Write pivot data; formatting below rewrites the displayed cells.
        pivot_final.to_excel(excel_writer, sheet_name='Pivot Table')

        workbook = excel_writer.book
        worksheet = excel_writer.sheets['Pivot Table']

        # Cell formats.
        header_fmt = workbook.add_format({
            'bold': True,
            'align': 'center',
            'border': 1,
            'bg_color': '#D9E1F2',
        })
        idx_fmt = workbook.add_format({
            'align': 'center',
            'border': 1,
        })
        idx_bold_fmt = workbook.add_format({
            'align': 'left',
            'border': 1,
            'bold': True,
        })
        center_fmt = workbook.add_format({
            'align': 'center',
            'border': 1,
            'num_format': '#,##0.00;(#,##0.00)',
        })
        center_bold = workbook.add_format({
            'align': 'center',
            'border': 1,
            'bold': True,
            'num_format': '#,##0.00;(#,##0.00)',
        })
        disposable_bold = workbook.add_format({
            'align': 'center',
            'border': 1,
            'bold': True,
            'num_format': '#,##0.00;(#,##0.00)',
        })
        blank_fmt = workbook.add_format({'border': 1})

        # Format the header row.
        worksheet.write(0, 0, 'Months', header_fmt)
        for col_idx, col_name in enumerate(pivot_final.columns, start=1):
            worksheet.write(0, col_idx, col_name, header_fmt)

        summary_labels = {
            'Grand Total',
            'Average',
            '70% Discounting',
            'Profitability @ 20%',
            'Disposable Income @ 25%',
        }
        discount_labels = {
            '70% Discounting',
            'Profitability @ 20%',
            'Disposable Income @ 25%',
        }

        # Format row labels and values.
        for row_idx, label in enumerate(pivot_final.index, start=1):
            label_fmt = idx_bold_fmt if label in summary_labels else idx_fmt
            worksheet.write(row_idx, 0, label, label_fmt)

            for col_idx, col_name in enumerate(pivot_final.columns, start=1):
                value = pivot_final.loc[label, col_name]

                if pd.isna(value):
                    worksheet.write(row_idx, col_idx, '', blank_fmt)
                elif label == 'Average':
                    worksheet.write_number(
                        row_idx, col_idx, float(value), center_fmt
                    )
                elif label == 'Grand Total':
                    worksheet.write_number(
                        row_idx, col_idx, float(value), center_bold
                    )
                elif label in discount_labels:
                    amount_fmt = (
                        disposable_bold
                        if label == 'Disposable Income @ 25%'
                        else center_fmt
                    )
                    worksheet.write_number(
                        row_idx, col_idx, float(value), amount_fmt
                    )
                else:
                    worksheet.write_number(
                        row_idx, col_idx, float(value), center_fmt
                    )

    except Exception as e:
        messagebox.showwarning(
            "Pivot Warning",
            f"Could not create pivot table: {e}",
        )
        raise


# ─────────────────────────────────────────────
# LOAN FILTER
# ─────────────────────────────────────────────
LOAN_COMPANIES = [
    "Abito Limited", "Absolute Credit Kenya Ltd", "Acquire Credit Limited",
    "Adjacent Possible Finance Limited", "Adroit Credit Limited",
    "Ajax Credit Kenya Limited", "Aleza Limited", "Ambush Capital Limited",
    "FlashPesa", "Anjoy Credit Limited", "ASA International Kenya Limited",
    "Asante FS East Africa Limited", "Asap Credit Limited",
    "Aspire Lending Ltd", "Autochek Limited", "Auxiliary Credit Ltd",
    "Avenews Ke Ltd", "Aventus Technology Limited", "Lendplus",
    "AVL Capital Ltd", "Azura Credit Limited", "BCF Kenya Limited",
    "Bidii Credit Limited", "Bimas Kenya Limited",
    "Bingwa Micro Capital Limited", "Blesmark Credit Limited",
    "Boostline Capital Limited", "Bossrich Credit Limited",
    "BRAC Kenya Company Limited", "Brisk Credit Limited",
    "Bytech Credit Limited", "Cashmart Capital Limited",
    "Ceres Tech Limited", "Chapeo Capital Limited",
    "Chelete Credit Limited", "Chime Capital Limited", "Credit",
    "Creditarea Capital Limited", "Kashbean", "Decimal Capital Limited",
    "Dexintec Kenya Limited", "Dime Credit Limited",
    "Dotcash Credit Limited", "East Africa Futures Company Limited",
    "Easy Asset Management Limited", "Easyways Credit Limited",
    "ED Partners Africa Limited", "Edenbridge Capital Ltd",
    "EDOMX Limited", "Elevate Credit Limited", "Ellegant Credit Limited",
    "Extend Money Services Limited", "Fabilo Credit Ltd",
    "Factorhouse Limited", "Fahari Point Capital Limited",
    "Fantom Capital Limited", "Fezotech Kenya Limited",
    "Finberry Capital Ltd", "Finboom Credit Kenya Limited",
    "Fincorp Credit Limited", "Fincredit Limited", "Finseil Limited",
    "Fortune Credit Limited", "Fourth Generation Capital Limited",
    "4G Capital", "Frictionless Enterprises Limited", "M-Power",
    "Futureinno Digital Tech Limited", "Geoland Credit Limited",
    "Getcash Capital Limited", "Giando Africa Limited",
    "Flash Credit Africa", "Girls First Kenya Limited",
    "Granary Capital Limited", "Guava Capital Limited",
    "Hanis Capital Limited", "Hela Capital Limited",
    "Helium Credit Limited", "Inspire Credit Limited",
    "Inventure Mobile Limited", "Tala", "Ismuk Credit Limited",
    "Jackfruit Associates Limited", "Jafari Credit Limited",
    "Jambofin Credit Limited", "Jijenge Credit Limited",
    "Juhudi Kilimo Company Limited", "Jumo Kenya Limited",
    "Keep Vision and Growth Credit Ltd", "KCB M-PESA",
    "Kifedha Ltd", "Kikwetu Credit Ltd",
    "Kweli Smart Solutions Limited", "Lasiri Capital Limited",
    "Leaf Credit Limited", "Leja Ltd",
    "Lenana Innovative Solutions Limited", "Letshego Kenya Ltd",
    "Liberty Afrika Technologies Ltd", "Lipa Later Limited",
    "Little Limited", "SpotIt", "Little Pesa Limited",
    "Loan Plus Digital Credit Provider Ltd", "Lobelitec Credit Limited",
    "LockBx Limited", "Longitude Capital Limited",
    "Lucason Capital Limited", "Maison Capital Limited",
    "Malicash Investment Limited", "Maralal Ledger Limited",
    "Marble Capital Solutions Limited", "Maxxton Enterprises Ltd",
    "Mayflower Capital Limited", "MCF 2 Kenya Limited",
    "Medical Credit Fund", "Mednow Capital Limited",
    "MFS Technologies Limited", "Milhan Access Capital Limited",
    "Mimi Credit Limited", "Mint Credit Limited",
    "MKash Solutions Limited", "MKM Capital Limited", "M-Shwari",
    "M-Kopa", "Mkulimapay Credit Ltd", "Modesty Credit Ltd", "Mogo",
    "Momentum Credit Limited", "Moneza Ltd", "Moto Hope Capital Limited",
    "Mular Credit Limited", "Musoni Capital Limited",
    "Mwananchi Credit Ltd", "Mwanzo Credit Limited",
    "Mycredit Limited", "MyWagepay Limited",
    "Natal Tech Limited", "Nawiri African Sprouts Ltd",
    "Fast Credit", "Newark Frontiers Limited", "Ngao Credit Limited",
    "Numida Technologies Kenya Limited", "ODI Credit Limited",
    "Okolea International Limited", "Onwards Swift Company Limited",
    "Opal Quick Limited", "Otas Credit Limited", "Pato Capital Limited",
    "Payablu Credit Limited", "Pembeni Cash Ltd",
    "Pesaglow Capital Limited", "Pesakuu Credit Limited", "iZiLoan",
    "Peshee Capital Limited", "Pezesha Africa Limited",
    "Phoenix Capital Limited", "Pi Capital Limited", "Bayes",
    "Platinum Credit Limited", "Premier Credit Limited",
    "Progressive Credit Limited", "Puphik Credit Limited",
    "Radi Credit Limited", "Real People Kenya Limited",
    "Reazilla DCP Limited", "Rewot Ciro Limited",
    "Risine Credit Limited", "Rosatap Credit Limited",
    "Seanala Credit Limited", "Select Management Services Limited",
    "Senti Capital Limited", "Sevi Innovation Limited",
    "Simbageld Ltd", "Simplepay Capital Limited",
    "Sipranda Capital Ltd", "Siti Mobility Technologies Limited",
    "Snowflex Capital Limited", "Sokohela Limited",
    "Spectrum Credit Ltd", "Swift Capital", "Spread Capital Ltd",
    "Steadfast Credit Ltd", "Stride Credit Limited", "Suffice Ltd",
    "Sure Cred Capital Limited",
    "Tanir Credit & Accounting Services Ltd", "Tazu Credit Limited",
    "Tenakata Enterprises Limited", "Tentacorp Holdings Limited",
    "Tinycost Credit Kenya Limited", "Tip-Point Capital Limited",
    "Transsnet Credit Limited", "Treasure Store Limited",
    "TrustGro SCA Limited", "UbaPesa Limited",
    "Umoja Fanisi Limited", "Unidirect Ltd", "Unifi Credit Limited",
    "Vision Edge Credit Partners Limited", "Fast Growth",
    "Wabema Credit Ltd", "Wakanda Credit Limited", "Watu Credit Ltd",
    "Westlip Credit Limited", "Zaidi Pato Limited",
    "Zamaradi Capital & Credit Group Ltd", "Zanifu Limited",
    "Zenka Digital Limited", "Promotion", "Salary", "Zash Loan", "Loan",
    "MICROMART AFRICA LIMITED", "DOWNTOWN AFRICA",
    "PAGE CAPITAL LIMITED",

    # These are written to the Overdrafts sheet, not the Loans sheet.
    "overdraft of credit party", "m-pesa overdraw",
    "mpesa overdraw", "od loan repayment",
]

_INCLUDE_PATTERN = re.compile(
    "|".join(re.escape(company) for company in LOAN_COMPANIES),
    re.IGNORECASE,
)

_OVERDRAFT_PATTERN = re.compile(
    r"overdraft of credit party|m-pesa overdraw|mpesa overdraw|od loan repayment",
    re.IGNORECASE,
)


def _write_total_row(
    dataframe: pd.DataFrame,
    detail_col: str,
    paid_col: str | None,
    withdrawn_col: str | None,
) -> pd.DataFrame:
    """Append a TOTAL row with Paid In and Withdrawn sums."""
    total = {col: np.nan for col in dataframe.columns}
    total[detail_col] = 'TOTAL'

    if paid_col:
        total[paid_col] = _normalise_numeric(dataframe[paid_col]).sum()
    if withdrawn_col:
        total[withdrawn_col] = _normalise_numeric(
            dataframe[withdrawn_col]
        ).sum()

    return pd.concat(
        [dataframe, pd.DataFrame([total])],
        ignore_index=True,
    )


def _format_money_columns(
    worksheet,
    workbook,
    dataframe: pd.DataFrame,
    money_cols: list,
) -> None:
    """Apply a number format to known monetary columns."""
    money_fmt = workbook.add_format({'num_format': '#,##0.00'})

    columns_to_format = set(money_cols or [])
    for expected in ['Paid In', 'Withdrawn', 'Balance', 'Amount']:
        actual = _find_column_case_insensitive(dataframe, expected)
        if actual:
            columns_to_format.add(actual)

    for col in columns_to_format:
        if col in dataframe.columns:
            col_idx = dataframe.columns.get_loc(col)
            worksheet.set_column(col_idx, col_idx, 18, money_fmt)


def filter_loans(
    df: pd.DataFrame,
    excel_writer: pd.ExcelWriter,
    money_cols: list = None,
) -> None:
    """
    Write loan transactions to Loans and overdraft-related transactions
    to Overdrafts. Each sheet includes Paid In and Withdrawn totals.
    """
    try:
        work = df.copy()
        detail_col = next(
            (
                col for col in work.columns
                if isinstance(col, str) and 'detail' in col.strip().lower()
            ),
            None,
        )

        if not detail_col:
            messagebox.showwarning(
                "Loan Filter Warning",
                "No 'Details' column found."
            )
            empty = pd.DataFrame(columns=work.columns)
            empty.to_excel(excel_writer, sheet_name='Loans', index=False)
            empty.to_excel(excel_writer, sheet_name='Overdrafts', index=False)
            return

        work[detail_col] = (
            work[detail_col]
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        # Find the amount columns on the full dataframe. This ensures they
        # are available even if no overdraft transactions match.
        paid_col = _find_column_case_insensitive(work, 'Paid In')
        withdrawn_col = _find_column_case_insensitive(work, 'Withdrawn')

        # Overdraft transactions go to their own sheet.
        overdraft_mask = work[detail_col].str.contains(
            _OVERDRAFT_PATTERN,
            na=False,
        )
        overdraft_df = work.loc[overdraft_mask].copy()

        for col in [paid_col, withdrawn_col]:
            if col:
                overdraft_df[col] = _normalise_numeric(overdraft_df[col])

        overdraft_output = _write_total_row(
            overdraft_df,
            detail_col,
            paid_col,
            withdrawn_col,
        )
        overdraft_output.to_excel(
            excel_writer,
            sheet_name='Overdrafts',
            index=False,
        )

        # Matching loan transactions, excluding overdrafts already listed
        # on the Overdrafts sheet.
        loan_mask = work[detail_col].str.contains(_INCLUDE_PATTERN, na=False)
        loan_df = work.loc[loan_mask & ~overdraft_mask].copy()

        for col in [paid_col, withdrawn_col]:
            if col:
                loan_df[col] = _normalise_numeric(loan_df[col])

        if not loan_df.empty:
            # Sort Paid In descending, then Withdrawn ascending.
            sort_cols = []
            sort_order = []

            if paid_col:
                sort_cols.append(paid_col)
                sort_order.append(False)
            if withdrawn_col:
                sort_cols.append(withdrawn_col)
                sort_order.append(True)

            if sort_cols:
                loan_df = loan_df.sort_values(
                    by=sort_cols,
                    ascending=sort_order,
                    na_position='last',
                    kind='stable',
                ).reset_index(drop=True)

            loan_output = _write_total_row(
                loan_df,
                detail_col,
                paid_col,
                withdrawn_col,
            )
        else:
            loan_output = loan_df

        loan_output.to_excel(
            excel_writer,
            sheet_name='Loans',
            index=False,
        )

        workbook = excel_writer.book

        overdraft_ws = excel_writer.sheets['Overdrafts']
        _format_money_columns(
            overdraft_ws,
            workbook,
            overdraft_output,
            money_cols,
        )

        loan_ws = excel_writer.sheets['Loans']
        _format_money_columns(
            loan_ws,
            workbook,
            loan_output,
            money_cols,
        )

        print(
            f"[Loans] {len(loan_df)} loan entries written; "
            f"{len(overdraft_df)} overdraft entries written."
        )

    except Exception as e:
        messagebox.showwarning(
            "Loan Filter Warning",
            f"Could not filter loans: {e}",
        )
        raise


# ─────────────────────────────────────────────
# MAIN CONVERSION
# ─────────────────────────────────────────────
def convert_pdf_to_excel(pdf_path: str, excel_path: str) -> None:
    """Convert the selected M-Pesa PDF to a formatted Excel workbook."""
    df = extract_tables_with_pdfplumber(pdf_path)
    df = format_dates(df)

    df_numeric, money_cols = coerce_money_columns(df)

    # Ensure pivot-required columns are numeric.
    for expected in ['Paid In', 'Withdrawn', 'Balance']:
        actual = _find_column_case_insensitive(df_numeric, expected)
        if actual:
            df_numeric[actual] = _normalise_numeric(df_numeric[actual])
            if actual not in money_cols:
                money_cols.append(actual)

    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        # Full Data sheet.
        df_numeric.to_excel(writer, sheet_name='Full Data', index=False)

        workbook = writer.book
        full_data_ws = writer.sheets['Full Data']
        _format_money_columns(
            full_data_ws,
            workbook,
            df_numeric,
            money_cols,
        )

        create_pivot_table(df_numeric, writer)
        filter_loans(df_numeric, writer, money_cols)


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
def main() -> None:
    """Prompt for a PDF and create the corresponding Excel workbook."""
    root = Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Select M-Pesa PDF",
        filetypes=[("PDF Files", "*.pdf")],
    )
    if not pdf_path:
        sys.exit()

    folder = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    excel_path = os.path.join(folder, f"{base_name}.xlsx")

    unlocked_path = None

    try:
        answer = messagebox.askyesno(
            "PDF Password",
            "Is the PDF password-protected?",
        )

        if answer:
            password = simpledialog.askstring(
                "PDF Password",
                "Enter the M-Pesa PDF password:",
                show='*',
            )
            if not password:
                sys.exit()

            unlocked_path = os.path.join(folder, TEMP_UNLOCKED_PDF)
            pdf_to_use = unlock_pdf(pdf_path, unlocked_path, password)
        else:
            pdf_to_use = pdf_path

        convert_pdf_to_excel(pdf_to_use, excel_path)
        messagebox.showinfo(
            "Success",
            f"Excel file saved:\n{excel_path}",
        )

    except Exception as e:
        messagebox.showerror(
            "Error",
            f"An unexpected error occurred: {e}",
        )

    finally:
        if unlocked_path and os.path.exists(unlocked_path):
            try:
                os.remove(unlocked_path)
                print(f"[Cleanup] Removed temp file: {unlocked_path}")
            except OSError:
                pass


if __name__ == "__main__":
    main()
