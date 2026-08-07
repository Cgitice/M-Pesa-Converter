# 📄 M-Pesa Statement Converter

A Python application that automates the conversion of Safaricom M-Pesa PDF statements into structured Excel reports for financial analysis and credit assessment.

The application extracts transaction data from M-Pesa statements, cleans and formats the information, generates summary pivot tables, and identifies transactions related to digital lenders, significantly reducing manual processing time.

---

## Overview

Financial institutions often receive M-Pesa statements in PDF format when assessing loan applications. Reviewing these statements manually is time-consuming and prone to errors.

This application streamlines the process by converting PDF statements into structured Excel workbooks containing:

- Complete transaction data
- Financial summary pivot tables
- Loan-related transaction reports
- Properly formatted dates and monetary values

The project was developed to improve efficiency in credit analysis workflows.

---

## Features

✔ Convert Safaricom M-Pesa PDF statements to Excel

✔ Support both password-protected and non-password-protected PDF statements

✔ Automatically extract transaction tables

✔ Clean and standardize transaction data

✔ Parse and format transaction dates

✔ Detect and format monetary values

✔ Generate six-month financial summary pivot tables

✔ Calculate:

- Average Monthly Paid In
- 70% Discounting
- Profitability @ 20%
- Disposable Income @ 25%

✔ Detect transactions from known Digital Credit Providers (DCPs)

✔ Export professionally formatted Excel reports

✔ Automatically delete temporary unlocked PDF files

---

## Excel Output

The generated workbook contains the following worksheets:

### 1. Full Data

Contains all extracted transactions from the PDF statement with cleaned and formatted data.

### 2. Pivot Table

Provides a six-month financial summary including:

- Total Paid In
- Total Withdrawn
- Average Balance
- Grand Total
- Average Monthly Paid In
- 70% Discounting
- Profitability @ 20%
- Disposable Income @ 25%

### 3. Loans

Lists transactions associated with registered Digital Credit Providers (DCPs) to support credit assessment.

---

## Technologies Used

- Python 3
- Pandas
- NumPy
- pdfplumber
- PikePDF
- XlsxWriter
- Tkinter

---

## Project Structure

```
R-M-Pesa-Converter
│
├── src/
│   └── converter.py
│
├── sample_data/
│
├── screenshots/
│
├── output/
│
├── tests/
│
├── docs/
│
├── README.md
├── requirements.txt
├── LICENSE
├── .gitignore
└── run.py
```

---

## Installation

Clone the repository:

```bash
git clone https://github.com/yourusername/R-M-Pesa-Converter.git
```

Move into the project folder:

```bash
cd R-M-Pesa-Converter
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it.

**Windows**

```bash
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Running the Application

Run:

```bash
python run.py
```

The application will:

1. Prompt you to select an M-Pesa PDF statement.
2. Ask whether the PDF is password protected.
3. Unlock the PDF (if required).
4. Extract transaction data.
5. Generate a formatted Excel report.
6. Remove any temporary files created during processing.

---

## Screenshots

### Input PDF

_Add screenshot here_

---

### Full Data Worksheet

_Add screenshot here_

---

### Pivot Table

_Add screenshot here_

---

### Loan Transactions

_Add screenshot here_

---

## Business Value

This application automates a repetitive manual process commonly performed during credit assessment.

Key benefits include:

- Reduced manual data entry
- Faster financial analysis
- Improved reporting consistency
- Automatic identification of loan transactions
- Standardized financial summaries

---

## Future Improvements

Potential enhancements include:

- Support for additional statement formats
- Interactive dashboard
- Batch processing of multiple statements
- Export to CSV
- Improved graphical user interface
- Loan interest calculator
- Credit scoring system

---

## Author

**Caroline Mwende Gitice**

Bachelor of Economics | Credit Risk Analyst | Data Analytics Enthusiast

GitHub: *(https://github.com/Cgitice)*

LinkedIn: *(www.linkedin.com/in/caroline-gitice)*

---

## License

This project is licensed under the MIT License.
