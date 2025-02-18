# 🐍 XLSX to CSV Converter for XTB platform

This Python script automates the conversion of data from `.xlsx` files to `.csv` format, providing a streamlined solution for extracting, transforming, and saving structured data. It is particularly useful for processing financial reports or other tabular data stored in Excel files.

## 🔍 Overview

The script scans a designated folder (`xlsxFiles`) for `.xlsx` files, processes each file by extracting relevant data from a specific worksheet, and saves the formatted data as `.csv` files in an output folder (`csvFiles`). The conversion process includes handling specific formatting needs, such as replacing dots with commas in numeric values and mapping transaction types for better readability.

This tool is designed to save time and reduce errors in manual data processing, making it ideal for use in financial analysis, reporting tasks, or any workflow requiring clean data exports.

---

## ✨ Features

1. **Batch Processing**: Automatically processes all `.xlsx` files in the input folder, ensuring no file is missed.
2. **Dynamic Field Mapping**:
   - Extracts specific columns like dates, transaction types, amounts, and symbols.
   - Converts "BUY" and "SELL" transaction types to "Kupno" and "Sprzedaj" for localized formatting.
3. **Custom Formatting**:
   - Adjusts numeric fields to use commas as decimal separators.
   - Ensures consistent semicolon-delimited CSV files.
4. **Error Handling**:
   - Automatically creates the output folder (`csvFiles`) if it doesn’t exist.
   - Skips files that don’t match the `.xlsx` format.
5. **Readable Output**:
   - Outputs CSV files with UTF-8 encoding and a clear column structure.

---

## 🖥️ How to Use

1. **Prepare the Input Files**: Place all `.xlsx` files you want to process in the `xlsxFiles` directory.
2. **Install Dependencies**:
   - Make sure you have Python 3.8 or higher installed.
   - Install the `openpyxl` library using:
     ....
     pip install openpyxl
     ....
3. **Run the Script**: Execute the script by running:
     ....
     python converter.py
     ....
4. **Check the Output**: The processed `.csv` files will be saved in the `csvFiles` directory.

---

## 📋 Example Workflow

Imagine you have an `.xlsx` file with financial data stored in the `xlsxFiles` directory. The script will:

1. Open the file and locate the relevant worksheet (e.g., a sheet containing transaction details).
2. Extract data row by row until it encounters a "Total" marker in column B, which signals the end of the data.
3. Map extracted data into a structured format, such as:
   - Transaction date, type, value, symbol, and additional details.
   - Ensure consistent formatting for numeric fields and currency symbols.
4. Save the extracted data into a `.csv` file with a clear, structured layout.

For instance, if the input file contains the following data:

| B         | C       | D      | E  | F        | G     | I       | M      | Q    |  
|-----------|---------|--------|----|----------|-------|---------|--------|------|  
| Data 12   | Symbol1 | BUY    | 10 | 01-01-23 | 100.5 | 150.0   | 5.0    | Note |  
| Total     |         |        |    |          |       |         |        |      |

The output `.csv` will look like:

....
Data;Wartość;Uwaga;Podatki;Konto przeciwstawne;Waluta transakcji;Kurs wymiany;Kwota brutto;Symbol giełdowy waloru;ISIN;Opłaty;WKN;Akcje;Kwota waluty brutto;Nazwa waloru;Czas;Typ;Konto pieniężne;Konto walorów
01-01-23;100,5;Note;;;USD;;;;Symbol1;;5,0;;;10;150,0;;;;Kupno;Bank123;
....

---

## 🛡️ Error Handling

- **Missing Files**: If no `.xlsx` files are found, the script will notify you.
- **Invalid Format**: Files without the `.xlsx` extension are skipped.
- **Output Directory**: Automatically creates the output directory if it doesn’t exist.

---

## 🚀 Future Enhancements

- Add support for additional Excel formats, such as `.xls`.
- Include detailed logging for better debugging and traceability.
- Implement configuration options to customize column mappings and worksheet selection.
- Develop a graphical user interface (GUI) for easier use by non-technical users.

---

This script is a practical and efficient tool for automating repetitive data processing tasks. Contributions and feedback are welcome to improve its functionality and usability. 😊
