import os
from openpyxl import load_workbook
import csv

# Input and output directories
input_directory = "xlsxFiles"
output_directory = "csvFiles"

# Ensure the output directory exists
os.makedirs(output_directory, exist_ok=True)

# Iterate over all files in the input directory
for filename in os.listdir(input_directory):
    if filename.endswith(".xlsx"):
        # Full path to the Excel file
        excel_path = os.path.join(input_directory, filename)

        # Load the workbook
        workbook = load_workbook(excel_path)

        # Select the sheet (adjust if needed to select the correct sheet)
        sheet = workbook.worksheets[1]
        print(f"Processing file: {filename}, sheet: {sheet.title}")

        # Read the currency value
        currency = sheet["L6"].value

        # Initialize the open position list
        openPositionList = []

        # Read rows until "Total" is encountered in column B
        row_index = 12
        while sheet[f"B{row_index}"].value != "Total":
            openPositionList.append({
                "Data": sheet[f"F{row_index}"].value.strftime('%Y-%m-%d %H:%M:%S'),
                "Wartość": str(sheet[f"I{row_index}"].value).replace(".", ","),
                "Uwaga": sheet[f"Q{row_index}"].value,
                "Podatki": "",
                "Konto przeciwstawne": "",
                "Waluta transakcji": currency,
                "Kurs wymiany": "",
                "Kwota brutto": "",
                "Symbol giełdowy waloru": sheet[f"C{row_index}"].value,
                "ISIN": "",
                "Opłaty": str(sheet[f"M{row_index}"].value).replace(".", ","),
                "WKN": "",
                "Akcje": str(sheet[f"E{row_index}"].value).replace(".", ","),
                "Kwota waluty brutto": str(sheet[f"I{row_index}"].value).replace(".", ","),
                "Nazwa waloru": "",
                "Czas": "",
                "Typ": sheet[f"D{row_index}"].value,
                "Konto pieniężne": sheet["I6"].value,
                "Konto walorów": ""
            })
            row_index += 1

        # displaying the openPositionList
        for i in openPositionList:
            print(i)

        # Convert the "Typ" column to "BUY" and "SELL"
            for i in openPositionList:
                if i["Typ"] == "BUY":
                    i["Typ"] = "Kupno"
                elif i["Typ"] == "SELL":
                    i["Typ"] = "Sprzedaj"

        # Generate the output CSV filename
        csv_filename = os.path.splitext(filename)[0] + ".csv"
        csv_path = os.path.join(output_directory, csv_filename)

        # Save the data to a CSV file
        with open(csv_path, 'w', newline='', encoding="utf-8") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=openPositionList[0].keys(), delimiter=';')
            writer.writeheader()
            for row in openPositionList:
                writer.writerow(row)

        print(f"Saved to: {csv_path}")

print("Processing complete.")
