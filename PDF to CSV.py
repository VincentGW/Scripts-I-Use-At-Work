import os
import inspect
import csv
import datetime
import pandas as pd
from PyPDF2 import PdfReader

# Get today's date and calculate the previous month
today = datetime.date.today()
previous_month = (today.replace(day=1) - datetime.timedelta(days=1)).strftime('%b')
directory = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))

fields = ['Date', 'Desc', 'Daily Factor', 'Transaction Amount', 'Outstanding Balance']

# Iterate through each folder in the directory
for folder_name in os.listdir(directory):
    folder_path = os.path.join(directory, folder_name)
    if os.path.isdir(folder_path):
        # Search for a PDF file containing the previous month's three-letter code
        for file_name in os.listdir(folder_path):
            if file_name.endswith('.pdf') and previous_month in file_name:
                pdf_path = os.path.join(folder_path, file_name)
                reader = PdfReader(pdf_path)

                data = []
                lines = []

                # Extract text from the PDF
                for page in reader.pages:
                    text = page.extract_text()
                    page_lines = text.splitlines()
                    lines.extend(page_lines)  # Append the lines from the current page to the main list

                # Process the extracted lines
                for line in lines:
                    if line[:2].isdigit() and '/' in line[:10]:  # Basic check for date format
                        parts = line.split()
                        if len(parts) >= 4:
                            date = parts[0]
                            desc = parts[1]
                            daily_factor = parts[-3] if parts[-3].startswith('0.') else 0
                            transaction_amount = parts[-2]
                            outstanding_balance = parts[-1]
                            data.append([date, desc, daily_factor, transaction_amount, outstanding_balance])

                # Create a DataFrame from the extracted data
                df = pd.DataFrame(data, columns=fields)
                selected_columns = df[['Date', 'Desc', 'Transaction Amount']]

                # Save the selected columns to a CSV file
                output_csv_name = f"{folder_name} {previous_month} Output.csv"
                output_csv_path = os.path.join(directory, output_csv_name)
                selected_columns.to_csv(output_csv_path, index=False)

                print(f"Processed {file_name} in folder {folder_name}. CSV saved to {output_csv_path}.")
