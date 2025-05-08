import datetime
import xlwings as xw
import os
import inspect
import pikepdf

save_directory = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
support_directory = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))) + "\\Support"
workbooks = []

for file in os.listdir(support_directory):
    file_path = os.path.join(support_directory, file)
    if file_path.endswith(".xlsm"):
        workbooks.append(file_path)

print("Merging PDF's...")

# Process each workbook
for workbook_path in workbooks:
    try:
        # Open the workbook
        app = xw.App(visible=False)  # Run Excel in the background
        wb = app.books.open(workbook_path)
        
        # Save the resulting PDF
        pdf_name = os.path.splitext(os.path.basename(workbook_path))[0] + ".pdf"
        pdf_path = os.path.join(save_directory, pdf_name)
        wb.api.ExportAsFixedFormat(0, pdf_path)  # 0 corresponds to PDF format
        
        print(f"Saving PDF: {pdf_path}")
        
        # Close the workbook
        wb.close()
        app.quit()

        # Append the newly created PDF with corresponding pre-existing PDFs
        five_digit_code = pdf_name.split(" ")[3][:5]  # Extract the five-digit code
        year_letter_combo = pdf_name.split(" ")[1] + " " + pdf_name.split(" ")[2][:1]  # Extract the year/letter combination

        # Find matching PDFs in the support directory
        matching_pdfs = [os.path.join(support_directory, f) for f in os.listdir(support_directory) 
                            if (f.startswith(five_digit_code) or year_letter_combo in f) and f.lower().endswith('.pdf')]

        if matching_pdfs:
            merged_pdf_path = os.path.join(save_directory, f"{pdf_name}")
            
            # Open the newly created PDF
            with pikepdf.Pdf.open(pdf_path, allow_overwriting_input=True) as merged_pdf:
                for matching_pdf in matching_pdfs:
                    with pikepdf.Pdf.open(matching_pdf) as pdf_to_append:
                        merged_pdf.pages.extend(pdf_to_append.pages)
                
                # Save the merged PDF
                merged_pdf.save(merged_pdf_path)
    except Exception as e:
        print(e)
        try:
            print("Second attempt...")
            merged_pdf.save(merged_pdf_path)
        except Exception as e:
            print("Failed to save merged PDF after second attempt.")