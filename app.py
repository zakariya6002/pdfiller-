import pandas as pd
from pdfrw import PdfReader, PdfWriter, PageMerge

# Load Excel file
excel_file = "data.xlsx"  # Path to the Excel file
sheet_name = "Sheet1"  # Change if your data is in another sheet
data = pd.read_excel(excel_file, sheet_name=sheet_name)

# Load the PDF template
pdf_template = "fillable_form.pdf"  # Path to your PDF form
pdf_output = "filled_forms_combined.pdf"  # Final combined output PDF

# Read the PDF form
pdf_reader = PdfReader(pdf_template)

# Create a PDF writer for output
pdf_writer = PdfWriter()

# Loop through each row in the Excel file (each row is a new page)
for index, row in data.iterrows():
    # Clone the template for each record
    pdf_clone = PdfReader(pdf_template, decompress=True)
    
    # Get the first page (modify if your form has multiple pages)
    pdf_page = pdf_clone.pages[0]

    # Fill form fields with data from Excel
    annotations = pdf_page.Annots
    if annotations:
        for annotation in annotations:
            field_name = annotation.T and annotation.T[1:-1]  # Extract field name
            if field_name and field_name in data.columns:
                annotation.V = f"({row[field_name]})"  # Fill the field

    # Add the modified page to the PDF writer
    pdf_writer.addpage(pdf_page)

# Save all filled forms into a single PDF
pdf_writer.write(pdf_output)

print(f"All records filled and saved in {pdf_output}")
