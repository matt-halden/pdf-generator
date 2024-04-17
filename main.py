import pandas as pd
# Create standard list for excel files with glob
import glob
from fpdf import FPDF
# pathlib allows us to extract file name
from pathlib import Path

# grab everything with .xlsx
filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit='mm', format="A4")
    pdf.add_page()

    # get filenames, Path gets us a special string, then using stem grabs just the name
    filename = Path(filepath).stem
    # just grab initial invoice number, split makes the full filename into a list
    invoice_nr = filename.split("-")[0]
    # grab date, extract second item from list. could do invoice_nr, date = filename.split("-")!!!
    date = filename.split("-")[1]

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice #: {invoice_nr}", ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=2)

    # read_excel not csv as this is an excel file, have to provide sheet_name
    # have to install py package: openpyxl for pandas to read excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # grab headers from excel file to make headers for table in PDF
    columns = df.columns

    # Clean up column headers
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", style="B", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1)

    # Add Rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)

        # create cells for each cell in table, no break line so they go next to each other
        # make these rows strings as row is actually giving us an integer output
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
