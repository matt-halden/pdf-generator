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
    # read_excel not csv as this is an excel file, have to provide sheet_name
    # have to install py package: openpyxl for pandas to read excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    pdf = FPDF(orientation="P", unit='mm', format="A4")
    pdf.add_page()
    # get filenames, Path gets us a special string, then using stem grabs just the name
    filename = Path(filepath).stem
    # just grab initial invoice number, split makes the full filename into a list
    invoice_nr = filename.split("-")[0]
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice #: {invoice_nr}")
    pdf.output(f"PDFs/{filename}.pdf")
