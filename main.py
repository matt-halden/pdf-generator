from fpdf import FPDF
import pandas as pd
# Create standard list for excel files
import glob

# grab everything with .xlsx
filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    # read_excel not csv as this is an excel file, have to provide sheet_name
    # have to install py package: openpyxl for pandas to read excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
