import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")

print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    filename = Path(filepath).stem
    invno = Path(filepath).stem[:5]
    print(filename)
    pdf.add_page()
    pdf.set_font("Times", size=16)
    pdf.cell(w=0, h=12, txt=f"Invoice NO. {invno}", align="L", ln=1, border=0)
    pdf.output(f"PDFs/{filename}.pdf")


