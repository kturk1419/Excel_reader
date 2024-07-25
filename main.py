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
    date = Path(filepath).stem[6:]
    print(filename)
    pdf.add_page()
    pdf.set_font("Times", size=16)
    pdf.cell(w=0, h=12, txt=f"Invoice NO. {invno}", align="L", ln=1, border=0)

    pdf.cell(w=0, h=12, txt=f"Date {date}", align="L", ln=1, border=0)

    pdf.set_font("Times", size=10, style="B")
    cols = list(df.columns)
    cols = [item.replace("_", " ").title() for item in cols]

    pdf.cell(w=30, h=8, txt=cols[0], border=1)
    pdf.cell(w=70, h=8, txt=cols[1], border=1)
    pdf.cell(w=35, h=8, txt=cols[2], border=1)
    pdf.cell(w=30, h=8, txt=cols[3], border=1)
    pdf.cell(w=30, h=8, txt=cols[4], border=1, ln=1)

    pdf.set_font("Times", size=10)
    for i, row in df.iterrows():
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # total price
    total_price = df["total_price"].sum()

    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), border=1, ln=1)


    #add sentence:
    pdf.cell(w=30, h=8, txt=f"The total amount is ${total_price} SAR", ln=1)
    pdf.cell(w=35, h=8, txt=f"Khaled Turkestani CO")
    pdf.image("images/pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")



