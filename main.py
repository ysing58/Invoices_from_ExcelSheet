import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
import time

filepaths = glob.glob("invoices/*.xlsx")
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="L", format="A4", unit="mm")
    pdf.add_page()

    Invoice_nr, date = Path(filepath).stem.split("-")
    pdf.set_font(family="Times", style="B", size=16)

    pdf.cell(w=30, h=10, txt=f"Invoice nr.{Invoice_nr} ", align="L", ln=1)
    pdf.cell(w=30, h=10, txt=f"Date {date}", align="L", ln=1)

    total_price=0

    columns= [items.replace("_", " ").title() for items in df.columns]

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=30, h=10, txt=f"{columns[0]}", align="L", ln=0, border=1)
    pdf.cell(w=90, h=10, txt=f"{columns[1]}", align="L", ln=0, border=1)
    pdf.cell(w=50, h=10, txt=f"{columns[2]}", align="L", ln=0, border=1)
    pdf.cell(w=50, h=10, txt=f"{columns[3]}", align="L", ln=0, border=1)
    pdf.cell(w=30, h=10, txt=f"total_price", align="L", ln=1, border=1)


    for index, row in df.iterrows():
        sum_of_price = row['price_per_unit']*row['amount_purchased']
        pdf.set_font(family="Times", style="B", size=16)
        pdf.cell(w=30, h=10, txt=f"{row['product_id']}", align="L", ln=0, border=1)
        pdf.cell(w=90, h=10, txt=f"{row['product_name']}", align="L", ln=0, border=1)
        pdf.cell(w=50, h=10, txt=f"{row['amount_purchased']}", align="L", ln=0, border=1)
        pdf.cell(w=50, h=10, txt=f"{row['price_per_unit']}", align="L", ln=0, border=1)
        pdf.cell(w=30, h=10, txt=f"{sum_of_price}", align="L", ln=1, border=1)
        total_price+=sum_of_price
    pdf.cell(w=30, h=10, txt=f"", align="L", ln=0, border=1)
    pdf.cell(w=90, h=10, txt=f"", align="L", ln=0, border=1)
    pdf.cell(w=50, h=10, txt=f"", align="L", ln=0, border=1)
    pdf.cell(w=50, h=10, txt=f"", align="L", ln=0, border=1)
    pdf.cell(w=30, h=10, txt=f"{total_price}", align="L", ln=1, border=1)
    pdf.ln(20)
    pdf.set_font(family="Times", style="B", size=13)
    pdf.cell(w=0, h=10, txt=f"The Total due amount is {total_price}", align="L", ln=1, border=0)
    pdf.cell(w=25, h=10, txt=f"Python How", align="L", ln=0, border=0)
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDF/{Invoice_nr}.pdf")