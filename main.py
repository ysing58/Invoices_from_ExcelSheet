import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="L", format="A4", unit="mm")
    pdf.add_page()
    Invoice_nr = Path(filepath).stem.split("-")[0]
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=30, h=10, txt=f"Invoice nr.{Invoice_nr} ", align="L", ln=1)
    pdf.output(f"PDF/{Invoice_nr}.pdf")