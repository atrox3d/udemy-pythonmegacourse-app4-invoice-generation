# import openpyxl
import pandas as pd
import glob
from pathlib import Path
from fpdf import FPDF


for filepath in glob.glob("invoices/*.xlsx"):
    print(filepath)
    df = pd.read_excel(filepath)
    print(df)

    filename = Path(filepath).stem
    invoice, date = filename.split('-')
    print(invoice, date)

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    pdf.write(h=16, txt=f"Invoice nr.{invoice}")

    outfile = Path(f"pdf/{filename}.pdf")
    outdir = outfile.parent
    if not outdir.exists():
        outdir.mkdir()
    pdf.output(outfile)