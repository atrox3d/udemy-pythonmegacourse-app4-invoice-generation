# import openpyxl
import pandas as pd
import glob
from pathlib import Path
from fpdf import FPDF


def pt2mm(points):
    return points * 0.35


for filepath in glob.glob("invoices/*.xlsx"):
    print(filepath)
    filename = Path(filepath).stem
    invoice, date = filename.split('-')
    print(invoice, date)

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", size=16, style="B")
    # pdf.write(h=8, txt=f"Invoice nr.{invoice}\n")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    # pdf.write(h=8, txt=f"Date: {date}\n")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath)

    # print headers dynamically
    headers = [header.replace('_', ' ').title() for header in df.columns]
    widths = [30, 70, 30, 30, 30]
    for header, width in zip(headers, widths):
        pdf.set_font(family="Times", size=10, style="B")
        pdf.set_text_color(80, 80, 80)
        print(header, width)
        pdf.cell(w=width, h=8, border=1, txt=header)
    pdf.ln()

    # print rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, border=1, txt=str(row.product_id))
        pdf.cell(w=70, h=8, border=1, txt=str(row.product_name))
        pdf.cell(w=30, h=8, border=1, txt=str(row.amount_purchased))
        pdf.cell(w=30, h=8, border=1, txt=str(row.amount_purchased))
        pdf.cell(w=30, h=8, border=1, txt=str(row.total_price))
        pdf.ln()


    # create PDF
    outfile = Path(f"pdf/{filename}.pdf")
    outdir = outfile.parent
    if not outdir.exists():
        outdir.mkdir()
    pdf.output(outfile)
