import zipfile
import glob
from fpdf import FPDF
from pathlib import Path

# extract txt files from zip
archivepath = "Text-Files.zip"
destdir = "."
with zipfile.ZipFile(archivepath, 'r') as archive:
    print(f"Extracting {archivepath}...")
    archive.extractall(destdir)

# create pdf
pdf = FPDF(orientation="P", unit="mm",  format="A4")
for textfile in glob.glob("*.txt"):
    # create one page for each txt file
    print(f"processing {textfile}...")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    title = Path(textfile).stem.capitalize()
    pdf.write(h=16, txt=title)

pdf.output("solution.pdf")



