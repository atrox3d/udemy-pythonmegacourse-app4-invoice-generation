# import openpyxl
import pandas as pd
import glob

for filepath in glob.glob("invoices/*.xlsx"):
    print(filepath)
    df = pd.read_excel(filepath)
    print(df)
