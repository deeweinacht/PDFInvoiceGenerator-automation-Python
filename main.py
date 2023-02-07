import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob('sales/*.xlsx')
for filepath in filepaths:
    file = pd.read_excel(filepath, sheet_name='Sheet 1')
    filename = Path(filepath).stem
    invoice_num = filename.split('-')[0]

    pdf = FPDF()
    pdf.add_page()

    # add header
    pdf.set_font(family="Times", style='B', size=16)
    pdf.cell(w=0, h=16, txt=f'Invoice #{invoice_num}')
    pdf.output(f'invoices/{filename}.pdf')