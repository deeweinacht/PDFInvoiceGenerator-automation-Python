"""
This application generates formatted PDF invoices from Excel documents.
Copyright (c) 2023 Dee Weinacht
"""

import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob('sales/*.xlsx')

for filepath in filepaths:
    file = pd.read_excel(filepath, sheet_name='Sheet 1')
    filename = Path(filepath).stem
    invoice_num, invoice_date = filename.split('-')
    invoice_date = invoice_date.replace('.', '-')

    pdf = FPDF()
    pdf.add_page()  # A4 size 210 x 297 mm

    # add header
    pdf.set_font(family="Times", style='B', size=18)
    pdf.set_y(10)
    pdf.cell(w=100, h=10, txt=f'Invoice date: {invoice_date}')
    pdf.cell(w=80, h=10, txt=f'Invoice #{invoice_num}', align='R', ln=1)
    pdf.line(10, 20, 200, 20)
    pdf.set_y(40)

    # add titles to invoice
    titles = file.columns
    titles = [title.replace('_', ' ').title() for title in titles]

    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=80, h=8, txt=titles[1], border=1)
    pdf.cell(w=25, h=8, txt=titles[0], border=1)
    pdf.cell(w=20, h=8, txt=titles[2][:7], border=1)
    pdf.cell(w=30, h=8, txt=titles[3], border=1)
    pdf.cell(w=25, h=8, txt=titles[4], border=1, ln=1)

    # add rows to invoice
    for i, row in file.iterrows():
        pdf.set_font(family='Times', size=12)
        pdf.cell(w=80, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=25, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=20, h=8, txt=str(row['amount_purchased']),
                 border=1, align='R')
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']),
                 border=1, align='R')
        pdf.cell(w=25, h=8, txt=str(row['total_price']), border=1,
                 ln=1, align='R')

    # add grand total row to invoice
    total_price = file['total_price'].sum()
    pdf.set_x(135)
    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=30, h=8, txt='Grand Total:', border=1)
    pdf.cell(w=25, h=8, txt=f'${total_price}', border=1, align='R', ln=1)
    pdf.cell(w=0, h=20, txt='', border=0, ln=1)

    # Invoice closing
    pdf.set_font(family='Times', size=12)
    pdf.cell(w=0, h=12, txt='Thank you for shopping with us!', ln=1)
    pdf.cell(w=70, h=12,
             txt=f'The total price for your order is: ${total_price}.', ln=0)
    pdf.cell(w=0, h=12,
             txt='Please provide your remittance as soon as possible', ln=1)
    pdf.cell(w=0, h=12,
             txt=f'Please contact us if you need any further assistance', ln=1)
    pdf.cell(w=0, h=20, txt='', border=0, ln=1)
    pdf.cell(w=0, h=12, txt=f'-The <YourBusiness> Team', ln=1)
    pdf.image('images/example logo.png', w=50)

    # save file
    pdf.output(f'invoices/{filename}.pdf')