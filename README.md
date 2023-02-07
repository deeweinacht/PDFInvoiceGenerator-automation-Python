# PDFInvoiceGenerator-automation-Python

This application generates formatted PDF invoices from Excel documents.

    Copyright (C) 2023  Dee Weinacht

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.


**Using the app:**  

Insert the Excel documents you wish to generate invoices from into the 
'sales' folder. They should be in the same format as the example Excel documents
with the following column names: 'product_id', 'product_name', 'amount_purchased',
'price_per_unit', 'total_price'.
Run the script and a PDF invoice will be generated for each Excel document
in the 'sales' folder. The generated PDFs will be saved in the 'invoices'
folder with the same file names as the sales Excel documents.
    

**Dependencies:**
- pandas 1.4.3 licensed with BSD 3-Clause
- openpyxl 3.1.0 licensed with MIT/Expat
- pyfpdf 1.7.2 licensed with LGPL-3.0
