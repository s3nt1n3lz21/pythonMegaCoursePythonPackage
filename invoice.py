import os

from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path


def get_title(key: str):
    match key:
        case 'product_id':
            return 'Product ID'
        case 'product_name':
            return 'Product Name'
        case 'amount_purchased':
            return 'Amount Purchased'
        case 'price_per_unit':
            return 'Price Per Unit'
        case 'total_price':
            return 'Total Price'
        case _:
            return ''

def get_column_width(key: str):
    match key:
        case "product_id":
            return 25
        case 'product_name':
            return 70
        case 'amount_purchased':
            return 40
        case 'price_per_unit':
            return 30
        case 'total_price':
            return 30
        case _:
            return 30

def generate(invoices_path, pdfs_path):
    """
    Convert invoices from Excel format into PDFs
    :param invoices_path:
    :param pdfs_path:
    :return:
    """

    filepaths = glob.glob(f"{invoices_path}/*.xlsx")

    for filepath in filepaths:
        df = pd.read_excel(filepath, sheet_name="Sheet 1")
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=False, margin=0)
        pdf.set_font(family="Times", style="B", size=18)
        pdf.set_text_color(100, 100, 100)
        page_height = 297
        page_width = 210

        filename = Path(filepath).stem
        invoice_number = filename.split("-")[0]
        date = filename.split("-")[1]

        # Create Initial page
        pdf.add_page()

        # Create title
        pdf.cell(w=0, h=12, txt="Invoice No: " + invoice_number, align="L", ln=1, border=0)
        pdf.cell(w=0, h=12, txt="Date: " + date, align="L", ln=1, border=0)

        # Create table
        line_height = pdf.font_size * 2.5
        col_width = (page_width - 40) / 5
        pdf.set_font(family="Times", style="B", size=10)

        # Create table headers
        pdf_y_initial = pdf.y
        pdf_x_initial = pdf.x
        pdf_y = pdf_y_initial
        pdf_x = pdf_x_initial
        for i, key in enumerate(df.keys()):
            pdf.set_y(pdf_y)
            pdf.set_x(pdf_x)
            pdf.multi_cell(
                get_column_width(key),
                line_height, get_title(key), border=1
            )
            pdf_x = pdf_x + get_column_width(key)

        # Create rows
        total = 0
        for i, row in df.iterrows():
            pdf_y = pdf.y
            pdf_x = pdf_x_initial
            for key in row.keys():
                pdf.set_y(pdf_y)
                pdf.set_x(pdf_x)
                pdf.multi_cell(get_column_width(key), line_height, str(row.get(key)), border=1)
                pdf_x = pdf_x + get_column_width(key)
                if key == "total_price":
                    total += row.get(key)

        # Create summary
        pdf.cell(w=0, h=12, txt="The total amount due is: " + str(total), align="L", ln=1, border=0)

        # Save file
        if not os.path.exists(pdfs_path):
            os.makedirs(pdfs_path)
        pdf_filepath = filename + '.pdf'
        pdf.output(f"{pdfs_path}/" + pdf_filepath)
