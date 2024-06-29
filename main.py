""" This script reads all the excel files in the invoices folder and creates a PDF file for each of them. """

import glob
from pathlib import Path
import pandas as pd
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename, date = Path(filepath).stem.split("-")
    # invoice number
    pdf.set_font("Arial", "B", 14)
    pdf.cell(190, 8, f"Invoice nr.{filename}", 0, 1, "R")
    # customer details
    pdf.set_font("Arial", "", 10)
    pdf.cell(190, 5, "Tomas Ferreras", 0, 1, "R")
    pdf.cell(190, 5, "4422 Adams Avenue", 0, 1, "R")
    pdf.cell(190, 5, "Spearfish, South Dakota", 0, 1, "R")
    pdf.cell(190, 5, "57783", 0, 1, "R")
    pdf.cell(190, 5, "United States", 0, 1, "R")
    pdf.set_font("Arial", "B", 10)
    pdf.cell(190, 5, f"{date}", 0, 1, "R")
    pdf.line(
        10,
        52,
        200,
        52,
    )
    pdf.ln(10)

    # invoice details
    columns = list(df.columns)
    pdf.set_font("Arial", "B", 8)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(30, 8, columns[0], border=1, fill=True)
    pdf.cell(70, 8, columns[1], border=1, fill=True)
    pdf.cell(30, 8, columns[2], border=1, fill=True)
    pdf.cell(
        30,
        8,
        columns[3],
        border=1,
        fill=True,
    )
    pdf.cell(30, 8, columns[4], border=1, fill=True, ln=1)
    for index, row in df.iterrows():
        pdf.set_text_color(80, 80, 80)
        pdf.cell(30, 8, str(row["product_id"]), border=1)
        pdf.cell(70, 8, str(row["product_name"]), border=1)
        pdf.cell(30, 8, str(row["amount_purchased"]), border=1)
        pdf.cell(30, 8, f'${str(row["price_per_unit"])}', border=1)
        pdf.cell(30, 8, f'${str(row["total_price"])}', border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_text_color(80, 80, 80)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(70, 8, "", border=1)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(30, 8, f"${total_sum}", border=1, ln=1)
    pdf.ln(10)

    pdf.image("images/amazon_icon.png", w=10)
    pdf.text(10, 10, "Amazon Inc.")

    pdf.output(f"PDFs/{filename}.pdf")
