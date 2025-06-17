import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("Invocies/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation='p', unit='mm', format='A4')

    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')

    pdf.set_font(family="Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date{date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    ##Tablodaki basliklari aliyoruz, tablodaki basliklari tek bir kere kullanacagimiz icin
    # yani basliklar sabit oldugu icin bu islemi asagidaki verileri aldigimiz  for loopun disinda yapiyoruz

    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=8, style='B')
    pdf.set_text_color(80, 80, 80)

    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    ## tablodaki verileri almak icin olusturdugumuz itteration
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)
        ## tablonun ilk icerigi olan product_id'deki verileri
        # ##kismini aldik simdi tablonun devamini alma vakti
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    #toplam fiyati ekleme
    pdf.set_font(family="Times", size=14, style='B')
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    ## Sirket adi ekleme
    pdf.set_font(family="Times", size=14, style='B')
    pdf.cell(w=30, h=8, txt=f"SeverCan.com", ln=1)

    pdf.output(f"PDFs/{invoice_nr}.pdf")