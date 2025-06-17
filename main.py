import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# "Invocies" klasörü altındaki tüm .xlsx dosyalarının dosya yollarını al
filepaths = glob.glob("Invocies/*.xlsx")

# Her bir Excel dosyası için işlem başlat
for filepath in filepaths:
    # A4 boyutunda, dikey (portrait) bir PDF nesnesi oluştur
    pdf = FPDF(orientation='p', unit='mm', format='A4')

    # Yeni bir sayfa ekle
    pdf.add_page()

    # Dosya adını uzantısız olarak al (örn. "10001-2023.1.18")
    filename = Path(filepath).stem

    # Dosya adını fatura numarası ve tarih olarak ayır ("10001", "2023.1.18")
    invoice_nr, date = filename.split('-')

    # Fatura numarasını PDF'e yaz (kalın yazı tipi ile)
    pdf.set_font(family="Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    # Tarihi PDF'e yaz
    pdf.set_font(family="Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date{date}", ln=1)

    # Excel dosyasındaki "Sheet 1" sayfasını oku
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Excel tablosundaki sütun başlıklarını al ve formatla (örnek: product_id → Product Id)
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]

    # Başlık satırını PDF'e yaz (sütun isimleri)
    pdf.set_font(family="Times", size=8, style='B')
    pdf.set_text_color(80, 80, 80)

    # Başlık hücrelerini yaz
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Excel tablosundaki her satır için PDF'e veri yaz
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)

        # Her hücreyi ilgili genişlikte ve kenarlıkla yaz
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Tablonun en altına toplam fiyatı yazmak için toplam satırını oluştur
    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Toplam fiyat bilgisini daha büyük yazı ile belirt
    pdf.set_font(family="Times", size=14, style='B')
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Şirket ismini PDF'e ekle
    pdf.set_font(family="Times", size=14, style='B')
    pdf.cell(w=30, h=8, txt=f"SeverCan.com", ln=1)

    # PDF dosyasını "PDFs" klasörüne fatura numarasıyla kaydet
    pdf.output(f"PDFs/{invoice_nr}.pdf")
