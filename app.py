import streamlit as st
import pandas as pd
import io
import re
import os
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.colors import yellow, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pypdf

FONT_NAME = 'Helvetica'
if os.path.exists('arial.ttf'):
    try:
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
        FONT_NAME = 'Arial'
    except Exception:
        pass
elif os.path.exists('DejaVuSans.ttf'):
    try:
        pdfmetrics.registerFont(TTFont('DejaVu', 'DejaVuSans.ttf'))
        FONT_NAME = 'DejaVu'
    except Exception:
        pass

st.set_page_config(page_title="ROYALGROSS Etiket Otomasyonu", layout="wide", page_icon="📦")
st.title("📦 ROYALGROSS Günlük Etiket ve Barkod Otomasyonu")
st.markdown("*(Trendyol 10x10 cm | Hepsiburada Orijinal Barkod Korunur + Personel Damgası)*")
st.markdown("---")

st.sidebar.header("⚙️ Personel Ayarları")
personel_input = st.sidebar.text_area(
    "Personel İsimleri veya Kodları (Her satıra bir tane)",
    "Ahmet (Kod: 01)\nAyşe (Kod: 02)\nMehmet (Kod: 03)"
)
personel_list = [p.strip() for p in personel_input.split('\n') if p.strip()]

st.sidebar.info("🖨️ **Yazıcı Boyutları:**\n- Trendyol: 10x10 cm\n- Hepsiburada: 10x15 cm")
st.sidebar.markdown("---")
st.sidebar.caption("💡 Türkçe karakterler için `arial.ttf` dosyasını repo'ya ekleyin.")

def generate_barcode_image(code):
    clean_code = re.sub(r'[^A-Za-z0-9]', '', str(code))
    if not clean_code:
        clean_code = "000000000000"
    code128 = Code128(clean_code, writer=ImageWriter())
    buffer = io.BytesIO()
    code128.write(buffer, options={"module_width": 0.2, "module_height": 10.0, "font_size": 8})
    return buffer

def process_trendyol_excel(uploaded_file, personel_list):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df.columns = df.columns.str.strip()
        
        grouped = df.groupby('Sipariş Kodu')
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=(10*cm, 10*cm))
        
        for idx, (order_code, group) in enumerate(grouped):
            assigned_person = personel_list[idx % len(personel_list)] if personel_list else "Atanmadı"
            barcode_val = str(group['Kampanya Kodu'].iloc[0]) if 'Kampanya Kodu' in group.columns and pd.notna(group['Kampanya Kodu'].iloc[0]) else str(order_code)
            
            customer = str(group['Sipariş Veren Cari'].iloc[0]) if 'Sipariş Veren Cari' in group.columns else "Müşteri"
            address = str(group['Alıcı Adres'].iloc[0]) if 'Alıcı Adres' in group.columns else ""
            city = str(group['Şehir/Semt/PK'].iloc[0]) if 'Şehir/Semt/PK' in group.columns else ""
            
            products = []
            for _, row in group.iterrows():
                prod_name = str(row['Sipariş Verilen Ürün(ler)']).split(' x')[0] if 'Sipariş Verilen Ürün(ler)' in row else "Ürün"
                qty = int(row['Adet']) if 'Adet' in row and pd.notna(row['Adet']) else 1
                products.append(f"{qty}x {prod_name}")
            
            barcode_img = generate_barcode_image(barcode_val)
            c.drawImage(barcode_img, 1*cm, 7.5*cm, width=8*cm, height=2*cm)
            c.setFont(FONT_NAME, 10)
            c.drawString(1*cm, 7*cm, f"Sipariş: {order_code}")
            c.setFont(FONT_NAME, 9)
            c.drawString(1*cm, 6.5*cm, f"Müşteri: {customer}")
            full_address = f"{address} {city}"
            if len(full_address) > 55:
                c.drawString(1*cm, 6*cm, full_address[:52] + "...")
            else:
                c.drawString(1*cm, 6*cm, full_address)
            c.setFont(FONT_NAME, 8)
            y_pos = 5.5*cm
            for prod in products:
                if len(prod) > 45:
                    c.drawString(1*cm, y_pos, prod[:42] + "...")
                    y_pos -= 0.4*cm
                    c.drawString(1.5*cm, y_pos, prod[42:])
                else:
                    c.drawString(1*cm, y_pos, prod)
                y_pos -= 0.5*cm
            c.setFillColor(yellow)
            c.rect(1*cm, 0.5*cm, 8*cm, 1.5*cm, fill=1, stroke=1)
            c.setFillColor(black)
            c.setFont(FONT_NAME, 9)
            c.drawCentredString(5*cm, 1.6*cm, "DEPO PERSONELİ:")
            c.setFont(FONT_NAME, 12)
            c.drawCentredString(5*cm, 0.9*cm, assigned_person)
            c.showPage()
        c.save()
        return buffer.getvalue(), len(grouped)
    except Exception as e:
        raise Exception(f"Excel okuma hatası: {str(e)}")

def extract_product_from_text(text):
    """Hepsiburada PDF'inden ürün adını çıkar"""
    lines = text.split('\n')
    for line in lines:
        if 'ÜRÜN KODU/ ADI' in line and 'ADET' in line:
            idx = lines.index(line)
            if idx + 1 < len(lines):
                product_line = lines[idx + 1].strip()
                if product_line and '/' in product_line:
                    parts = product_line.split('/')
                    if len(parts) >= 2:
                        return parts[1].strip()
    return "Diger"

def process_hepsiburada_pdf(uploaded_file, personel_list):
    reader = pypdf.PdfReader(uploaded_file)
    writer = pypdf.PdfWriter()
    
    page_info = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        order_match = re.search(r'SİPARİŞ KODU:\s*([0-9\-]+)', text)
        order_code = order_match.group(1).strip() if order_match else f"Sayfa_{i+1}"
        product_name = extract_product_from_text(text)
        page_info.append({
            "index": i,
            "product": product_name,
            "order": order_code,
            "page": page
        })
    
    # 1. ÖNCE ÜRÜN BAZINDA SIRALA
    page_info.sort(key=lambda x: (x["product"].lower(), x["order"]))
    
    # 2. SONRA PERSONELİ SIRALI DAĞIT (aynı ürünler aynı personele gelsin)
    personel_index = 0
    current_product = None
    
    for idx, info in enumerate(page_info):
        # Eğer ürün değiştiyse, bir sonraki personele geç
        if info["product"] != current_product:
            current_product = info["product"]
            # Personel formatını düzenle: "Ahmet (Kod: 01)" -> "DEPO[01]"
            if personel_list:
                personel_entry = personel_list[personel_index % len(personel_list)]
                # Kodu çıkar: "Ahmet (Kod: 01)" -> "01"
                code_match = re.search(r'Kod:\s*(\d+)', personel_entry)
                if code_match:
                    assigned_person = f"DEPO[{code_match.group(1)}]"
                else:
                    assigned_person = f"DEPO[{personel_index % len(personel_list) + 1:02d}]"
            else:
                assigned_person = "DEPO[00]"
            personel_index += 1
        else:
            # Aynı ürün, aynı personel
            if personel_list:
                prev_personel_entry = personel_list[(personel_index - 1) % len(personel_list)]
                code_match = re.search(r'Kod:\s*(\d+)', prev_personel_entry)
                if code_match:
                    assigned_person = f"DEPO[{code_match.group(1)}]"
                else:
                    assigned_person = f"DEPO[{(personel_index - 1) % len(personel_list) + 1:02d}]"
            else:
                assigned_person = "DEPO[00]"
        
        info["assigned_person"] = assigned_person
        page = info["page"]
        page_box = page.mediabox
        width = float(page_box.width)
        height = float(page_box.height)
        
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))
        c.setFillColor(yellow)
        c.rect(width - 140, 10, 130, 45, fill=1, stroke=1)
        c.setFillColor(black)
        c.setFont(FONT_NAME, 10)
        c.drawCentredString(width - 75, 40, "DEPO PERSONELİ:")
        c.setFont(FONT_NAME, 13)
        c.drawCentredString(width - 75, 20, assigned_person)
        c.save()
        
        packet.seek(0)
        stamp_pdf = pypdf.PdfReader(packet)
        page.merge_page(stamp_pdf.pages[0])
        writer.add_page(page)
    
    out_buffer = io.BytesIO()
    writer.write(out_buffer)
    return out_buffer.getvalue(), len(page_info)

st.subheader("📦 TRENDYOL İŞLEMLERİ (10x10 cm)")
st.write("Sentos Excel çıktısını yükleyin. Aynı siparişteki ürünler tek etikette birleştirilir, personel kodu eklenir.")
trendyol_file = st.file_uploader("Trendyol Excel dosyasını seçin (.xlsx)", type=["xlsx"], key="trendyol_uploader")
if trendyol_file is not None:
    st.success(f"✅ Dosya yüklendi: {trendyol_file.name}")
    if st.button("Trendyol Etiketlerini Oluştur", type="primary", use_container_width=True):
        with st.spinner("İşleniyor, lütfen bekleyin..."):
            try:
                pdf_bytes, count = process_trendyol_excel(trendyol_file, personel_list)
                st.success(f"✅ {count} adet sipariş için 10x10 cm etiket oluşturuldu!")
                st.download_button("Trendyol PDF'ini İndir", data=pdf_bytes, file_name="Trendyol_Etiketler.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ Hata oluştu: {str(e)}")
                st.info("💡 Lütfen Trendyol Excel dosyanızın formatını kontrol edin veya örnek bir dosya paylaşın.")
else:
    st.info("👆 Yukarıdan bir Excel dosyası seçin.")

st.markdown("---")

st.subheader("📦 HEPSİBURADA İŞLEMLERİ (10x15 cm)")
st.write("PDF yükleyin. Sistem barkodu BOZMAZ, sayfaları **ÜRÜN BAZINDA** sıralar ve **personeli gruplar** (DEPO[01] formatında).")
hb_file = st.file_uploader("Hepsiburada PDF dosyasını seçin (.pdf)", type=["pdf"], key="hepsiburada_uploader")
if hb_file is not None:
    st.success(f"✅ Dosya yüklendi: {hb_file.name}")
    if st.button("Hepsiburada Etiketlerini Oluştur", type="primary", use_container_width=True):
        with st.spinner("İşleniyor, lütfen bekleyin..."):
            try:
                pdf_bytes, count = process_hepsiburada_pdf(hb_file, personel_list)
                st.success(f"✅ {count} adet sayfa işlendi ve ürün bazında sıralandı!")
                st.download_button("Hepsiburada PDF'ini İndir", data=pdf_bytes, file_name="Hepsiburada_Etiketler.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ Hata oluştu: {str(e)}")
else:
    st.info("👆 Yukarıdan bir PDF dosyası seçin.")

st.markdown("---")
st.caption("🏭 ROYALGROSS EV GEREÇLERİ DIŞ TİCARET LİMİTED ŞİRKETİ - Otomatik Etiket Sistemi")
