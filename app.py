import streamlit as st
import pandas as pd
import io
import re
import os
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import cm
from reportlab.lib.colors import Color, yellow, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Türkçe karakter desteği için font ayarı
# NOT: GitHub repo'suna 'arial.ttf' veya 'DejaVuSans.ttf' dosyasını eklerseniz kusursuz çalışır.
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
st.markdown("*(Trendyol 10x10 cm | Hepsiburada 10x15 cm)*")

st.sidebar.header("⚙️ Personel Ayarları")
personel_input = st.sidebar.text_area(
    "Personel İsimleri (Her satıra bir isim yazın)", 
    "Ahmet\nAyşe\nMehmet"
)
personel_list = [p.strip() for p in personel_input.split('\n') if p.strip()]

def generate_barcode_image(code):
    """Verilen koddan Code128 barkod görseli oluşturur"""
    clean_code = re.sub(r'[^A-Za-z0-9]', '', str(code))
    if not clean_code:
        clean_code = "000000000000"
    
    code128 = Code128(clean_code, writer=ImageWriter())
    buffer = io.BytesIO()
    code128.write(buffer, options={"module_width": 0.2, "module_height": 10.0, "font_size": 8})
    return buffer

def process_trendyol_excel(uploaded_file):
    """Trendyol Excel'ini okur, çoklu ürünleri birleştirir ve 10x10 cm PDF üretir"""
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    
    # Sipariş Koduna göre grupla (Çoklu ürünleri tek etikette toplar)
    grouped = df.groupby('Sipariş Kodu')
    
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(10*cm, 10*cm))
    
    for order_code, group in grouped:
        barcode_val = str(group['Kampanya Kodu'].iloc[0]) if 'Kampanya Kodu' in group.columns and pd.notna(group['Kampanya Kodu'].iloc[0]) else str(order_code)
        
        customer = str(group['Sipariş Veren Cari'].iloc[0]) if 'Sipariş Veren Cari' in group.columns else "Müşteri"
        address = str(group['Alıcı Adres'].iloc[0]) if 'Alıcı Adres' in group.columns else ""
        city = str(group['Şehir/Semt/PK'].iloc[0]) if 'Şehir/Semt/PK' in group.columns else ""
        
        products = []
        for _, row in group.iterrows():
            prod_name = str(row['Sipariş Verilen Ürün(ler)']).split(' x')[0] if 'Sipariş Verilen Ürün(ler)' in row else "Ürün"
            qty = int(row['Adet']) if 'Adet' in row and pd.notna(row['Adet']) else 1
            products.append(f"{qty}x {prod_name}")
        
        # --- 10x10 cm ÇİZİM ---
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
            
        c.showPage()
    
    c.save()
    return buffer.getvalue(), len(grouped)

def process_hepsiburada_pdf(uploaded_file):
    """Hepsiburada PDF'ini okur, personel atar ve 10x15 cm PDF üretir"""
    orders = []
    with pdfplumber.open(uploaded_file) as pdf:
        text = ""
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    
    # "SİPARİŞ KODU:" ifadesine göre parçala
    blocks = text.split("SİPARİŞ KODU:")
    for block in blocks[1:]:
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        order_data = {"order_code": "", "customer": "", "address": "", "products": []}
        
        for i, line in enumerate(lines):
            if "SİPARİŞ TARİHİ:" in line:
                order_data["order_code"] = line.split("-")[0].replace("SİPARİŞ KODU:", "").strip()
            elif "ALICI BİLGİLERİ:" in line:
                order_data["customer"] = line.replace("ALICI BİLGİLERİ:", "").strip()
            elif "Adres:" in line and "Telefon:" in line:
                parts = line.split("Telefon:")
                order_data["address"] = parts[0].replace("Adres:", "").strip()
            elif "ÜRÜN KODU/ ADI" in line or "ADET" in line:
                continue
            elif re.match(r'^[A-Za-z0-9].*\/.*', line) and len(line) > 10 and "ROYALGROSS" not in line:
                parts = line.split('/')
                if len(parts) >= 2:
                    prod_name = parts[1].strip()
                    qty_match = re.search(r'(\d+)\s*$', line)
                    qty = qty_match.group(1) if qty_match else "1"
                    # Aynı ürünü tekrar ekleme
                    prod_str = f"{qty}x {prod_name}"
                    if prod_str not in order_data["products"]:
                        order_data["products"].append(prod_str)
        
        if order_data["order_code"]:
            orders.append(order_data)

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(10*cm, 15*cm))
    
    for i, order in enumerate(orders):
        assigned_person = personel_list[i % len(personel_list)] if personel_list else "Atanmadı"
        barcode_val = order["order_code"].replace("-", "").strip()
        
        # --- 10x15 cm ÇİZİM ---
        barcode_img = generate_barcode_image(barcode_val)
        c.drawImage(barcode_img, 1*cm, 12*cm, width=8*cm, height=2*cm)
        
        c.setFont(FONT_NAME, 10)
        c.drawString(1*cm, 11.5*cm, f"Sipariş: {order['order_code']}")
        c.drawString(1*cm, 11*cm, f"Müşteri: {order['customer']}")
        
        addr_text = order['address']
        if len(addr_text) > 50:
            c.drawString(1*cm, 10.5*cm, addr_text[:47] + "...")
            c.drawString(1*cm, 10*cm, addr_text[47:])
        else:
            c.drawString(1*cm, 10.5*cm, addr_text)
            
        c.setFont(FONT_NAME, 8)
        y_pos = 9.5*cm
        for prod in order["products"]:
            c.drawString(1*cm, y_pos, prod)
            y_pos -= 0.4*cm
            
        # Personel Damgası (Sarı Kutu)
        c.setFillColor(yellow)
        c.rect(6*cm, 1*cm, 3.5*cm, 1.5*cm, fill=1, stroke=1)
        c.setFillColor(black)
        c.setFont(FONT_NAME, 10)
        c.drawCentredString(7.75*cm, 1.75*cm, "PERSONEL:")
        c.setFont(FONT_NAME, 12)
        c.drawCentredString(7.75*cm, 1.25*cm, assigned_person)
        
        c.showPage()
    
    c.save()
    return buffer.getvalue(), len(orders)

# --- ARAYÜZ ---
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("📦 Trendyol İşlemleri (10x10 cm)")
    st.write("Sentos Excel çıktısını yükleyin. Aynı siparişteki ürünler otomatik tek etikette birleştirilir.")
    trendyol_file = st.file_uploader("Trendyol Excel (.xlsx)", type=["xlsx"])
    if st.button("Trendyol Etiketlerini Oluştur", type="primary", use_container_width=True):
        if trendyol_file:
            with st.spinner("İşleniyor, lütfen bekleyin..."):
                try:
                    pdf_bytes, count = process_trendyol_excel(trendyol_file)
                    st.success(f"✅ {count} adet sipariş için etiket oluşturuldu!")
                    st.download_button("📥 PDF İndir", data=pdf_bytes, file_name="Trendyol_Etiketler.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Hata oluştu: {str(e)}")
        else:
            st.warning("Lütfen bir Excel dosyası yükleyin.")

with col2:
    st.subheader("📦 Hepsiburada İşlemleri (10x15 cm)")
    st.write("Hepsiburada PDF etiketlerini yükleyin. Personel dağılımı otomatik ve dengeli yapılır.")
    hb_file = st.file_uploader("Hepsiburada PDF (.pdf)", type=["pdf"])
    if st.button("Hepsiburada Etiketlerini Oluştur", type="primary", use_container_width=True):
        if hb_file:
            with st.spinner("İşleniyor, lütfen bekleyin..."):
                try:
                    pdf_bytes, count = process_hepsiburada_pdf(hb_file)
                    st.success(f"✅ {count} adet sipariş için etiket oluşturuldu!")
                    st.download_button("📥 PDF İndir", data=pdf_bytes, file_name="Hepsiburada_Etiketler.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Hata oluştu: {str(e)}")
        else:
            st.warning("Lütfen bir PDF dosyası yükleyin.")

st.markdown("---")
st.caption("💡 **Not:** Türkçe karakterlerin (ş, ı, ğ, ü, ö, ç) kusursuz görünmesi için proje klasörüne `arial.ttf` dosyasını eklemeyi unutmayın.")