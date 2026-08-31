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
import pypdf

# Türkçe karakter desteği için font ayarı
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

st.sidebar.header("⚙️ Personel Ayarları")
personel_input = st.sidebar.text_area(
    "Personel İsimleri veya Kodları (Her satıra bir tane)", 
    "Ahmet (Kod: 01)\nAyşe (Kod: 02)\nMehmet (Kod: 03)"
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

def process_trendyol_excel(uploaded_file, personel_list):
    """Trendyol Excel'ini okur, çoklu ürünleri birleştirir, personel atar ve 10x10 cm PDF üretir"""
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    
    # Sipariş Koduna göre grupla (Çoklu ürünleri tek etikette toplar)
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
            
        # --- PERSONEL DAMGASI (Trendyol) ---
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

def process_hepsiburada_pdf(uploaded_file, personel_list):
    """Hepsiburada PDF'ini okur, ÜRÜN BAZLI sıralar, ORİJİNAL barkodu korur ve personel damgası ekler"""
    reader = pypdf.PdfReader(uploaded_file)
    writer = pypdf.PdfWriter()
    
    page_info = []
    # 1. Tüm sayfaları gez, ürün adını ve sipariş kodunu çıkar
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        
        # Ürün adını yakala (Hepsiburada formatına göre)
        product_match = re.search(r'ÜRÜN KODU/ ADI\s+ADET\s+(.+?)(?=\n|$)', text, re.IGNORECASE | re.DOTALL)
        product_name = product_match.group(1).strip() if product_match else f"Sayfa {i+1}"
        
        # Sipariş kodunu da yedek olarak al
        order_match = re.search(r'SİPARİŞ KODU:\s*(.+?)(?=\n|$)', text)
        order_code = order_match.group(1).strip() if order_match else ""
        
        page_info.append({
            "index": i,
            "product": product_name,
            "order": order_code,
            "page": page
        })
    
    # 2. ÜRÜN BAZLI SIRALAMA (Aynı ürünler yan yana gelir)
    page_info.sort(key=lambda x: x["product"])
    
    # 3. Personel ataması yap ve damgayı orijinal sayfanın üzerine ekle
    for idx, info in enumerate(page_info):
        assigned_person = personel_list[idx % len(personel_list)] if personel_list else "Atanmadı"
        page = info["page"]
        
        # Sayfa boyutlarını al (10x15 cm veya başka bir boyut olabilir, dinamik olarak alıyoruz)
        page_box = page.mediabox
        width = float(page_box.width)
        height = float(page_box.height)
        
        # Personel damgası için geçici bir PDF oluştur
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))
        
        # Sağ alt köşeye veya ortaya sarı kutu çiz
        box_w, box_h = 130, 45
        c.setFillColor(yellow)
        c.rect(width - 140, 10, box_w, box_h, fill=1, stroke=1)
        c.setFillColor(black)
        c.setFont(FONT_NAME, 10)
        c.drawCentredString(width - 75, 40, "DEPO PERSONELİ:")
        c.setFont(FONT_NAME, 13)
        c.drawCentredString(width - 75, 20, assigned_person)
        c.save()
        
        packet.seek(0)
        stamp_pdf = pypdf.PdfReader(packet)
        
        # Orijinal sayfaya damgayı birleştir (Barkod ve diğer her şey %100 korunur)
        page.merge_page(stamp_pdf.pages[0])
        writer.add_page(page)
        
    out_buffer = io.BytesIO()
    writer.write(out_buffer)
    return out_buffer.getvalue(), len(page_info)

# --- ARAYÜZ ---
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("📦 Trendyol İşlemleri (10x10 cm)")
    st.write("Sentos Excel çıktısını yükleyin. Aynı siparişteki ürünler tek etikette birleştirilir, personel kodu eklenir.")
    trendyol_file = st.file_uploader("Trendyol Excel (.xlsx)", type=["xlsx"])
    if st.button("Trendyol Etiketlerini Oluştur", type="primary", use_container_width=True):
        if trendyol_file:
            with st.spinner("İşleniyor, lütfen bekleyin..."):
                try:
                    pdf_bytes, count = process_trendyol_excel(trendyol_file, personel_list)
                    st.success(f"✅ {count} adet sipariş için etiket oluşturuldu!")
                    st.download_button("📥 PDF İndir", data=pdf_bytes, file_name="Trendyol_Etiketler.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Hata oluştu: {str(e)}")
        else:
            st.warning("Lütfen bir Excel dosyası yükleyin.")

with col2:
    st.subheader("📦 Hepsiburada İşlemleri (Orijinal Barkod Korunur)")
    st.write("PDF yükleyin. Sistem barkodu BOZMAZ, sadece sayfaları ürün bazında sıralar ve personel kodunu damgalar.")
    hb_file = st.file_uploader("Hepsiburada PDF (.pdf)", type=["pdf"])
    if st.button("Hepsiburada Etiketlerini Oluştur", type="primary", use_container_width=True):
        if hb_file:
            with st.spinner("İşleniyor, lütfen bekleyin..."):
                try:
                    pdf_bytes, count = process_hepsiburada_pdf(hb_file, personel_list)
                    st.success(f"✅ {count} adet sayfa işlendi ve sıralandı!")
                    st.download_button("📥 PDF İndir", data=pdf_bytes, file_name="Hepsiburada_Etiketler.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Hata oluştu: {str(e)}")
        else:
            st.warning("Lütfen bir PDF dosyası yükleyin.")

st.markdown("---")
st.caption("💡 **Not:** Türkçe karakterlerin (ş, ı, ğ, ü, ö, ç) kusursuz görünmesi için proje klasörüne `arial.ttf` dosyasını eklemeyi unutmayın.")
