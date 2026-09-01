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

# Türkçe karakter desteği
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
st.markdown("*(Trendyol 10x10 cm | Hepsiburada Orijinal Barkod Korunur + Sayfa Sıralama + Personel Damgası)*")
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

def extract_trendyol_product_name(product_text):
    try:
        text = str(product_text)
        if '/' in text:
            parts = text.split('/', 1)
            name_part = parts[1].strip()
            name_part = re.split(r'\s*x\s*\d+', name_part)[0].strip()
            return name_part[:50] if name_part else "Diger"
        return text[:50] if text else "Diger"
    except:
        return "Diger"

def extract_hepsiburada_product_name(text):
    """Hepsiburada PDF'inden ürün adını %100 doğru çeker"""
    try:
        # "ÜRÜN KODU/ ADI ADET [KOD]/ [ÜRÜN ADI] [ADET]" formatını yakalar
        match = re.search(r'ÜRÜN KODU/\s*ADI\s*ADET\s+(.+)', text)
        if match:
            product_line = match.group(1).strip()
            # '/' işaretinden sonrasını al (Ürün Adı kısmı)
            if '/' in product_line:
                name_part = product_line.split('/', 1)[1].strip()
            else:
                name_part = product_line
            
            # Sondaki adet rakamını (örn: " 1", " 2") temizle
            name_part = re.sub(r'\s+\d+\s*$', '', name_part).strip()
            return name_part[:60] if name_part else "Diger_Urun"
        
        return "Diger_Urun"
    except:
        return "Diger_Urun"

def assign_personel_by_chunks(total_pages, personel_list):
    """Toplam sayfa / personel sayısı mantığıyla adil blok dağıtım"""
    if not personel_list:
        return ["DEPO[00]"] * total_pages
    
    n_personel = len(personel_list)
    base_count = total_pages // n_personel
    remainder = total_pages % n_personel
    
    assignments = []
    for i in range(total_pages):
        if i < remainder * (base_count + 1):
            current_idx = i // (base_count + 1)
        else:
            adjusted_i = i - remainder * (base_count + 1)
            current_idx = remainder + (adjusted_i // base_count)
        
        p_entry = personel_list[current_idx % n_personel]
        code_match = re.search(r'Kod:\s*(\d+)', p_entry)
        if code_match:
            code = code_match.group(1).zfill(2)
            assignments.append(f"DEPO[{code}]")
        else:
            code = str((current_idx % n_personel) + 1).zfill(2)
            assignments.append(f"DEPO[{code}]")
            
    return assignments

def process_trendyol_excel(uploaded_file, personel_list):
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    df.columns = df.columns.str.strip()
    grouped = df.groupby('Sipariş Kodu')
    
    order_products = []
    for order_code, group in grouped:
        first_product = str(group['Sipariş Verilen Ürün(ler)'].iloc[0]) if 'Sipariş Verilen Ürün(ler)' in group.columns else "Diger"
        product_name = extract_trendyol_product_name(first_product)
        order_products.append({"order_code": order_code, "product": product_name, "group": group})
    
    order_products.sort(key=lambda x: x["product"].lower())
    assignments = assign_personel_by_chunks(len(order_products), personel_list)
    
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(10*cm, 10*cm))
    
    for idx, order_info in enumerate(order_products):
        order_code = order_info["order_code"]
        group = order_info["group"]
        assigned_person = assignments[idx]
        
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
        c.drawString(1*cm, 6*cm, full_address[:52] + "..." if len(full_address) > 55 else full_address)
        
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
    return buffer.getvalue(), len(order_products)

def process_hepsiburada_pdf(uploaded_file, personel_list):
    reader = pypdf.PdfReader(uploaded_file)
    writer = pypdf.PdfWriter()
    
    page_info = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        order_match = re.search(r'SİPARİŞ KODU:\s*([0-9\-]+)', text)
        order_code = order_match.group(1).strip() if order_match else f"Sayfa_{i+1}"
        
        # KRİTİK DÜZELTME: Ürün adını artık doğru yakalıyor
        product_name = extract_hepsiburada_product_name(text)
        
        page_info.append({
            "index": i,
            "product": product_name,
            "order": order_code,
            "page": page
        })
    
    # 1. ADIM: Sayfaları ÜRÜN ADINA GÖRE alfabetik sırala (Aynı ürünler yan yana gelir)
    page_info.sort(key=lambda x: (x["product"].lower(), x["order"]))
    
    # 2. ADIM: Sıralanmış sayfaları personele blok halinde dağıt
    assignments = assign_personel_by_chunks(len(page_info), personel_list)
    
    # 3. ADIM: Yeni PDF'i oluştur ve damgaları ekle
    for idx, info in enumerate(page_info):
        assigned_person = assignments[idx]
        page = info["page"]
        page_box = page.mediabox
        width = float(page_box.width)
        height = float(page_box.height)
        
        # Personel damgası overlay (Sağ alt köşe)
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
        
        # Orijinal sayfaya damgayı birleştir (Barkod %100 korunur)
        page.merge_page(stamp_pdf.pages[0])
        writer.add_page(page)
    
    out_buffer = io.BytesIO()
    writer.write(out_buffer)
    return out_buffer.getvalue(), len(page_info)

# ============================================
# ARAYÜZ
# ============================================
st.subheader("📦 TRENDYOL İŞLEMLERİ (10x10 cm)")
st.write("Sentos Excel çıktısını yükleyin. Siparişler **ÜRÜN BAZINDA** sıralanır, personel **adil dağıtılır**.")
trendyol_file = st.file_uploader("Trendyol Excel dosyasını seçin (.xlsx)", type=["xlsx"], key="trendyol_uploader")
if trendyol_file is not None:
    st.success(f"✅ Dosya yüklendi: {trendyol_file.name}")
    if st.button("🚀 Trendyol Etiketlerini Oluştur", type="primary", use_container_width=True):
        with st.spinner("İşleniyor, lütfen bekleyin..."):
            try:
                pdf_bytes, count = process_trendyol_excel(trendyol_file, personel_list)
                st.success(f"✅ {count} adet sipariş için 10x10 cm etiket oluşturuldu!")
                st.download_button("📥 Trendyol PDF'ini İndir", data=pdf_bytes, file_name="Trendyol_Etiketler.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ Hata oluştu: {str(e)}")
else:
    st.info("👆 Yukarıdan bir Excel dosyası seçin.")

st.markdown("---")

st.subheader("📦 HEPSİBURADA İŞLEMLERİ (10x15 cm)")
st.write("PDF yükleyin. Sistem barkodu BOZMAZ, sayfaları **AYNI ÜRÜN BAZINDA YAN YANA** dizer ve personeli blok halinde dağıtır.")
hb_file = st.file_uploader("Hepsiburada PDF dosyasını seçin (.pdf)", type=["pdf"], key="hepsiburada_uploader")
if hb_file is not None:
    st.success(f"✅ Dosya yüklendi: {hb_file.name}")
    if st.button("🚀 Hepsiburada Etiketlerini Oluştur", type="primary", use_container_width=True):
        with st.spinner("İşleniyor, lütfen bekleyin..."):
            try:
                pdf_bytes, count = process_hepsiburada_pdf(hb_file, personel_list)
                st.success(f"✅ {count} adet sayfa ÜRÜN BAZINDA sıralandı ve personel atandı!")
                st.download_button("📥 Hepsiburada PDF'ini İndir", data=pdf_bytes, file_name="Hepsiburada_Etiketler.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ Hata oluştu: {str(e)}")
else:
    st.info("👆 Yukarıdan bir PDF dosyası seçin.")

st.markdown("---")
st.caption("🏭 ROYALGROSS EV GEREÇLERİ DIŞ TİCARET LİMİTED ŞİRKETİ - Otomatik Etiket Sistemi")
