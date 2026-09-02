import streamlit as st
import pandas as pd
import io
import re
import os
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.colors import yellow, black, HexColor
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

st.set_page_config(page_title="ROYALGROSS Etiket Otomasyonu", layout="wide", page_icon="📦")
st.title("📦 ROYALGROSS Günlük Etiket ve Barkod Otomasyonu")
st.markdown("*(Trendyol 10x10 cm | Hepsiburada Orijinal Barkod Korunur + Ürün Bazlı Sıralama)*")
st.markdown("---")

st.sidebar.header("⚙️ Personel Ayarları")
personel_input = st.sidebar.text_area(
    "Personel İsimleri veya Kodları (Her satıra bir tane)",
    "Ahmet (Kod: 01)\nAyşe (Kod: 02)\nMehmet (Kod: 03)"
)
personel_list = [p.strip() for p in personel_input.split('\n') if p.strip()]

st.sidebar.info("🖨️ **Yazıcı Boyutları:**\n- Trendyol: 10x10 cm\n- Hepsiburada: 10x15 cm")
st.sidebar.markdown("---")

def generate_barcode_image(code):
    clean_code = re.sub(r'[^A-Za-z0-9]', '', str(code))
    if not clean_code:
        clean_code = "000000000000"
    code128 = Code128(clean_code, writer=ImageWriter())
    buffer = io.BytesIO()
    code128.write(buffer, options={"module_width": 0.2, "module_height": 10.0, "font_size": 8})
    return buffer

def get_column(df, possible_names, default=None):
    """Kolon ismini esnek olarak bul"""
    for name in possible_names:
        if name in df.columns:
            return name
    return default

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
    """Hepsiburada PDF'ten ürün adını çıkar - GÜÇLENDİRİLMİŞ"""
    try:
        # Format: "ÜRÜN KODU/ ADI ADET [KOD]/ [ÜRÜN ADI] [ADET]"
        # Birden fazla satır olabilir, hepsini birleştir
        text = re.sub(r'\s+', ' ', text)
        
        match = re.search(r'ÜRÜN\s+KODU/\s*ADI\s+ADET\s+(.+?)(?:DEPO\s+PERSONELİ|$)', text, re.IGNORECASE)
        if match:
            raw_text = match.group(1).strip()
            # '/' işaretinden sonrasını al
            if '/' in raw_text:
                name_part = raw_text.split('/', 1)[1].strip()
            else:
                name_part = raw_text
            
            # Sondaki adet rakamını temizle
            name_part = re.sub(r'\s+\d+\s*$', '', name_part).strip()
            name_part = " ".join(name_part.split())
            return name_part[:60] if name_part else "Diger_Urun"
        
        return "Diger_Urun"
    except:
        return "Diger_Urun"

def assign_personel_by_chunks(total_items, personel_list):
    if not personel_list:
        return ["DEPO[01]"] * total_items
    
    n_personel = len(personel_list)
    assignments = []
    pages_per_person = [total_items // n_personel] * n_personel
    for i in range(total_items % n_personel):
        pages_per_person[i] += 1
    
    current_personel_idx = 0
    pages_assigned = 0
    
    for i in range(total_items):
        p_entry = personel_list[current_personel_idx % n_personel]
        code_match = re.search(r'Kod:\s*(\d+)', p_entry)
        if code_match:
            code = code_match.group(1).zfill(2)
            assignments.append(f"DEPO[{code}]")
        else:
            code = str((current_personel_idx % n_personel) + 1).zfill(2)
            assignments.append(f"DEPO[{code}]")
        
        pages_assigned += 1
        if pages_assigned >= pages_per_person[current_personel_idx % n_personel]:
            current_personel_idx += 1
            pages_assigned = 0
    
    return assignments

def process_trendyol_excel(uploaded_file, personel_list):
    """Trendyol Excel işleme - ESNEK KOLON İSİMLERİ"""
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    df.columns = df.columns.str.strip()
    
    # Esnek kolon eşleştirme
    order_col = get_column(df, ['Sipariş Numarası', 'Sipariş Kodu', 'Sipariş ID'])
    barcode_col = get_column(df, ['Barkod', 'Kampanya Kodu'])
    customer_col = get_column(df, ['Alıcı', 'Sipariş Veren Cari', 'Alıcı Ad/Soyad'])
    address_col = get_column(df, ['Teslimat Adresi', 'Alıcı Adres'])
    city_col = get_column(df, ['İl'])
    district_col = get_column(df, ['İlçe'])
    product_col = get_column(df, ['Ürün Adı', 'Sipariş Verilen Ürün(ler)', 'Ürün Platform İsmi'])
    qty_col = get_column(df, ['Adet'])
    
    if not order_col:
        raise Exception(f"Excel'de 'Sipariş Numarası' veya 'Sipariş Kodu' kolonu bulunamadı. Mevcut kolonlar: {list(df.columns)}")
    
    # Sipariş bazında grupla
    grouped = df.groupby(order_col)
    
    order_products = []
    for order_code, group in grouped:
        first_product = str(group[product_col].iloc[0]) if product_col and product_col in group.columns else "Diger"
        product_name = extract_trendyol_product_name(first_product)
        order_products.append({
            "order_code": order_code,
            "product": product_name,
            "group": group
        })
    
    # ÜRÜN BAZINDA SIRALAMA
    order_products.sort(key=lambda x: x["product"].lower())
    
    # Personel ataması
    assignments = assign_personel_by_chunks(len(order_products), personel_list)
    
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(10*cm, 10*cm))
    
    for idx, order_info in enumerate(order_products):
        order_code = order_info["order_code"]
        group = order_info["group"]
        assigned_person = assignments[idx]
        
        barcode_val = str(group[barcode_col].iloc[0]) if barcode_col and barcode_col in group.columns and pd.notna(group[barcode_col].iloc[0]) else str(order_code)
        customer = str(group[customer_col].iloc[0]) if customer_col and customer_col in group.columns else "Müşteri"
        address = str(group[address_col].iloc[0]) if address_col and address_col in group.columns else ""
        city = str(group[city_col].iloc[0]) if city_col and city_col in group.columns else ""
        district = str(group[district_col].iloc[0]) if district_col and district_col in group.columns else ""
        
        products = []
        for _, row in group.iterrows():
            prod_name = str(row[product_col]).split(' x')[0] if product_col and product_col in row else "Ürün"
            qty = int(row[qty_col]) if qty_col and qty_col in row and pd.notna(row[qty_col]) else 1
            products.append(f"{qty}x {prod_name}")
        
        # Barkod
        barcode_img = generate_barcode_image(barcode_val)
        c.drawImage(barcode_img, 1*cm, 7.5*cm, width=8*cm, height=2*cm)
        
        # Bilgiler
        c.setFont(FONT_NAME, 10)
        c.drawString(1*cm, 7*cm, f"Sipariş: {order_code}")
        c.setFont(FONT_NAME, 9)
        c.drawString(1*cm, 6.5*cm, f"Müşteri: {customer}")
        full_address = f"{address} {district}/{city}" if district else f"{address} {city}"
        c.drawString(1*cm, 6*cm, full_address[:55] + "..." if len(full_address) > 55 else full_address)
        
        # Ürün listesi
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
        
        # Personel damgası
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
    """Hepsiburada PDF işleme - ÜRÜN BAZINDA SIRALAMA + FİZİKSEL YER DEĞİŞTİRME"""
    reader = pypdf.PdfReader(uploaded_file)
    
    page_info = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        
        # Sipariş kodunu al
        order_match = re.search(r'SİPARİŞ\s+KODU:\s*([0-9\-]+)', text)
        order_code = order_match.group(1).strip() if order_match else f"Sayfa_{i+1}"
        
        # Ürün adını al - GÜÇLENDİRİLMİŞ
        product_name = extract_hepsiburada_product_name(text)
        
        page_info.append({
            "index": i,
            "product": product_name,
            "order": order_code,
            "page": page
        })
    
    # DEBUG: Çıkarılan ürün adlarını göster
    st.write("**🔍 DEBUG - Çıkarılan Ürün Adları (ilk 10):**")
    for info in page_info[:10]:
        st.write(f"Sayfa {info['index']+1}: {info['product']}")
    
    # ÜRÜN BAZINDA SIRALAMA
    page_info.sort(key=lambda x: (x["product"].lower(), x["order"]))
    
    # Personel ataması
    assignments = assign_personel_by_chunks(len(page_info), personel_list)
    
    # YENİ PDF OLUŞTUR - Sayfaları sıralı şekilde ekle
    writer = pypdf.PdfWriter()
    
    for idx, info in enumerate(page_info):
        assigned_person = assignments[idx]
        original_page = info["page"]
        
        # Yeni boş sayfa oluştur (sıralamayı garanti eder)
        new_page = pypdf.PageObject.create_blank_page(
            width=original_page.mediabox.width,
            height=original_page.mediabox.height
        )
        new_page.merge_page(original_page)
        
        # Personel damgası overlay
        width = float(original_page.mediabox.width)
        height = float(original_page.mediabox.height)
        
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
        new_page.merge_page(stamp_pdf.pages[0])
        writer.add_page(new_page)
    
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
                st.info("💡 Excel kolon isimlerini kontrol edin: Sipariş Numarası, Barkod, Alıcı, Teslimat Adresi, İl, İlçe, Ürün Adı, Adet")
else:
    st.info(" Yukarıdan bir Excel dosyası seçin.")

st.markdown("---")

st.subheader(" HEPSİBURADA İŞLEMLERİ (10x15 cm)")
st.write("PDF yükleyin. Sistem barkodu BOZMAZ, sayfaları **ÜRÜN BAZINDA** sıralar ve personel kodunu damgalar.")
hb_file = st.file_uploader("Hepsiburada PDF dosyasını seçin (.pdf)", type=["pdf"], key="hepsiburada_uploader")
if hb_file is not None:
    st.success(f"✅ Dosya yüklendi: {hb_file.name}")
    if st.button(" Hepsiburada Etiketlerini Oluştur", type="primary", use_container_width=True):
        with st.spinner("İşleniyor, lütfen bekleyin..."):
            try:
                pdf_bytes, count = process_hepsiburada_pdf(hb_file, personel_list)
                st.success(f"✅ {count} adet sayfa işlendi ve ürün bazında sıralandı!")
                st.download_button("📥 Hepsiburada PDF'ini İndir", data=pdf_bytes, file_name="Hepsiburada_Etiketler.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ Hata oluştu: {str(e)}")
else:
    st.info("👆 Yukarıdan bir PDF dosyası seçin.")

st.markdown("---")
st.caption("🏭 ROYALGROSS EV GEREÇLERİ DIŞ TİCARET LİMİTED ŞİRKETİ - Otomatik Etiket Sistemi")
