import streamlit as st
import pandas as pd
import io
import re
import os
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, black
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
st.markdown("*(Trendyol 10x10 cm | Hepsiburada Orijinal Barkod Korunur + Personel Damgası)*")
st.markdown("---")

st.sidebar.header("⚙️ Personel Ayarları")
personel_input = st.sidebar.text_area(
    "Personel İsimleri veya Kodları (Her satıra bir tane)",
    "Ahmet (Kod: 01)\nAyşe (Kod: 02)\nMehmet (Kod: 03)"
)
personel_list = [p.strip() for p in personel_input.split('\n') if p.strip()]

st.sidebar.info("️ **Yazıcı Boyutları:**\n- Trendyol: 10x10 cm\n- Hepsiburada: 10x15 cm")
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
    try:
        match = re.search(r'ÜRÜN KODU/\s*ADI\s*ADET\s+(.+)', text, re.IGNORECASE | re.DOTALL)
        if match:
            raw_text = match.group(1).strip()
            if '/' in raw_text:
                name_part = raw_text.split('/', 1)[1].strip()
            else:
                name_part = raw_text
            name_part = re.sub(r'\s+\d+\s*$', '', name_part).strip()
            name_part = " ".join(name_part.split())
            return name_part[:60] if name_part else "Diger_Urun"
        return "Diger_Urun"
    except:
        return "Diger_Urun"

def get_personel_royal_format(personel_list, index):
    """Personel listesinden 'Royal001' formatında kod üretir"""
    if not personel_list or index >= len(personel_list):
        return "Royal001"
    p_entry = personel_list[index]
    code_match = re.search(r'Kod:\s*(\d+)', p_entry)
    if code_match:
        code = code_match.group(1).zfill(3)
        return f"Royal{code}"
    else:
        code = str(index + 1).zfill(3)
        return f"Royal{code}"

def assign_personel_by_chunks(total_items, personel_list):
    """Ürün bazlı gruplu personel ataması"""
    if not personel_list:
        return ["Royal001"] * total_items
    
    n_personel = len(personel_list)
    pages_per_person = [total_items // n_personel] * n_personel
    for i in range(total_items % n_personel):
        pages_per_person[i] += 1
    
    assignments = []
    current_personel_idx = 0
    pages_assigned = 0
    
    for i in range(total_items):
        assignments.append(get_personel_royal_format(personel_list, current_personel_idx % n_personel))
        pages_assigned += 1
        if pages_assigned >= pages_per_person[current_personel_idx % n_personel]:
            current_personel_idx += 1
            pages_assigned = 0
    
    return assignments

def process_trendyol_excel(uploaded_file, personel_list):
    """Yeni şablona uygun Trendyol etiket üretimi"""
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    df.columns = df.columns.str.strip()
    
    grouped = df.groupby('Sipariş Kodu')
    
    order_products = []
    for order_code, group in grouped:
        first_product = str(group['Sipariş Verilen Ürün(ler)'].iloc[0]) if 'Sipariş Verilen Ürün(ler)' in group.columns else "Diger"
        product_name = extract_trendyol_product_name(first_product)
        order_products.append({"order_code": order_code, "product": product_name, "group": group})
    
    # Ürün bazlı sıralama
    order_products.sort(key=lambda x: x["product"].lower())
    
    # Personel ataması (ürün bazlı gruplu)
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
        
        # Ürün ve miktar bilgilerini hazırla
        products = []
        for _, row in group.iterrows():
            prod_name = str(row['Sipariş Verilen Ürün(ler)']).split(' x')[0] if 'Sipariş Verilen Ürün(ler)' in row else "Ürün"
            qty = int(row['Adet']) if 'Adet' in row and pd.notna(row['Adet']) else 1
            products.append((prod_name, qty))
        
        # İl/İlçe formatını hazırla
        city_district = city if city else ""
        
        # Kargo firması
        kargo_firmasi = "Trendyol Express Marketplace"
        
        # === YENİ ŞABLON TASARIMI ===
        
        # 1. ÜST SATIR: Müşteri adı | ROYALGROSS | Sipariş Kodu
        c.setFont(FONT_NAME, 11)
        c.drawString(0.5*cm, 9.2*cm, customer)
        
        c.setFont(FONT_NAME, 12)
        # ROYALGROSS'u ortaya hizala
        royal_width = c.stringWidth("ROYALGROSS", FONT_NAME, 12)
        c.drawString((10*cm - royal_width)/2, 9.2*cm, "ROYALGROSS")
        
        c.setFont(FONT_NAME, 11)
        c.drawRightString(9.5*cm, 9.2*cm, str(order_code))
        
        # 2. İKİNCİ SATIR: İlçe/İl | Kargo firması
        c.setFont(FONT_NAME, 9)
        c.drawString(0.5*cm, 8.7*cm, city_district)
        c.drawRightString(9.5*cm, 8.7*cm, kargo_firmasi)
        
        # 3. BARKOD (ortada)
        barcode_img = generate_barcode_image(barcode_val)
        c.drawImage(barcode_img, 1*cm, 6.5*cm, width=8*cm, height=2*cm)
        
        # Barkod numarası (barkodun altında, ortada)
        c.setFont(FONT_NAME, 10)
        barcode_text_width = c.stringWidth(str(barcode_val), FONT_NAME, 10)
        c.drawString((10*cm - barcode_text_width)/2, 6.2*cm, str(barcode_val))
        
        # 4. ÜRÜN ve MİKTAR başlıkları (altı çizili)
        c.setFont(FONT_NAME, 9)
        c.drawString(0.5*cm, 5.5*cm, "ÜRÜN")
        c.drawRightString(9.5*cm, 5.5*cm, "MİKTAR")
        
        # Altı çizgi
        c.line(0.5*cm, 5.4*cm, 9.5*cm, 5.4*cm)
        
        # 5. Ürün bilgileri
        c.setFont(FONT_NAME, 9)
        y_pos = 5.0*cm
        for prod_name, qty in products:
            # Ürün adı (sol)
            if len(prod_name) > 55:
                c.drawString(0.5*cm, y_pos, prod_name[:52] + "...")
                y_pos -= 0.4*cm
                c.drawString(0.5*cm, y_pos, prod_name[52:])
            else:
                c.drawString(0.5*cm, y_pos, prod_name)
            
            # Miktar (sağ)
            c.drawRightString(9.5*cm, y_pos, f"{qty} adet")
            y_pos -= 0.5*cm
        
        # 6. PERSONEL KUTUSU (sağ alt - "Royal002" formatında)
        box_w = 2.2*cm
        box_h = 0.9*cm
        box_x = 10*cm - box_w - 0.5*cm
        box_y = 0.5*cm
        
        # Krem/sarı arka plan
        c.setFillColor(HexColor('#FFF8DC'))  # Krem rengi
        c.setStrokeColor(HexColor('#CC3333'))  # Kırmızımsı kenarlık
        c.setLineWidth(1.5)
        c.roundRect(box_x, box_y, box_w, box_h, 3, fill=1, stroke=1)
        
        # Personel kodu (kırmızımsı yazı)
        c.setFillColor(HexColor('#CC3333'))
        c.setFont(FONT_NAME, 11)
        code_width = c.stringWidth(assigned_person, FONT_NAME, 11)
        c.drawString(box_x + (box_w - code_width)/2, box_y + 0.25*cm, assigned_person)
        
        c.showPage()
    
    c.save()
    return buffer.getvalue(), len(order_products)

def process_hepsiburada_pdf(uploaded_file, personel_list):
    """Hepsiburada: Ürün bazlı sıralama + personel damgası (mevcut çalışan kod)"""
    reader = pypdf.PdfReader(uploaded_file)
    writer = pypdf.PdfWriter()
    
    page_info = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        order_match = re.search(r'SİPARİŞ KODU:\s*([0-9\-]+)', text)
        order_code = order_match.group(1).strip() if order_match else f"Sayfa_{i+1}"
        product_name = extract_hepsiburada_product_name(text)
        page_info.append({
            "index": i,
            "product": product_name,
            "order": order_code,
            "page": page
        })
    
    # Ürün bazlı sıralama
    page_info.sort(key=lambda x: (x["product"].lower(), x["order"]))
    
    # Personel ataması (ürün bazlı gruplu)
    assignments = assign_personel_by_chunks(len(page_info), personel_list)
    
    for idx, info in enumerate(page_info):
        assigned_person = assignments[idx]
        original_page = info["page"]
        
        new_page = pypdf.PageObject.create_blank_page(
            width=original_page.mediabox.width,
            height=original_page.mediabox.height
        )
        new_page.merge_page(original_page)
        
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(float(original_page.mediabox.width), float(original_page.mediabox.height)))
        
        # Personel damgası - "Royal001" formatında, sağ alt
        box_w = 130
        box_h = 45
        width = float(original_page.mediabox.width)
        height = float(original_page.mediabox.height)
        
        # Krem arka plan, kırmızı kenarlık
        c.setFillColor(HexColor('#FFF8DC'))
        c.setStrokeColor(HexColor('#CC3333'))
        c.setLineWidth(1.5)
        c.roundRect(width - 145, 10, box_w, box_h, 3, fill=1, stroke=1)
        
        # Personel kodu
        c.setFillColor(HexColor('#CC3333'))
        c.setFont(FONT_NAME, 14)
        code_width = c.stringWidth(assigned_person, FONT_NAME, 14)
        c.drawString(width - 145 + (box_w - code_width)/2, 25, assigned_person)
        
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
st.write("Sentos Excel çıktısını yükleyin. Siparişler **ÜRÜN BAZINDA** sıralanır, personel **Royal001** formatında atanır.")
trendyol_file = st.file_uploader("Trendyol Excel dosyasını seçin (.xlsx)", type=["xlsx"], key="trendyol_uploader")
if trendyol_file is not None:
    st.success(f"✅ Dosya yüklendi: {trendyol_file.name}")
    if st.button("🚀 Trendyol Etiketlerini Oluştur", type="primary", use_container_width=True):
        with st.spinner("İşleniyor, lütfen bekleyin..."):
            try:
                pdf_bytes, count = process_trendyol_excel(trendyol_file, personel_list)
                st.success(f"✅ {count} adet sipariş için 10x10 cm etiket oluşturuldu!")
                st.download_button(" Trendyol PDF'ini İndir", data=pdf_bytes, file_name="Trendyol_Etiketler.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ Hata oluştu: {str(e)}")
else:
    st.info("👆 Yukarıdan bir Excel dosyası seçin.")

st.markdown("---")

st.subheader("📦 HEPSİBURADA İŞLEMLERİ (10x15 cm)")
st.write("PDF yükleyin. Sistem barkodu BOZMAZ, sayfaları **ÜRÜN BAZINDA** sıralar ve **Royal001** formatında personel kodu ekler.")
hb_file = st.file_uploader("Hepsiburada PDF dosyasını seçin (.pdf)", type=["pdf"], key="hepsiburada_uploader")
if hb_file is not None:
    st.success(f"✅ Dosya yüklendi: {hb_file.name}")
    if st.button("🚀 Hepsiburada Etiketlerini Oluştur", type="primary", use_container_width=True):
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
