import streamlit as st
import pandas as pd
import io
import re
import os
import math
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
    """Trendyol Excel'den ürün adını çıkarır"""
    try:
        text = str(product_text)
        if '/' in text:
            parts = text.split('/', 1)
            name_part = parts[1].strip()
            # 'x1 Adet' gibi kısmı kes
            name_part = re.split(r'\s*x\s*\d+', name_part)[0].strip()
            # Uzunsa kısalt (sıralama için ilk 50 karakter yeter)
            return name_part[:50] if name_part else "Diger"
        return text[:50] if text else "Diger"
    except:
        return "Diger"

def extract_hepsiburada_product_name(text):
    """Hepsiburada PDF'ten ürün adını çıkarır"""
    try:
        # Format: "ÜRÜN KODU/ ADI ADET RG-BASCEK-CIRPICI/ ROYALGROSS Basçek..."
        match = re.search(r'ÜRÜN KODU/\s*ADI\s*ADET\s+(\S+)/\s*(.+?)(?:\d+)?$', text, re.DOTALL | re.IGNORECASE)
        if match:
            product_name = match.group(2).strip()
            # Sondaki rakamı (adet) temizle
            product_name = re.sub(r'\d+\s*$', '', product_name).strip()
            return product_name[:50] if product_name else "Diger"
        
        # Alternatif: sadece ürün kodunu al
        match2 = re.search(r'ÜRÜN KODU/\s*ADI\s*ADET\s+(\S+)', text)
        if match2:
            return match2.group(1)[:50]
        
        return "Diger"
    except:
        return "Diger"

def assign_personel_by_chunks(total_items, personel_list):
    """Adil personel dağılımı: toplam / personel sayısı kadar her birine"""
    if not personel_list:
        return ["At
