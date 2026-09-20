import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.title("Mühendislik Proje Raporu Otomasyonu")
st.write("Lütfen kurumsal kapak ve genel bilgiler kısımlarını doldurun:")

# --- 1. SEKME / BÖLÜM: KAPAK BİLGİLERİ ---
st.header("1. Kapak Bilgileri")
sirket_adi = st.text_input("Şirket / Kuruluş İsmi", "FUGAMEKANİK MÜHENDİSLİK A.Ş.")
is_adi = st.text_input("İşin Adı / Proje Başlığı", "Merkezi Isıtma ve Havalandırma Tesisatı Projesi")
rapor_turu = st.text_input("Rapor Türü", "MEKANİK TESİSAT HESAP RAPORU")
hazirlayan = st.text_input("Hazırlayan Mühendis", "Mehmet KÜÇÜK (Makine Mühendisi)")
tarih = st.text_input("Rapor Tarihi", "Eylül 2026")

# --- 2. SEKME / BÖLÜM: GENEL BİLGİLER ---
st.header("2. Genel Bilgiler ve Tasarım Kriterleri")
proje_yeri = st.text_input("Projenin Yeri / İl", "Ankara")
dis_hava_sicaklik = st.text_input("Dış Hava Tasarım Sıcaklığı (°C)", "-12 °C")
ic_hava_sicaklik = st.text_input("İç Ortam Tasarım Sıcaklığı (°C)", "20 °C")
yonetmelik_standart = st.text_input("Esas Alınan Standart / Yönetmelik", "TS 825, ASHRAE, Binalarda Yangın Korunması Yönetmeliği")
proje_aciklamasi = st.text_area("Ek Açıklamalar", "Bu rapor, ilgili bina mekanik tesisat sistemlerinin, yürürlükteki standartlar ve yönetmeliklere uygun olarak tasarlanması amacıyla hazırlanmıştır.")

# Rapor Oluştur Butonu
if st.button("Genel Bilgiler Dahil Word Raporu Oluştur"):
    if is_adi and sirket_adi:
        doc = Document()
        
        # Sayfa boşlukları
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1.5)
            section.bottom_margin = Inches(1.5)
            section.left_margin = Inches(1.2)
            section.right_margin = Inches(1.2)

        # --- KAPAK SAYFASI ---
        p_sirket = doc.add_paragraph()
        p_sirket.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sirket = p_sirket.add_run(sirket_adi.upper())
        run_sirket.font.size = Pt(18)
        run_sirket.font.bold = True
        run_sirket.font.name = 'Arial'
        
        doc.add_paragraph()
        doc.add_paragraph()

        p_is = doc.add_paragraph()
        p_is.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_is_baslik = p_is.add_run("PROJE ADI:\n")
        run_is_baslik.font.size = Pt(11)
        run_is_baslik.font.name = 'Arial'
        
        run_is = p_is.add_run(is_adi)
        run_is.font.size = Pt(16)
        run_is.font.bold = True
        run_is.font.name = 'Arial'

        doc.add_paragraph()

        p_tur = doc.add_paragraph()
        p_tur.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_tur = p_tur.add_run(rapor_turu.upper())
        run_tur.font.size = Pt(14)
        run_tur.font.bold = True
        run_tur.font.name = 'Arial'

        for _ in range(4):
            doc.add_paragraph()

        p_alt = doc.add_paragraph()
        p_alt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_hazirlayan = p_alt.add_run(f"Hazırlayan:\n{hazirlayan}\n\nTarih:\n{tarih}")
        run_hazirlayan.font.size = Pt(11)
        run_hazirlayan.font.name = 'Arial'

        # Yeni sayfaya geçiş (Genel Bilgiler için)
        doc.add_page_break()

        # --- BÖLÜM 1: GENEL BİLGİLER ---
        doc.add_heading("1. GENEL BİLGİLER VE TASARIM KRİTERLERİ", level=1)
        
        # İstediğiniz Sabit Giriş Metni (Proje Adı dinamik ekleniyor)
        giris_metni = f"Bu raporda '{is_adi}' için tasarlanan mekanik tesisatlar açıklanmış ve tüm uygulama ve detay projelerine esas teşkil eden tasarım kriterleri ve mekanik tesisat sistem çözümleri tespit edilmiştir."
        doc.add_paragraph(giris_metni)
        
        if proje_aciklamasi:
            doc.add_paragraph(proje_aciklamasi)
        
        # Proje Parametreleri Başlığı ve Listesi
        doc.add_heading("1.1. Proje Parametreleri", level=2)
        
        p_bilgi = doc.add_paragraph()
        p_bilgi.add_run(f"• Proje Yeri / İl: ").bold = True
        p_bilgi.add_run(f"{proje_yeri}\n")
        p_bilgi.add_run(f"• Dış Hava Tasarım Sıcaklığı: ").bold = True
        p_bilgi.add_run(f"{dis_hava_sicaklik}\n")
        p_bilgi.add_run(f"• İç Ortam Tasarım Sıcaklığı: ").bold = True
        p_bilgi.add_run(f"{ic_hava_sicaklik}\n")
        p_bilgi.add_run(f"• Esas Alınan Standartlar: ").bold = True
        p_bilgi.add_run(f"{yonetmelik_standart}\n")

        # Hafızada dosya oluşturma
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("Sabit metin ve genel bilgiler rapora eklendi!")
        
        st.download_button(
            label="📥 Güncel Word Dosyasını İndir (.docx)",
            data=buffer,
            file_name=f"{is_adi.replace(' ', '_')}_Genel_Bilgiler.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Lütfen Şirket İsmi ve Proje Adı alanlarını doldurun.")
