import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.title("Mühendislik Proje Raporu Otomasyonu")
st.write("Lütfen kurumsal kapak bilgilerini giriniz:")

# Girdi Alanları
sirket_adi = st.text_input("Şirket / Kuruluş İsmi", "FUGAMEKANİK MÜHENDİSLİK A.Ş.")
is_adi = st.text_input("İşin Adı / Proje Başlığı", "Merkezi Isıtma ve Havalandırma Tesisatı Projesi")
rapor_turu = st.text_input("Rapor Türü", "MEKANİK TESİSAT HESAP RAPORU")
hazirlayan = st.text_input("Hazırlayan Mühendis", "Ahmet Yılmaz (Makine Mühendisi)")
tarih = st.text_input("Rapor Tarihi", "Eylül 2026")

# Rapor Oluştur Butonu
if st.button("Profesyonel Kapaklı Word Raporu Oluştur"):
    if is_adi and sirket_adi:
        doc = Document()
        
        # Sayfa boşluklarını ayarlayalım (Kapak için ferah olsun)
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1.5)
            section.bottom_margin = Inches(1.5)
            section.left_margin = Inches(1.2)
            section.right_margin = Inches(1.2)

        # 1. Şirket / Kuruluş İsmi (Büyük ve Ortalanmış)
        p_sirket = doc.add_paragraph()
        p_sirket.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sirket = p_sirket.add_run(sirket_adi.upper())
        run_sirket.font.size = Pt(18)
        run_sirket.font.bold = True
        run_sirket.font.name = 'Arial'
        
        # Araya biraz boşluk bırakalım
        doc.add_paragraph()
        doc.add_paragraph()

        # 2. İşin Adı (Proje Başlığı)
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

        # 3. Rapor Türü (Mekanik Tesisat Hesap Raporu)
        p_tur = doc.add_paragraph()
        p_tur.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_tur = p_tur.add_run(rapor_turu.upper())
        run_tur.font.size = Pt(14)
        run_tur.font.bold = True
        run_tur.font.name = 'Arial'

        # Sayfanın alt kısmına doğru boşluk bırakmak için birkaç paragraf ekleyelim
        for _ in range(4):
            doc.add_paragraph()

        # 4. En Altta Hazırlayan ve Tarih (Sağa veya Ortaya Hizalı)
        p_alt = doc.add_paragraph()
        p_alt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run_hazirlayan = p_alt.add_run(f"Hazırlayan:\n{hazirlayan}\n\nTarih:\n{tarih}")
        run_hazirlayan.font.size = Pt(11)
        run_hazirlayan.font.name = 'Arial'

        # Hafızada dosya tutma ve indirme hazırlığı
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("Kapak sayfası tasarımı başarıyla oluşturuldu!")
        
        st.download_button(
            label="📥 Düzenli Kapaklı Word Dosyasını İndir (.docx)",
            data=buffer,
            file_name=f"{is_adi.replace(' ', '_')}_Kapak_Raporu.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Lütfen Şirket İsmi ve Proje Adı alanlarını doldurun.")
