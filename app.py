import streamlit as st
from docx import Document
import io

st.title("Mühendislik Proje Raporu Otomasyonu")
st.write("Lütfen proje kapak bilgilerini giriniz:")

# Kapak Bilgileri Alanları
proje_adi = st.text_input("Proje Adı", "Örnek Mekanik Proje")
isveren = st.text_input("İşveren / Kurum", "Örnek Kurum A.Ş.")
tarih = st.text_input("Rapor Tarihi", "2026-09-20")
hazirlayan = st.text_input("Hazırlayan Mühendis", "Mühendis")

# Rapor Oluştur Butonu
if st.button("Word Raporu Oluştur"):
    if proje_adi:
        # Word belgesini oluşturalım
        doc = Document()
        doc.add_heading(proje_adi, 0)
        doc.add_paragraph(f"İşveren / Kurum: {isveren}")
        doc.add_paragraph(f"Rapor Tarihi: {tarih}")
        doc.add_paragraph(f"Hazırlayan Mühendis: {hazirlayan}")
        
        # Hafızada dosya tutma (BytesIO)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("Word raporu başarıyla oluşturuldu!")
        
        # İndirme Butonu
        st.download_button(
            label="📥 Word Dosyasını İndir (.docx)",
            data=buffer,
            file_name=f"{proje_adi.replace(' ', '_')}_Rapor.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Lütfen en azından Proje Adı alanını doldurun.")
