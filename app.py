import streamlit as st

st.title("Mühendislik Proje Raporu Otomasyonu")
st.write("Lütfen proje kapak bilgilerini giriniz:")

# Kapak Bilgileri Alanları
proje_adi = st.text_input("Proje Adı")
isveren = st.text_input("İşveren / Kurum")
tarih = st.text_input("Rapor Tarihi")
hazirlayan = st.text_input("Hazırlayan Mühendis")

# Rapor Oluştur Butonu
if st.button("Word Raporu Oluştur"):
    if proje_adi:
        st.success(f"Harika! '{proje_adi}' projesi için veriler alındı.")
    else:
        st.warning("Lütfen en azından Proje Adı alanını doldurun.")
