import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from datetime import datetime

# Türkçe ay isimleri için sözlük
aylar = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
}

bugun = datetime.now()
bugun_ay_yil = f"{aylar[bugun.month]} {bugun.year}"

st.title("Mühendislik Proje Raporu Otomasyonu")
st.write("Lütfen kurumsal kapak ve genel bilgiler kısımlarını doldurun:")

# --- 1. SEKME / BÖLÜM: KAPAK BİLGİLERİ ---
st.header("1. Kapak Bilgileri")
sirket_adi = st.text_input("Şirket / Kuruluş İsmi", "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ")
is_adi = st.text_input("İşin Adı / Proje Başlığı", "")
rapor_turu = st.text_input("Rapor Türü", "MEKANİK TESİSAT UYGULAMA PROJESİ HESAP RAPORU")
hazirlayan = st.text_input("Hazırlayan Mühendis", "Mehmet Küçük")
mmo_no = st.text_input("MMO Oda No", "109913")
tarih = st.text_input("Rapor Tarihi", bugun_ay_yil)

# --- 2. SEKME / BÖLÜM: GENEL BİLGİLER ---
st.header("2. Genel Bilgiler")
st.write("Genel bilgiler bölümü standart kurumsal metinlerle otomatik olarak oluşturulacaktır.")

# İleride 3. sekmeden gelecek şehir değişkeni
sehir = "" 

# Rapor Oluştur Butonu
if st.button("Genel Bilgiler Dahil Word Raporu Oluştur"):
    if not is_adi:
        st.warning("⚠️ Dikkat: İşin Adı / Proje Başlığı girilmedi. Rapor oluşturuluyor ancak kapak başlığı boş bırakılacak.")
    
    aktif_sirket = sirket_adi if sirket_adi else "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ"
    aktif_is = is_adi if is_adi else ""

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
    run_sirket = p_sirket.add_run(aktif_sirket.upper())
    run_sirket.font.size = Pt(13)  # Uzun ünvan için en ideal boyut
    run_sirket.font.bold = True
    run_sirket.font.name = 'Arial'
    
    doc.add_paragraph()
    doc.add_paragraph()

    # Eğer proje adı girildiyse kapakta gösterelim, girilmediyse boş bırakalım
    if aktif_is:
        p_is = doc.add_paragraph()
        p_is.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_is_baslik = p_is.add_run("PROJE ADI:\n")
        run_is_baslik.font.size = Pt(11)
        run_is_baslik.font.name = 'Arial'
        
        run_is = p_is.add_run(aktif_is)
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
    run_hazirlayan = p_alt.add_run(f"Hazırlayan:\n{hazirlayan} (Makine Mühendisi)\nMMO Oda No: {mmo_no}\n\nTarih:\n{tarih}")
    run_hazirlayan.font.size = Pt(11)
    run_hazirlayan.font.name = 'Arial'

    # Yeni sayfaya geçiş (Genel Bilgiler için)
    doc.add_page_break()

    # --- BÖLÜM 1: GENEL BİLGİLER ---
    doc.add_heading("1. GENEL BİLGİLER", level=1)
    
    # 1. Sabit Giriş Metni
    proje_ifade = f"'{aktif_is}'" if aktif_is else "ilgili proje"
    giris_metni = f"Bu raporda {proje_ifade} için tasarlanan mekanik tesisatlar açıklanmış ve tüm uygulama ve detay projelerine esas teşkil eden tasarım kriterleri ve mekanik tesisat sistem çözümleri tespit edilmiştir."
    doc.add_paragraph(giris_metni)
    
    # 2. Mahal Cümlesi
    sehir_ifadesi = f"{sehir}'nda" if sehir else "''de"
    yapi_metni = f"Yapı {sehir_ifadesi} inşa edilecektir. Yapıda aşağıdaki mahaller bulunmaktadır."
    doc.add_paragraph(yapi_metni)

    # Hafızada dosya oluşturma
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.success("Word dosyası başarıyla hazırlandı!")
    
    dosya_adi = f"{aktif_is.replace(' ', '_')}_Genel_Bilgiler.docx" if is_adi else "Mekanik_Uygulama_Raporu.docx"
    
    st.download_button(
        label="📥 Word Dosyasını İndir (.docx)",
        data=buffer,
        file_name=dosya_adi,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
