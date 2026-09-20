import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
from datetime import datetime

# Türkçe ay isimleri için sözlük
aylar = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
}

bugun = datetime.now()
bugun_ay_yil = f"{aylar[bugun.month]} {bugun.year}"

# Otomatik İçindekiler Tablosu (TOC) Alanı Ekleyen Fonksiyon
def add_toc(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)

st.title("Mühendislik Proje Raporu Otomasyonu")
st.write("Lütfen kurumsal kapak ve ilgili proje bölümlerini doldurun:")

# --- 1. SEKME / BÖLÜM: KAPAK BİLGİLERİ ---
st.header("1. Kapak Bilgileri")
sirket_adi = st.text_input("Şirket / Kuruluş İsmi", "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ")
is_adi = st.text_input("İşin Adı / Proje Başlığı", "")
rapor_turu = st.text_input("Rapor Türü", "MEKANİK TESİSAT UYGULAMA PROJESİ HESAP RAPORU")
hazirlayan = st.text_input("Hazırlayan Mühendis", "Mehmet Küçük")
mmo_no = st.text_input("MMO Oda No", "109913")
tarih = st.text_input("Rapor Tarihi", bugun_ay_yil)

# --- 2. SEKME / BÖLÜM: UYGULANACAK STANDART VE YÖNETMELİKLER ---
st.header("2. UYGULANACAK STANDART VE YÖNETMELİKLER")
st.write("Projeye esas alınacak standart ve yönetmelikleri aşağıda düzenleyebilirsiniz:")

varsayilan_standartlar = (
    "• 5 Aralık 2008 tarih, 27075 sayılı resmi gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ” şartları ve 1 Nisan 2010 tarih, 27539 sayılı resmi gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ”\n"
    "• 09 Eylül 2009 tarih ve 27344 numaralı sayısında yayımlanan \" BİNALARIN YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK\"\n"
    "• TS 825 - BİNALARDA ISI YALITIM KURALLARI\n"
    "• TS 1258 – TEMİZSU TESİSATI HESAP KURALLARI\n"
    "• TS 826 – BİNALARDA PİSSU TESİSATI HESAPLAMA KURALLARI\n"
    "• TS 2164 - KALORİFER TESİSATI PROJELENDİRME KURALLARI\n"
    "• TS3419–HAVALANDIRMA VE İKLİMLENDİRME TESİSLERİ PROJELENDİRME KURALLARI\n"
    "• TS EN 12056-2 – CAZİBELİ DRENAJ SİSTEMLERİ -BİNA İÇİ- TASARIM VE HESAPLAMA\n"
    "• TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ – OTOMATİK SPRİNKLER SİSTEMLERİ- TASARIM, MONTAJ VE BAKIM\n"
    "• MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:84)\n"
    "• MMO KALORİFER TESİSATI (Y.NO:352/5)\n"
    "• MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI(Y.NO:122)\n"
    "• MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:133)\n"
    "• MMO KAZAN VE BACA(Y.NO:155)"
)
standartlar_input = st.text_area("Standart ve Yönetmelik Listesi", varsayilan_standartlar, height=250)

sehir = "" 

# Rapor Oluştur Butonu
if st.button("Raporu Oluştur (.docx)"):
    if not is_adi:
        st.warning("⚠️ Dikkat: İşin Adı / Proje Başlığı girilmedi. Rapor oluşturuluyor ancak kapak başlığı boş bırakılacak.")
    
    aktif_sirket = sirket_adi if sirket_adi else "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ"
    aktif_is = is_adi if is_adi else ""

    doc = Document()
    
    # Sayfa boşlukları
    cover_section = doc.sections[0]
    cover_section.top_margin = Inches(1.5)
    cover_section.bottom_margin = Inches(1.5)
    cover_section.left_margin = Inches(1.2)
    cover_section.right_margin = Inches(1.2)

    # ==========================================
    # 1. SAYFA: KAPAK SAYFASI
    # ==========================================
    p_sirket = doc.add_paragraph()
    p_sirket.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sirket = p_sirket.add_run(aktif_sirket.upper())
    run_sirket.font.size = Pt(13)
    run_sirket.font.bold = True
    run_sirket.font.name = 'Arial'
    
    doc.add_paragraph()
    doc.add_paragraph()

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

    # ==========================================
    # 2. SAYFA: İÇİNDEKİLER SAYFASI
    # ==========================================
    doc.add_page_break()
    
    doc.add_heading("İÇİNDEKİLER", level=1)
    
    p_toc = doc.add_paragraph()
    add_toc(p_toc)
    
    p_bilgi_notu = doc.add_paragraph()
    run_not = p_bilgi_notu.add_run("(Not: Belgeyi Word'de açtığınızda üstüne sağ tıklayıp 'Alanı Güncelle' diyerek başlıkları ve sayfa numaralarını güncelleyebilirsiniz.)")
    run_not.font.size = Pt(9)
    run_not.font.italic = True
    run_not.font.color.rgb = RGBColor(128, 128, 128)

    # ==========================================
    # 3. SAYFA: GENEL BİLGİLER VE STANDARTLAR
    # ==========================================
    doc.add_page_break()

    body_section = doc.add_section()
    body_section.top_margin = Inches(1.2)
    body_section.bottom_margin = Inches(1.2)
    body_section.left_margin = Inches(1.2)
    body_section.right_margin = Inches(1.2)

    doc.add_heading("1. GENEL BİLGİLER", level=1)
    
    proje_ifade = f"'{aktif_is}'" if aktif_is else "ilgili proje"
    giris_metni = f"Bu raporda {proje_ifade} için tasarlanan mekanik tesisatlar açıklanmış ve tüm uygulama ve detay projelerine esas teşkil eden tasarım kriterleri ve mekanik tesisat sistem çözümleri tespit edilmiştir."
    doc.add_paragraph(giris_metni)
    
    yapi_metni = f"Yapı {sehir if sehir else '...'} 'nda inşa edilecektir. Yapıda aşağıdaki mahaller bulunmaktadır."
    doc.add_paragraph(yapi_metni)

    # --- UYGULANACAK STANDART VE YÖNETMELİKLER BÖLÜMÜ ---
    doc.add_heading("2. UYGULANACAK STANDART VE YÖNETMELİKLER", level=1)
    
    standart_giris = "Bu projenin tasarım ve uygulamasında aşağıda belirtilen ulusal ve uluslararası standartlar ile yönetmelikler esas alınmıştır:"
    doc.add_paragraph(standart_giris)
    
    # Madde işaretlerini temizleyip docx listesi olarak ekleyelim (ya da doğrudan paragraf olarak)
    standartlar_listesi = [s.strip().lstrip("•").strip() for s in standartlar_input.split("\n") if s.strip()]
    for std in standartlar_listesi:
        doc.add_paragraph(std, style='List Bullet')

    # Hafızada dosya oluşturma
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.success("Standartlar ve yönetmelikler listesi eklenerek Word dosyası hazırlandı!")
    
    dosya_adi = f"{aktif_is.replace(' ', '_')}_Rapor.docx" if is_adi else "Mekanik_Uygulama_Raporu.docx"
    
    st.download_button(
        label="📥 Word Dosyasını İndir (.docx)",
        data=buffer,
        file_name=dosya_adi,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
