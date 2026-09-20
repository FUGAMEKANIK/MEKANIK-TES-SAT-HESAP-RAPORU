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
st.write("Raporda yer almasını istediğiniz standart ve yönetmelikleri seçin:")

std_ts_825 = st.checkbox("TS 825 - BİNALARDA ISI YALITIM KURALLARI", value=True)
std_yangin = st.checkbox('09 Eylül 2009 tarih ve 27344 numaralı sayısında yayımlanan " BİNALARIN YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK"', value=True)
std_bep_2008_2010 = st.checkbox("5 Aralık 2008 tarih, 27075 sayılı resmi gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ” ve 1 Nisan 2010 tarih, 27539 sayılı resmi gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ”", value=True)
std_ts_1258 = st.checkbox("TS 1258 – TEMİZSU TESİSATI HESAP KURALLARI", value=True)
std_ts_826 = st.checkbox("TS 826 – BİNALARDA PİSSU TESİSATI HESAPLAMA KURALLARI", value=True)
std_ts_2164 = st.checkbox("TS 2164 - KALORİFER TESİSATI PROJELENDİRME KURALLARI", value=True)
std_ts_3419 = st.checkbox("TS 3419 – HAVALANDIRMA VE İKLİMLENDİRME TESİSLERİ PROJELENDİRME KURALLARI", value=True)
std_ts_en_12056_2 = st.checkbox("TS EN 12056-2 – CAZİBELİ DRENAJ SİSTEMLERİ -BİNA İÇİ- TASARIM VE HESAPLAMA", value=True)
std_ts_en_12845 = st.checkbox("TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ – OTOMATİK SPRİNKLER SİSTEMLERİ- TASARIM, MONTAJ VE BAKIM", value=True)
std_mmo_84 = st.checkbox("MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:84)", value=True)
std_mmo_352_5 = st.checkbox("MMO KALORİFER TESİSATI (Y.NO:352/5)", value=True)
std_mmo_122 = st.checkbox("MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI(Y.NO:122)", value=True)
std_mmo_133 = st.checkbox("MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:133)", value=True)
std_mmo_155 = st.checkbox("MMO KAZAN VE BACA(Y.NO:155)", value=True)

std_ashrae = st.checkbox("ASHRAE Standartları", value=True)
std_su = st.checkbox("İçmesuyu Temizleme ve Dağıtım Sistemleri Standartları", value=True)
std_klima = st.checkbox("Klima ve Havalandırma Tesisatı Yönetmelikleri", value=True)
std_tesisat = st.checkbox("Merkezi Isıtma ve Sıhhi Sıcak Su Sistemlerinde Isı Maliyetlerinin Paylaştırılmasına İlişkin Yönetmelik", value=False)
std_kanal = st.checkbox("Kanalizasyon Şebekesi Olmayan Yerlerde Yapılacak Çukurlar", value=False)
std_asansor = st.checkbox("Asansör Yönetmeliği ve İlgili Standartlar", value=False)
std_deprem = st.checkbox("Türkiye Bina Deprem Yönetmeliği (Mekanik Ekipman Askı ve Destekleri)", value=True)
std_akustik = st.checkbox("Binaların Gürültüye Karşı Korunması Yönetmeliği", value=False)
std_isg = st.checkbox("İş Sağlığı ve Güvenliği Kanunu ve İlgili Yönetmelikler", value=True)

ek_standartlar = st.text_area("Eklemek istediğiniz ilave standartlar ve açıklamaları (Her satıra bir tane yazabilirsiniz)", "", height=80)

# --- 3. SEKME / BÖLÜM: MEKANİK TESİSAT PROJE KAPSAMI ---
st.header("3. MEKANİK TESİSAT PROJE KAPSAMI")
st.write("Proje kapsamında yer alacak mekanik tesisat sistemlerini seçebilirsiniz:")

kapsam_isitma = st.checkbox("Isıtma tesisatı,", value=True)
kapsam_sogutma = st.checkbox("Soğutma tesisatı,", value=True)
kapsam_soguk_su = st.checkbox("Kullanma soğuk suyu tesisatı,", value=True)
kapsam_sicak_su = st.checkbox("Kullanma sıcak suyu tesisatı,", value=True)
kapsam_yangin_depo = st.checkbox("Yangın ve kullanma suyu depolaması ve dağıtımı,", value=True)
kapsam_atik_su = st.checkbox("Yapı içinde atık su tesisatı (Yapı çıkış rögarına),", value=True)
kapsam_yangin_dagitim = st.checkbox("Yangın suyu iç ve dış dağıtım sistemleri,", value=True)
kapsam_kazan_dairesi = st.checkbox("Merkezi ısıtma kazan dairesi ve tali teknik hacimler,", value=True)
kapsam_havalandirma = st.checkbox("Havalandırma Tesisatı", value=True)
kapsam_basinc_hava = st.checkbox("Basınçlı hava tesisatı,", value=False)
kapsam_medikal_gaz = st.checkbox("Medikal gaz tesisatı", value=False)
kapsam_otomatik = st.checkbox("Otomatik kontrol sistemi kavramı tanımı,", value=True)

ek_kapsam = st.text_area("Eklemek istediğiniz ilave proje kapsam maddeleri (Her satıra bir tane yazabilirsiniz)", "", height=80)

# --- 4. SEKME / BÖLÜM: TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI ---
st.header("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI")
st.write("Tesisat sistemlerinde kullanılacak akışkan sıcaklık ve basınç rejimlerini sekmeli olarak yapılandırın:")

# Arayüz İçi Sekmeler (Tabs)
tab_isitma_sogutma, tab_buhar_kizgin = st.tabs(["Isıtma / Soğutma Tesisatları", "Buhar ve Kızgın Su Tesisatları"])

sicaklik_secenekleri = ["80/60", "70/50", "60/40", "50/30", "50/40", "7/12", "6/11", "10/60"]
buhar_secenekleri = ["1 atm (100 °C)", "2 bar (120 °C)", "3 bar (133 °C)", "4 bar (143 °C)", "6 bar (165 °C)", "8 bar (175 °C)"]
kizgin_su_secenekleri = ["120/90", "130/70", "140/90", "150/100", "160/110", "180/130"]

with tab_isitma_sogutma:
    col1, col2 = st.columns(2)
    with col1:
        rej_kalorifer = st.selectbox("Kalorifer tesisatı sıcaklık rejimi:", sicaklik_secenekleri, index=0)
        rej_fco_ist = st.selectbox("Fan-Coil ısıtma tesisatı sıcaklık rejimi:", sicaklik_secenekleri, index=0)
        rej_fco_sog = st.selectbox("Fan-Coil Soğutma tesisatı sıcaklık rejimi:", sicaklik_secenekleri, index=5)
        rej_ks_ist = st.selectbox("Klima santrali ısıtma tesisatı sıcaklık rejimi:", sicaklik_secenekleri, index=0)
    with col2:
        rej_ks_sog = st.selectbox("Klima santrali Soğutma tesisatı sıcaklık rejimi:", sicaklik_secenekleri, index=5)
        rej_boyler = st.selectbox("Boyler ısıtma tesisatı sıcaklık rejimi:", sicaklik_secenekleri, index=0)
        rej_k_sicak = st.selectbox("Kullanma sıcak suyu sıcaklık rejimi:", sicaklik_secenekleri, index=7)  # 10/60
        rej_doseme = st.selectbox("Döşemeden ısıtma tesisatı sıcaklık rejimi:", sicaklik_secenekleri, index=4)  # 50/40

with tab_buhar_kizgin:
    col3, col4 = st.columns(2)
    with col3:
        aktif_buhar = st.checkbox("Buhar tesisatı dahil et", value=False)
        if aktif_buhar:
            rej_buhar = st.selectbox("Buhar tesisatı basınç/sıcaklık seçimi:", buhar_secenekleri, index=1)
        else:
            rej_buhar = ""
    with col4:
        aktif_kizgin = st.checkbox("Kızgın su tesisatı dahil et", value=False)
        if aktif_kizgin:
            rej_kizgin = st.selectbox("Kızgın su tesisatı sıcaklık rejimi:", kizgin_su_secenekleri, index=0)
        else:
            rej_kizgin = ""

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
    # 3. SAYFA: GENEL BİLGİLER VE RAPOR İÇERİĞİ
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

    # --- 2. UYGULANACAK STANDART VE YÖNETMELİKLER ---
    doc.add_heading("2. UYGULANACAK STANDART VE YÖNETMELİKLER", level=1)
    
    standart_giris = "Bu projenin tasarım ve uygulamasında seçilen ulusal ve uluslararası standartlar ile yönetmelikler esas alınmıştır:"
    doc.add_paragraph(standart_giris)
    
    secilen_standartlar = []
    if std_ts_825: secilen_standartlar.append("TS 825 - BİNALARDA ISI YALITIM KURALLARI")
    if std_yangin: secilen_standartlar.append('09 Eylül 2009 tarih ve 27344 numaralı sayısında yayımlanan " BİNALARIN YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK"')
    if std_bep_2008_2010: secilen_standartlar.append("5 Aralık 2008 tarih, 27075 sayılı resmi gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ” ve 1 Nisan 2010 tarih, 27539 sayılı resmi gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ”")
    if std_ts_1258: secilen_standartlar.append("TS 1258 – TEMİZSU TESİSATI HESAP KURALLARI")
    if std_ts_826: secilen_standartlar.append("TS 826 – BİNALARDA PİSSU TESİSATI HESAPLAMA KURALLARI")
    if std_ts_2164: secilen_standartlar.append("TS 2164 - KALORİFER TESİSATI PROJELENDİRME KURALLARI")
    if std_ts_3419: secilen_standartlar.append("TS 3419 – HAVALANDIRMA VE İKLİMLENDİRME TESİSLERİ PROJELENDİRME KURALLARI")
    if std_ts_en_12056_2: secilen_standartlar.append("TS EN 12056-2 – CAZİBELİ DRENAJ SİSTEMLERİ -BİNA İÇİ- TASARIM VE HESAPLAMA")
    if std_ts_en_12845: secilen_standartlar.append("TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ – OTOMATİK SPRİNKLER SİSTEMLERİ- TASARIM, MONTAJ VE BAKIM")
    if std_mmo_84: secilen_standartlar.append("MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:84)")
    if std_mmo_352_5: secilen_standartlar.append("MMO KALORİFER TESİSATI (Y.NO:352/5)")
    if std_mmo_122: secilen_standartlar.append("MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI(Y.NO:122)")
    if std_mmo_133: secilen_standartlar.append("MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:133)")
    if std_mmo_155: secilen_standartlar.append("MMO KAZAN VE BACA(Y.NO:155)")
    if std_ashrae: secilen_standartlar.append("ASHRAE Standartları")
    if std_su: secilen_standartlar.append("İçmesuyu Temizleme ve Dağıtım Sistemleri Standartları")
    if std_klima: secilen_standartlar.append("Klima ve Havalandırma Tesisatı Yönetmelikleri")
    if std_tesisat: secilen_standartlar.append("Merkezi Isıtma ve Sıhhi Sıcak Su Sistemlerinde Isı Maliyetlerinin Paylaştırılmasına İlişkin Yönetmelik")
    if std_kanal: secilen_standartlar.append("Kanalizasyon Şebekesi Olmayan Yerlerde Yapılacak Çukurlar")
    if std_asansor: secilen_standartlar.append("Asansör Yönetmeliği ve İlgili Standartlar")
    if std_deprem: secilen_standartlar.append("Türkiye Bina Deprem Yönetmeliği (Mekanik Ekipman Askı ve Destekleri)")
    if std_akustik: secilen_standartlar.append("Binaların Gürültüye Karşı Korunması Yönetmeliği")
    if std_isg: secilen_standartlar.append("İş Sağlığı ve Güvenliği Kanunu ve İlgili Yönetmelikler")
    
    if ek_standartlar.strip():
        for ek in ek_standartlar.split("\n"):
            if ek.strip():
                secilen_standartlar.append(ek.strip())

    secilen_standartlar.sort()

    if secilen_standartlar:
        for std in secilen_standartlar:
            doc.add_paragraph(std, style='List Bullet')
    else:
        doc.add_paragraph("Herhangi bir standart seçilmemiştir.", style='Italic')

    # --- 3. MEKANİK TESİSAT PROJE KAPSAMI ---
    doc.add_heading("3. MEKANİK TESİSAT PROJE KAPSAMI", level=1)
    
    doc.add_paragraph("Yapılarda aşağıdaki mekanik tesisat sistemleri uygulanacaktır.")
    
    secilen_kapsam = []
    if kapsam_isitma: secilen_kapsam.append("Isıtma tesisatı,")
    if kapsam_sogutma: secilen_kapsam.append("Soğutma tesisatı,")
    if kapsam_soguk_su: secilen_kapsam.append("Kullanma soğuk suyu tesisatı,")
    if kapsam_sicak_su: secilen_kapsam.append("Kullanma sıcak suyu tesisatı,")
    if kapsam_yangin_depo: secilen_kapsam.append("Yangın ve kullanma suyu depolaması ve dağıtımı,")
    if kapsam_atik_su: secilen_kapsam.append("Yapı içinde atık su tesisatı (Yapı çıkış rögarına),")
    if kapsam_yangin_dagitim: secilen_kapsam.append("Yangın suyu iç ve dış dağıtım sistemleri,")
    if kapsam_kazan_dairesi: secilen_kapsam.append("Merkezi ısıtma kazan dairesi ve tali teknik hacimler,")
    if kapsam_havalandirma: secilen_kapsam.append("Havalandırma Tesisatı")
    if kapsam_basinc_hava: secilen_kapsam.append("Basınçlı hava tesisatı,")
    if kapsam_medikal_gaz: secilen_kapsam.append("Medikal gaz tesisatı")
    if kapsam_otomatik: secilen_kapsam.append("Otomatik kontrol sistemi kavramı tanımı,")
    
    if ek_kapsam.strip():
        for ekk in ek_kapsam.split("\n"):
            if ekk.strip():
                secilen_kapsam.append(ekk.strip())

    if secilen_kapsam:
        for k in secilen_kapsam:
            doc.add_paragraph(k, style='List Bullet')
    else:
        doc.add_paragraph("Herhangi bir proje kapsam maddesi seçilmemiştir.", style='Italic')

    # --- 4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI ---
    doc.add_heading("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI", level=1)
    
    doc.add_paragraph("Tesisat sistemlerinde aşağıdaki ısı iletim akışkanları ve sıcaklık rejimleri kullanılacaktır:")
    
    akiskan_maddeleri = [
        f"Kalorifer tesisatında {rej_kalorifer} °C sıcak su.",
        f"Fan-Coil ısıtma tesisatında {rej_fco_ist} °C sıcak su.",
        f"Fan-Coil Soğutma tesisatında {rej_fco_sog} °C soğuk su.",
        f"Klima santrali ısıtma tesisatında {rej_ks_ist} °C sıcak su.",
        f"Klima santrali Soğutma tesisatında {rej_ks_sog} °C soğuk su.",
        f"Boyler ısıtma tesisatında {rej_boyler} °C sıcak su.",
        f"Kullanma sıcak suyunda {rej_k_sicak} °C sıcak su.",
        f"Döşemeden ısıtma tesisatında {rej_doseme} °C sıcak su."
    ]

    if aktif_buhar and rej_buhar:
        akiskan_maddeleri.append(f"Buhar tesisatında {rej_buhar} buhar.")
    if aktif_kizgin and rej_kizgin:
        akiskan_maddeleri.append(f"Kızgın su tesisatında {rej_kizgin} °C sıcak su.")

    for akiskan in akiskan_maddeleri:
        doc.add_paragraph(akiskan, style='List Bullet')

    # Hafızada dosya oluşturma
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.success("Sekmeli Isı İletim Akışkanları ve Rejim seçimleri tamamlanarak Word dosyası hazırlandı!")
    
    dosya_adi = f"{aktif_is.replace(' ', '_')}_Rapor.docx" if is_adi else "Mekanik_Uygulama_Raporu.docx"
    
    st.download_button(
        label="📥 Word Dosyasını İndir (.docx)",
        data=buffer,
        file_name=dosya_adi,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
