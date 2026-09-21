from datetime import datetime
import io
import json
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import streamlit as st

# Türkçe ay isimleri için sözlük
aylar = {
    1: "Ocak",
    2: "Şubat",
    3: "Mart",
    4: "Nisan",
    5: "Mayıs",
    6: "Haziran",
    7: "Temmuz",
    8: "Ağustos",
    9: "Eylül",
    10: "Ekim",
    11: "Kasım",
    12: "Aralık",
}

bugun = datetime.now()
bugun_ay_yil = f"{aylar[bugun.month]} {bugun.year}"


# İklim verilerini JSON dosyasından yükleme fonksiyonu
def iklim_verisini_yukle():
  dosya_adi = "iklim_verileri.json"
  if os.path.exists(dosya_adi):
    with open(dosya_adi, "r", encoding="utf-8") as f:
      return json.load(f)
  else:
    # Dosya henüz yoksa varsayılan geçici yapı
    return {
        "Ankara": {
            "Çankaya": {
                "kis_kt": -12.0,
                "kis_yt": -13.2,
                "yaz_kt": 33.0,
                "yaz_yt": 19.0,
                "enlem": "39° 57' Kuzey",
                "boylam": "32° 53' Doğu",
                "rakim": 949,
                "gsf": 15.5,
            },
            "Keçiören": {
                "kis_kt": -12.5,
                "kis_yt": -13.7,
                "yaz_kt": 32.5,
                "yaz_yt": 18.5,
                "enlem": "39° 58' Kuzey",
                "boylam": "32° 51' Doğu",
                "rakim": 930,
                "gsf": 16.0,
            },
        },
        "İstanbul": {
            "Kadıköy": {
                "kis_kt": -2.0,
                "kis_yt": -3.5,
                "yaz_kt": 31.0,
                "yaz_yt": 23.0,
                "enlem": "40° 59' Kuzey",
                "boylam": "29° 02' Doğu",
                "rakim": 30,
                "gsf": 9.0,
            }
        },
    }


iklim_veritabani = iklim_verisini_yukle()


# Otomatik İçindekiler Tablosu (TOC) Alanı Ekleyen Fonksiyon
def add_toc(paragraph):
  run = paragraph.add_run()
  fldChar1 = OxmlElement("w:fldChar")
  fldChar1.set(qn("w:fldCharType"), "begin")
  instrText = OxmlElement("w:instrText")
  instrText.set(qn("xml:space"), "preserve")
  instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
  fldChar2 = OxmlElement("w:fldChar")
  fldChar2.set(qn("w:fldCharType"), "separate")
  fldChar3 = OxmlElement("w:fldChar")
  fldChar3.set(qn("w:fldCharType"), "end")

  r = run._r
  r.append(fldChar1)
  r.append(instrText)
  r.append(fldChar2)
  r.append(fldChar3)


st.title("Mühendislik Proje Raporu Otomasyonu")
st.write("Lütfen kurumsal kapak ve ilgili proje bölümlerini doldurun:")

# --- 1. SEKME / BÖLÜM: KAPAK BİLGİLERİ ---
st.header("1. Kapak Bilgileri")
sirket_adi = st.text_input(
    "Şirket / Kuruluş İsmi",
    "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ",
)
is_adi = st.text_input("İşin Adı / Proje Başlığı", "")
rapor_turu = st.text_input(
    "Rapor Türü", "MEKANİK TESİSAT UYGULAMA PROJESİ HESAP RAPORU"
)
hazirlayan = st.text_input("Hazırlayan Mühendis", "Mehmet Küçük")
mmo_no = st.text_input("MMO Oda No", "109913")
tarih = st.text_input("Rapor Tarihi", bugun_ay_yil)

# --- 2. SEKME / BÖLÜM: UYGULANACAK STANDART VE YÖNETMELİKLER ---
st.header("2. UYGULANACAK STANDART VE YÖNETMELİKLER")
st.write("Raporda yer almasını istediğiniz standart ve yönetmelikleri seçin:")

std_ts_825 = st.checkbox("TS 825 - BİNALARDA ISI YALITIM KURALLARI", value=True)
std_yangin = st.checkbox(
    '09 Eylül 2009 tarih ve 27344 numaralı sayısında yayımlanan " BİNALARIN'
    ' YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK"',
    value=True,
)
std_bep_2008_2010 = st.checkbox(
    "5 Aralık 2008 tarih, 27075 sayılı resmi gazetede yayımlanan “BİNALARDA"
    " ENERJİ PERFORMANSI YÖNETMELİĞİ” ve 1 Nisan 2010 tarih, 27539 sayılı resmi"
    " gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ”",
    value=True,
)
std_ts_1258 = st.checkbox(
    "TS 1258 – TEMİZSU TESİSATI HESAP KURALLARI", value=True
)
std_ts_826 = st.checkbox(
    "TS 826 – BİNALARDA PİSSU TESİSATI HESAPLAMA KURALLARI", value=True
)
std_ts_2164 = st.checkbox(
    "TS 2164 - KALORİFER TESİSATI PROJELENDİRME KURALLARI", value=True
)
std_ts_3419 = st.checkbox(
    "TS 3419 – HAVALANDIRMA VE İKLİMLENDİRME TESİSLERİ PROJELENDİRME"
    " KURALLARI",
    value=True,
)
std_ts_en_12056_2 = st.checkbox(
    "TS EN 12056-2 – CAZİBELİ DRENAJ SİSTEMLERİ -BİNA İÇİ- TASARIM VE HESAPLAMA",
    value=True,
)
std_ts_en_12845 = st.checkbox(
    "TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ – OTOMATİK SPRİNKLER"
    " SİSTEMLERİ- TASARIM, MONTAJ VE BAKIM",
    value=True,
)
std_mmo_84 = st.checkbox(
    "MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:84)", value=True
)
std_mmo_352_5 = st.checkbox("MMO KALORİFER TESİSATI (Y.NO:352/5)", value=True)
std_mmo_122 = st.checkbox(
    "MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI(Y.NO:122)", value=True
)
std_mmo_133 = st.checkbox(
    "MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:133)", value=True
)
std_mmo_155 = st.checkbox("MMO KAZAN VE BACA(Y.NO:155)", value=True)

std_ashrae = st.checkbox("ASHRAE Standartları", value=True)
std_su = st.checkbox(
    "İçmesuyu Temizleme ve Dağıtım Sistemleri Standartları", value=True
)
std_klima = st.checkbox(
    "Klima ve Havalandırma Tesisatı Yönetmelikleri", value=True
)
std_tesisat = st.checkbox(
    "Merkezi Isıtma ve Sıhhi Sıcak Su Sistemlerinde Isı Maliyetlerinin"
    " Paylaştırılmasına İlişkin Yönetmelik",
    value=False,
)
std_kanal = st.checkbox(
    "Kanalizasyon Şebekesi Olmayan Yerlerde Yapılacak Çukurlar", value=False
)
std_asansor = st.checkbox(
    "Asansör Yönetmeliği ve İlgili Standartlar", value=False
)
std_deprem = st.checkbox(
    "Türkiye Bina Deprem Yönetmeliği (Mekanik Ekipman Askı ve Destekleri)",
    value=True,
)
std_akustik = st.checkbox(
    "Binaların Gürültüye Karşı Korunması Yönetmeliği", value=False
)
std_isg = st.checkbox(
    "İş Sağlığı ve Güvenliği Kanunu ve İlgili Yönetmelikler", value=True
)

ek_standartlar = st.text_area(
    "Eklemek istediğiniz ilave standartlar ve açıklamaları (Her satıra bir tane"
    " yazabilirsiniz)",
    "",
    height=80,
)

# --- 3. SEKME / BÖLÜM: MEKANİK TESİSAT PROJE KAPSAMI ---
st.header("3. MEKANİK TESİSAT PROJE KAPSAMI")
st.write(
    "Proje kapsamında yer alacak mekanik tesisat sistemlerini seçebilirsiniz:"
)

kapsam_isitma = st.checkbox("Isıtma tesisatı,", value=True)
kapsam_sogutma = st.checkbox("Soğutma tesisatı,", value=True)
kapsam_soguk_su = st.checkbox("Kullanma soğuk suyu tesisatı,", value=True)
kapsam_sicak_su = st.checkbox("Kullanma sıcak suyu tesisatı,", value=True)
kapsam_yangin_depo = st.checkbox(
    "Yangın ve kullanma suyu depolaması ve dağıtımı,", value=True
)
kapsam_atik_su = st.checkbox(
    "Yapı içinde atık su tesisatı (Yapı çıkış rögarına),", value=True
)
kapsam_yangin_dagitim = st.checkbox(
    "Yangın suyu iç ve dış dağıtım sistemleri,", value=True
)
kapsam_kazan_dairesi = st.checkbox(
    "Merkezi ısıtma kazan dairesi ve tali teknik hacimler,", value=True
)
kapsam_havalandirma = st.checkbox("Havalandırma Tesisatı", value=True)
kapsam_basinc_hava = st.checkbox("Basınçlı hava tesisatı,", value=False)
kapsam_medikal_gaz = st.checkbox("Medikal gaz tesisatı", value=False)
kapsam_otomatik = st.checkbox(
    "Otomatik kontrol sistemi kavramı tanımı,", value=True
)

ek_kapsam = st.text_area(
    "Eklemek istediğiniz ilave proje kapsam maddeleri (Her satıra bir tane"
    " yazabilirsiniz)",
    "",
    height=80,
)

# --- 4. SEKME / BÖLÜM: TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI ---
st.header("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI")
st.write(
    "Raporda yer almasını istediğiniz ısı iletim akışkanlarını seçin ve"
    " rejimlerini belirleyin:"
)

sicaklik_secenekleri = [
    "80/60",
    "70/50",
    "60/40",
    "50/30",
    "50/40",
    "7/12",
    "6/11",
    "10/60",
]
buhar_secenekleri = [
    "1 atm (100 °C)",
    "2 bar (120 °C)",
    "3 bar (133 °C)",
    "4 bar (143 °C)",
    "6 bar (165 °C)",
    "8 bar (175 °C)",
]
kizgin_su_secenekleri = [
    "120/90",
    "130/70",
    "140/90",
    "150/100",
    "160/110",
    "180/130",
]

col1, col2 = st.columns(2)

with col1:
  chk_kalorifer = st.checkbox("1. Kalorifer tesisatı", value=True)
  rej_kalorifer = st.selectbox(
      "Kalorifer Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_fco_ist = st.checkbox("2. Fan-Coil ısıtma tesisatı", value=True)
  rej_fco_ist = st.selectbox(
      "Fan-Coil Isıtma Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_fco_sog = st.checkbox("3. Fan-Coil Soğutma tesisatı", value=True)
  rej_fco_sog = st.selectbox(
      "Fan-Coil Soğutma Rejimi:", sicaklik_secenekleri, index=5
  )

  chk_ks_ist = st.checkbox("4. Klima santrali ısıtma tesisatı", value=True)
  rej_ks_ist = st.selectbox(
      "Klima Santrali Isıtma Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_buhar = st.checkbox("9. Buhar tesisatı", value=False)
  rej_buhar = st.selectbox("Buhar Seçimi:", buhar_secenekleri, index=1)

with col2:
  chk_ks_sog = st.checkbox("5. Klima santrali Soğutma tesisatı", value=True)
  rej_ks_sog = st.selectbox(
      "Klima Santrali Soğutma Rejimi:", sicaklik_secenekleri, index=5
  )

  chk_boyler = st.checkbox("6. Boyler ısıtma tesisatı", value=True)
  rej_boyler = st.selectbox(
      "Boyler Isıtma Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_k_sicak = st.checkbox("7. Kullanma sıcak suyu", value=True)
  rej_k_sicak = st.selectbox(
      "Kullanma Sıcak Suyu Rejimi:", sicaklik_secenekleri, index=7
  )

  chk_doseme = st.checkbox("8. Döşemeden ısıtma tesisatı", value=True)
  rej_doseme = st.selectbox(
      "Döşemeden Isıtma Rejimi:", sicaklik_secenekleri, index=4
  )

  chk_kizgin = st.checkbox("10. Kızgın su tesisatı", value=False)
  rej_kizgin = st.selectbox(
      "Kızgın Su Rejimi:", kizgin_su_secenekleri, index=0
  )

# --- 5. BÖLÜM: İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ ---
st.header("5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ")
st.subheader("5.1 DIŞ HAVA TASARIM KRİTERLERİ")

# İl ve İlçe Seçimi (Dinamik Yapı)
iller_listesi = sorted(list(iklim_veritabani.keys()))
secilen_il = st.selectbox(
    "Yapının inşa edileceği ili seçin:", iller_listesi, index=0
)

ilceler_listesi = sorted(list(iklim_veritabani[secilen_il].keys()))
secilen_ilce = st.selectbox(
    "Yapının inşa edileceği ilçeyi seçin:", ilceler_listesi
)

# Seçilen ilçe verilerini çek
iklim_veri = iklim_veritabani[secilen_il][secilen_ilce]

# Arayüzde anlık gösterim
st.markdown(f"### 📌 {secilen_il} / {secilen_ilce} İklim Verileri")
c_1, c_2 = st.columns(2)
with c_1:
  st.markdown(
      f"• **KIŞ**: `{iklim_veri['kis_kt']}` °C KT , `{iklim_veri['kis_yt']}`"
      " °C YT"
  )
  st.markdown(
      f"• **YAZ**: `{iklim_veri['yaz_kt']}` °C KT , `{iklim_veri['yaz_yt']}`"
      " °C YT"
  )
  st.markdown(f"• **Günlük Sıcaklık Farkı**: `{iklim_veri['gsf']}` °C")
with c_2:
  st.markdown(f"• **Enlem**: `{iklim_veri['enlem']}`")
  st.markdown(f"• **Boylam**: `{iklim_veri['boylam']}`")
  st.markdown(
      f"• **Deniz seviyesinden yüksekliği**: `{iklim_veri['rakim']}` m."
  )

# --- 6. BÖLÜM: SIHHİ TESİSAT ---
st.header("6. SIHHİ TESİSAT")
st.write(
    "Sıhhi tesisat sistemi için raporda yer almasını istediğiniz kriterleri"
    " seçin:"
)

sihhi_su_temini = st.checkbox(
    "Şebeke suyu veya hidrofor sistemi ile basınçlı su temini", value=True
)
sihhi_sicak_su = st.checkbox(
    "Merkezi / bireysel sistem ile sıhhi sıcak su hazırlanması ve sirkülasyonu",
    value=True,
)
sihhi_pis_su = st.checkbox(
    "Bina içi atık su (pis su) ve havalandırma boratları", value=True
)
sihhi_yagmur = st.checkbox(
    "Çatı yağmur suyu drenajı ve uzaklaştırılması", value=True
)

ek_sihhi = st.text_area(
    "Eklemek istediğiniz ilave sıhhi tesisat esasları (Her satıra bir tane"
    " yazabilirsiniz)",
    "",
    height=80,
)

# Rapor Oluştur Butonu
if st.button("Raporu Oluştur (.docx)"):
  if not is_adi:
    st.warning(
        "⚠️ Dikkat: İşin Adı / Proje Başlığı girilmedi. Rapor oluşturuluyor"
        " ancak kapak başlığı boş bırakılacak."
    )

  aktif_sirket = (
      sirket_adi
      if sirket_adi
      else "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ"
  )
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
  run_sirket.font.name = "Arial"

  doc.add_paragraph()
  doc.add_paragraph()

  if aktif_is:
    p_is = doc.add_paragraph()
    p_is.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_is_baslik = p_is.add_run("PROJE ADI:\n")
    run_is_baslik.font.size = Pt(11)
    run_is_baslik.font.name = "Arial"

    run_is = p_is.add_run(aktif_is)
    run_is.font.size = Pt(16)
    run_is.font.bold = True
    run_is.font.name = "Arial"

    doc.add_paragraph()

  p_tur = doc.add_paragraph()
  p_tur.alignment = WD_ALIGN_PARAGRAPH.CENTER
  run_tur = p_tur.add_run(rapor_turu.upper())
  run_tur.font.size = Pt(14)
  run_tur.font.bold = True
  run_tur.font.name = "Arial"

  for _ in range(4):
    doc.add_paragraph()

  p_alt = doc.add_paragraph()
  p_alt.alignment = WD_ALIGN_PARAGRAPH.CENTER
  run_hazirlayan = p_alt.add_run(
      f"Hazırlayan:\n{hazirlayan} (Makine Mühendisi)\nMMO Oda No:"
      f" {mmo_no}\n\nTarih:\n{tarih}"
  )
  run_hazirlayan.font.size = Pt(11)
  run_hazirlayan.font.name = "Arial"

  # ==========================================
  # 2. SAYFA: İÇİNDEKİLER SAYFASI
  # ==========================================
  doc.add_page_break()

  doc.add_heading("İÇİNDEKİLER", level=1)

  p_toc = doc.add_paragraph()
  add_toc(p_toc)

  p_bilgi_notu = doc.add_paragraph()
  run_not = p_bilgi_notu.add_run(
      "(Not: Belgeyi Word'de açtığınızda üstüne sağ tıklayıp 'Alanı Güncelle'"
      " diyerek başlıkları ve sayfa numaralarını güncelleyebilirsiniz.)"
  )
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
  giris_metni = (
      f"Bu raporda {proje_ifade} için tasarlanan mekanik tesisatlar"
      " açıklanmış ve tüm uygulama ve detay projelerine esas teşkil eden"
      " tasarım kriterleri ve mekanik tesisat sistem çözümleri tespit"
      " edilmiştir."
  )
  doc.add_paragraph(giris_metni)

  yapi_metni = (
      f"Yapı {secilen_il} ili {secilen_ilce} ilçesinde inşa edilecektir."
  )
  doc.add_paragraph(yapi_metni)

  # --- 2. UYGULANACAK STANDART VE YÖNETMELİKLER ---
  doc.add_heading("2. UYGULANACAK STANDART VE YÖNETMELİKLER", level=1)

  standart_giris = (
      "Bu projenin tasarım ve uygulamasında seçilen ulusal ve uluslararası"
      " standartlar ile yönetmelikler esas alınmıştır:"
  )
  doc.add_paragraph(standart_giris)

  secilen_standartlar = []
  if std_ts_825:
    secilen_standartlar.append("TS 825 - BİNALARDA ISI YALITIM KURALLARI")
  if std_yangin:
    secilen_standartlar.append(
        '09 Eylül 2009 tarih ve 27344 numaralı sayısında yayımlanan " BİNALARIN'
        ' YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK"'
    )
  if std_bep_2008_2010:
    secilen_standartlar.append(
        "5 Aralık 2008 tarih, 27075 sayılı resmi gazetede yayımlanan “BİNALARDA"
        " ENERJİ PERFORMANSI YÖNETMELİĞİ” ve 1 Nisan 2010 tarih, 27539 sayılı"
        " resmi gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ”"
    )
  if std_ts_1258:
    secilen_standartlar.append("TS 1258 – TEMİZSU TESİSATI HESAP KURALLARI")
  if std_ts_826:
    secilen_standartlar.append(
        "TS 826 – BİNALARDA PİSSU TESİSATI HESAPLAMA KURALLARI"
    )
  if std_ts_2164:
    secilen_standartlar.append(
        "TS 2164 - KALORİFER TESİSATI PROJELENDİRME KURALLARI"
    )
  if std_ts_3419:
    secilen_standartlar.append(
        "TS 3419 – HAVALANDIRMA VE İKLİMLENDİRME TESİSLERİ PROJELENDİRME"
        " KURALLARI"
    )
  if std_ts_en_12056_2:
    secilen_standartlar.append(
        "TS EN 12056-2 – CAZİBELİ DRENAJ SİSTEMLERİ -BİNA İÇİ- TASARIM VE"
        " HESAPLAMA"
    )
  if std_ts_en_12845:
    secilen_standartlar.append(
        "TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ – OTOMATİK SPRİNKLER"
        " SİSTEMLERİ- TASARIM, MONTAJ VE BAKIM"
    )
  if std_mmo_84:
    secilen_standartlar.append(
        "MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:84)"
    )
  if std_mmo_352_5:
    secilen_standartlar.append("MMO KALORİFER TESİSATI (Y.NO:352/5)")
  if std_mmo_122:
    secilen_standartlar.append(
        "MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI(Y.NO:122)"
    )
  if std_mmo_133:
    secilen_standartlar.append(
        "MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:133)"
    )
  if std_mmo_155:
    secilen_standartlar.append("MMO KAZAN VE BACA(Y.NO:155)")
  if std_ashrae:
    secilen_standartlar.append("ASHRAE Standartları")
  if std_su:
    secilen_standartlar.append(
        "İçmesuyu Temizleme ve Dağıtım Sistemleri Standartları"
    )
  if std_klima:
    secilen_standartlar.append(
        "Klima ve Havalandırma Tesisatı Yönetmelikleri"
    )
  if std_tesisat:
    secilen_standartlar.append(
        "Merkezi Isıtma ve Sıhhi Sıcak Su Sistemlerinde Isı Maliyetlerinin"
        " Paylaştırılmasına İlişkin Yönetmelik"
    )
  if std_kanal:
    secilen_standartlar.append(
        "Kanalizasyon Şebekesi Olmayan Yerlerde Yapılacak Çukurlar"
    )
  if std_asansor:
    secilen_standartlar.append(
        "Asansör Yönetmeliği ve İlgili Standartlar"
    )
  if std_deprem:
    secilen_standartlar.append(
        "Türkiye Bina Deprem Yönetmeliği (Mekanik Ekipman Askı ve Destekleri)"
    )
  if std_akustik:
    secilen_standartlar.append(
        "Binaların Gürültüye Karşı Korunması Yönetmeliği"
    )
  if std_isg:
    secilen_standartlar.append(
        "İş Sağlığı ve Güvenliği Kanunu ve İlgili Yönetmelikler"
    )

  if ek_standartlar.strip():
    for ek in ek_standartlar.split("\n"):
      if ek.strip():
        secilen_standartlar.append(ek.strip())

  secilen_standartlar.sort()

  if secilen_standartlar:
    for std in secilen_standartlar:
      doc.add_paragraph(std, style="List Bullet")
  else:
    doc.add_paragraph("Herhangi bir standart seçilmemiştir.", style="Italic")

  # --- 3. MEKANİK TESİSAT PROJE KAPSAMI ---
  doc.add_heading("3. MEKANİK TESİSAT PROJE KAPSAMI", level=1)
  doc.add_paragraph("Yapılarda aşağıdaki mekanik tesisat sistemleri uygulanacaktır.")

  secilen_kapsam = []
  if kapsam_isitma:
    secilen_kapsam.append("Isıtma tesisatı,")
  if kapsam_sogutma:
    secilen_kapsam.append("Soğutma tesisatı,")
  if kapsam_soguk_su:
    secilen_kapsam.append("Kullanma soğuk suyu tesisatı,")
  if kapsam_sicak_su:
    secilen_kapsam.append("Kullanma sıcak suyu tesisatı,")
  if kapsam_yangin_depo:
    secilen_kapsam.append("Yangın ve kullanma suyu depolaması ve dağıtımı,")
  if kapsam_atik_su:
    secilen_kapsam.append("Yapı içinde atık su tesisatı (Yapı çıkış rögarına),")
  if kapsam_yangin_dagitim:
    secilen_kapsam.append("Yangın suyu iç ve dış dağıtım sistemleri,")
  if kapsam_kazan_dairesi:
    secilen_kapsam.append(
        "Merkezi ısıtma kazan dairesi ve tali teknik hacimler,"
    )
  if kapsam_havalandirma:
    secilen_kapsam.append("Havalandırma Tesisatı")
  if kapsam_basinc_hava:
    secilen_kapsam.append("Basınçlı hava tesisatı,")
  if kapsam_medikal_gaz:
    secilen_kapsam.append("Medikal gaz tesisatı")
  if kapsam_otomatik:
    secilen_kapsam.append("Otomatik kontrol sistemi kavramı tanımı,")

  if ek_kapsam.strip():
    for ekk in ek_kapsam.split("\n"):
      if ekk.strip():
        secilen_kapsam.append(ekk.strip())

  if secilen_kapsam:
    for k in secilen_kapsam:
      doc.add_paragraph(k, style="List Bullet")
  else:
    doc.add_paragraph(
        "Herhangi bir proje kapsam maddesi seçilmemiştir.", style="Italic"
    )

  # --- 4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI ---
  doc.add_heading("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI", level=1)
  doc.add_paragraph(
      "Tesisat sistemlerinde aşağıdaki ısı iletim akışkanları ve sıcaklık"
      " rejimleri kullanılacaktır:"
  )

  akiskan_maddeleri = []
  if chk_kalorifer:
    akiskan_maddeleri.append(
        f"Kalorifer tesisatında {rej_kalorifer} °C sıcak su."
    )
  if chk_fco_ist:
    akiskan_maddeleri.append(
        f"Fan-Coil ısıtma tesisatında {rej_fco_ist} °C sıcak su."
    )
  if chk_fco_sog:
    akiskan_maddeleri.append(
        f"Fan-Coil Soğutma tesisatında {rej_fco_sog} °C soğuk su."
    )
  if chk_ks_ist:
    akiskan_maddeleri.append(
        f"Klima santrali ısıtma tesisatında {rej_ks_ist} °C sıcak su."
    )
  if chk_ks_sog:
    akiskan_maddeleri.append(
        f"Klima santrali Soğutma tesisatında {rej_ks_sog} °C soğuk su."
    )
  if chk_boyler:
    akiskan_maddeleri.append(
        f"Boyler ısıtma tesisatında {rej_boyler} °C sıcak su."
    )
  if chk_k_sicak:
    akiskan_maddeleri.append(
        f"Kullanma sıcak suyunda {rej_k_sicak} °C sıcak su."
    )
  if chk_doseme:
    akiskan_maddeleri.append(
        f"Döşemeden ısıtma tesisatında {rej_doseme} °C sıcak su."
    )
  if chk_buhar:
    akiskan_maddeleri.append(f"Buhar tesisatında {rej_buhar} buhar.")
  if chk_kizgin:
    akiskan_maddeleri.append(
        f"Kızgın su tesisatında {rej_kizgin} °C sıcak su."
    )

  if akiskan_maddeleri:
    for akiskan in akiskan_maddeleri:
      doc.add_paragraph(akiskan, style="List Bullet")
  else:
    doc.add_paragraph(
        "Herhangi bir ısı iletim akışkanı seçilmemiştir.", style="Italic"
    )

  # --- 5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ ---
  doc.add_heading("5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ", level=1)

  doc.add_heading("5.1 DIŞ HAVA TASARIM KRİTERLERİ", level=2)
  dih_hava_metni = (
      f"Yapının inşa edileceği {secilen_il} ili {secilen_ilce} ilçesi için"
      " kabul edilen dış hava koşulları ve coğrafi konum bilgileri"
      " aşağıdadır:"
  )
  doc.add_paragraph(dih_hava_metni)

  doc.add_paragraph(
      f"• KIŞ: {iklim_veri['kis_kt']} °C Kuru Termometre (KT) ,"
      f" {iklim_veri['kis_yt']} °C Yaş Termometre (YT)",
      style="List Bullet",
  )
  doc.add_paragraph(
      f"• YAZ: {iklim_veri['yaz_kt']} °C Kuru Termometre (KT) ,"
      f" {iklim_veri['yaz_yt']} °C Yaş Termometre (YT)",
      style="List Bullet",
  )
  doc.add_paragraph(f"• Enlem: {iklim_veri['enlem']}", style="List Bullet")
  doc.add_paragraph(f"• Boylam: {iklim_veri['boylam']}", style="List Bullet")
  doc.add_paragraph(
      f"• Deniz seviyesinden yüksekliği (Rakım): {iklim_veri['rakim']} m.",
      style="List Bullet",
  )
  doc.add_paragraph(
      f"• Günlük Sıcaklık Farkı (GSF): {iklim_veri['gsf']} °C",
      style="List Bullet",
  )

  # --- 6. SIHHİ TESİSAT ---
  doc.add_heading("6. SIHHİ TESİSAT", level=1)
  doc.add_paragraph(
      "Yapının sıhhi tesisat projelendirmesinde ilgili standartlar (TS 1258 ve"
      " TS 826) esas alınarak aşağıdaki sistemler tasarlanmıştır:"
  )

  secilen_sihhi = []
  if sihhi_su_temini:
    secilen_sihhi.append(
        "Şebeke suyu veya hidrofor sistemi ile basınçlı su temini,"
    )
  if sihhi_sicak_su:
    secilen_sihhi.append(
        "Merkezi / bireysel sistem ile sıhhi sıcak su hazırlanması ve"
        " sirkülasyonu,"
    )
  if sihhi_pis_su:
    secilen_sihhi.append("Bina içi atık su (pis su) ve havalandırma boratları,")
  if sihhi_yagmur:
    secilen_sihhi.append("Çatı yağmur suyu drenajı ve uzaklaştırılması,")

  if ek_sihhi.strip():
    for esh in ek_sihhi.split("\n"):
      if esh.strip():
        secilen_sihhi.append(esh.strip())

  if secilen_sihhi:
    for sih in secilen_sihhi:
      doc.add_paragraph(sih, style="List Bullet")
  else:
    doc.add_paragraph(
        "Herhangi bir sıhhi tesisat kapsam maddesi seçilmemiştir.", style="Italic"
    )

  # Hafızada dosya oluşturma
  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)

  st.success(
      "Sıhhi tesisat ve tüm önceki bölümler eklenerek rapor hazırlandı!"
  )

  dosya_adi = (
      f"{aktif_is.replace(' ', '_')}_Rapor.docx"
      if is_adi
      else "Mekanik_Uygulama_Raporu.docx"
  )

  st.download_button(
      label="📥 Word Dosyasını İndir (.docx)",
      data=buffer,
      file_name=dosya_adi,
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
  )
