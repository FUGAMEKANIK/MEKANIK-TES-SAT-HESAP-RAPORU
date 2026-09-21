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
            }
        }
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

# --- 1. KAPAK BİLGİLERİ ---
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

# --- 2. UYGULANACAK STANDART VE YÖNETMELİKLER ---
st.header("2. UYGULANACAK STANDART VE YÖNETMELİKLER")
std_ts_825 = st.checkbox("TS 825 - BİNALARDA ISI YALITIM KURALLARI", value=True)
std_yangin = st.checkbox(
    '09 Eylül 2009 tarih ve 27344 numaralı sayısında yayımlanan " BİNALARIN'
    ' YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK"',
    value=True,
)
std_bep_2008_2010 = st.checkbox(
    "5 Aralık 2008 tarih, 27075 sayılı resmi gazetede yayımlanan “BİNALARDA"
    " ENERJİ PERFORMANSI YÖNETMELİĞİ”",
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
    "TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ", value=True
)
std_mmo_84 = st.checkbox(
    "MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI (Y.NO:84)", value=True
)
std_mmo_122 = st.checkbox(
    "MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI (Y.NO:122)", value=True
)
std_mmo_133 = st.checkbox(
    "MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI (Y.NO:133)", value=True
)
std_deprem = st.checkbox(
    "Türkiye Bina Deprem Yönetmeliği (Mekanik Ekipman Askı ve Destekleri)",
    value=True,
)

ek_standartlar = st.text_area(
    "Eklemek istediğiniz ilave standartlar (Her satıra bir tane)", "", height=80
)

# --- 3. MEKANİK TESİSAT PROJE KAPSAMI ---
st.header("3. MEKANİK TESİSAT PROJE KAPSAMI")
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
kapsam_havalandirma = st.checkbox("Havalandırma Tesisatı", value=True)

ek_kapsam = st.text_area(
    "Eklemek istediğiniz ilave proje kapsam maddeleri", "", height=80
)

# --- 4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI ---
st.header("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI")
sicaklik_secenekleri = ["80/60", "70/50", "60/40", "50/30", "50/40", "7/12"]

col1, col2 = st.columns(2)
with col1:
  chk_kalorifer = st.checkbox("1. Kalorifer tesisatı", value=True)
  rej_kalorifer = st.selectbox(
      "Kalorifer Rejimi:", sicaklik_secenekleri, index=0
  )
  chk_fco_sog = st.checkbox("3. Fan-Coil Soğutma tesisatı", value=True)
  rej_fco_sog = st.selectbox(
      "Fan-Coil Soğutma Rejimi:", sicaklik_secenekleri, index=5
  )
with col2:
  chk_k_sicak = st.checkbox("7. Kullanma sıcak suyu", value=True)
  rej_k_sicak = st.selectbox(
      "Kullanma Sıcak Suyu Rejimi:", ["10/60", "50/40"], index=0
  )

# --- 5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ ---
st.header("5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ")
st.subheader("5.1 DIŞ HAVA TASARIM KRİTERLERİ")

iller_listesi = sorted(list(iklim_veritabani.keys()))
secilen_il = st.selectbox("İl seçin:", iller_listesi, index=0)
ilceler_listesi = sorted(list(iklim_veritabani[secilen_il].keys()))
secilen_ilce = st.selectbox("İlçe seçin:", ilceler_listesi)
iklim_veri = iklim_veritabani[secilen_il][secilen_ilce]

# --- 6. SIHHİ TESİSAT ---
st.header("6. SIHHİ TESİSAT")
st.subheader("6.1 SIHHİ TESİSAT ÖN BİLGİLER")
st.write(
    "Raporun 6.1 maddesinde yer almasını istediğiniz ön bilgi esaslarını"
    " seçin:"
)

sih_sec_1 = st.checkbox(
    "Bütün tesisin kullanma soğuk su ihtiyacı şehir şebekesinden sağlanacaktır.",
    value=True,
)
sih_sec_2 = st.checkbox(
    "Bütün tesisin kullanma soğuk su ihtiyacı kampüs içi su deposu dağıtım"
    " hattından sağlanacaktır.",
    value=False,
)
sih_sec_3 = st.checkbox(
    "Temiz su boru çapları yükleme birimine verilmiştir. 3/8” ’lik bir"
    " musluğun su verimi olan 0.25 lt/sn yükleme birimi olarak alınacaktır."
    " Diğer bütün sarfiyatlar bu birime tamamlanacaktır.",
    value=True,
)

sih_sec_4 = st.checkbox(
    "Bütün binanın kullanma soğuk su ihtiyacı soğuk su deposundan sağlanacaktır."
    " Basıncın yetersizliği ve su kesilmelerine karşın depo hidrofor sistemi"
    " uygulanmıştır. TS 1258 ve ilgili standartlar esas alınacaktır.",
    value=True,
)
sih_depo_konum = st.selectbox(
    "Soğuk Su Deposu Konumu:",
    ["Bodrum kat", "Zemin kat", "1. kat", "2. kat", "Çatı katı"],
    index=0,
)

sih_sec_depo_tipi = st.checkbox(
    "Bina da kullanım soğuk su depolaması için belirtilen tipte su deposu"
    " kullanılmıştır.",
    value=True,
)
sih_depo_tipi = st.selectbox(
    "Kullanma Soğuk Su Deposu Tipi:",
    [
        "Paslanmaz Çelik Modüler su deposu",
        "Galvaniz Çelik Modüler su deposu",
        "GRP (Cam Takviyeli Polyester) Modüler su deposu",
        "Betonarme Su deposu",
        "Silindirik Plastik Su deposu",
    ],
    index=0,
)

sih_sec_5 = st.checkbox(
    "Binada kullanılacak sıhhi tesisat elemanları birinci sınıf beyaz vitrifiye"
    " seramik olacaktır.",
    value=True,
)
sih_sec_6 = st.checkbox(
    "Tesisatta kullanılacak malzemeler ekstra sınıf olacak ve mimari projede"
    " belirtilen yerlere techiz edilecektir.",
    value=True,
)

sih_sec_7 = st.checkbox(
    "Kullanma Sıcak suyu üretimi ısı merkezindeki sistem vasıtasıyla"
    " yapılacaktır.",
    value=True,
)
sih_sicak_su_yontemi = st.selectbox(
    "Sıcak Su Üretim Sistemi / Yöntemi:",
    [
        "Dik tip hijyenik tek serpantinli boyler",
        "Dik tip hijyenik çift serpantinli boyler",
        "Elektrikli Sıcak Su üreticisi",
        "Kombi",
        "Plakalı eşanjör akümülasyon tankı",
        "Isıtma kazanı",
    ],
)

sih_sec_8 = st.checkbox(
    "Sıhhi tesisat işlerinde ana dağıtım boruları galvaniz çelik, mahal içi"
    " dağıtım boruları PPRC tipte seçilecektir.",
    value=True,
)
sih_sec_9 = st.checkbox(
    "Çamaşırhane, laboratuvar, mutfak mahallerinde yumuşak su kullanılacaktır.",
    value=True,
)
sih_sec_10 = st.checkbox(
    "Kullanım sıcak suyu hazırlanması için ısıtma kazanı ve güneş enerjisi"
    " sistemi kullanılacaktır.",
    value=True,
)
sih_sec_11 = st.checkbox(
    "Kullanım sıcak suyu hazırlanması için ısıtma kazanı kullanılacaktır.",
    value=False,
)
sih_sec_12 = st.checkbox(
    "Yağmur suyu toplama yönetmeliğine göre 2 bin metrekareden büyük"
    " parsellerde inşa edilecek tüm binaların çatılarında toplanan yağmur"
    " sularının, bahçe sulama veya arıtılarak bina ihtiyacında kullanılmak"
    " üzere bahçe zemini altında bir depoda toplaması amacıyla 'yağmur suyu"
    " toplama sistemi' yapılması zorunluluğu getirildiği için yağmur hasadı"
    " tesisatı yapılmıştır.",
    value=True,
)

ek_sihhi_on_bilgi = st.text_area(
    "İlave Sıhhi Tesisat Ön Bilgi Maddesi (Her satıra bir tane)", "", height=80
)

# Rapor Oluştur Butonu
if st.button("Raporu Oluştur (.docx)"):
  doc = Document()

  # Sayfa Yapısı
  cover_section = doc.sections[0]
  cover_section.top_margin = Inches(1.5)
  cover_section.bottom_margin = Inches(1.5)
  cover_section.left_margin = Inches(1.2)
  cover_section.right_margin = Inches(1.2)

  aktif_sirket = (
      sirket_adi
      if sirket_adi
      else "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ"
  )
  aktif_is = is_adi if is_adi else ""

  # 1. KAPAK
  p_s = doc.add_paragraph()
  p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
  r_s = p_s.add_run(aktif_sirket.upper())
  r_s.font.size = Pt(13)
  r_s.font.bold = True

  doc.add_paragraph()
  doc.add_paragraph()

  if aktif_is:
    p_i = doc.add_paragraph()
    p_i.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_i.add_run("PROJE ADI:\n").font.size = Pt(11)
    r_is = p_i.add_run(aktif_is)
    r_is.font.size = Pt(16)
    r_is.font.bold = True
    doc.add_paragraph()

  p_t = doc.add_paragraph()
  p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
  r_t = p_t.add_run(rapor_turu.upper())
  r_t.font.size = Pt(14)
  r_t.font.bold = True

  for _ in range(4):
    doc.add_paragraph()

  p_a = doc.add_paragraph()
  p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_a.add_run(
      f"Hazırlayan:\n{hazirlayan} (Makine Mühendisi)\nMMO Oda No:"
      f" {mmo_no}\n\nTarih:\n{tarih}"
  ).font.size = Pt(11)

  # 2. İÇİNDEKİLER
  doc.add_page_break()
  doc.add_heading("İÇİNDEKİLER", level=1)
  add_toc(doc.add_paragraph())

  # 3. GÖVDE
  doc.add_page_break()
  body_section = doc.add_section()
  body_section.top_margin = Inches(1.2)
  body_section.bottom_margin = Inches(1.2)
  body_section.left_margin = Inches(1.2)
  body_section.right_margin = Inches(1.2)

  # 1. GENEL BİLGİLER
  doc.add_heading("1. GENEL BİLGİLER", level=1)
  doc.add_paragraph(
      f"Bu raporda '{aktif_is}' için tasarlanan mekanik tesisatlar açıklanmıştır."
  )
  doc.add_paragraph(
      f"Yapı {secilen_il} ili {secilen_ilce} ilçesinde inşa edilecektir."
  )

  # 2. UYGULANACAK STANDART VE YÖNETMELİKLER
  doc.add_heading("2. UYGULANACAK STANDART VE YÖNETMELİKLER", level=1)
  st_list = []
  if std_ts_825:
    st_list.append("TS 825 - BİNALARDA ISI YALITIM KURALLARI")
  if std_yangin:
    st_list.append("BİNALARIN YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK")
  if std_bep_2008_2010:
    st_list.append("BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ")
  if std_ts_1258:
    st_list.append("TS 1258 – TEMİZSU TESİSATI HESAP KURALLARI")
  if std_ts_826:
    st_list.append("TS 826 – BİNALARDA PİSSU TESİSATI HESAPLAMA KURALLARI")
  if std_ts_2164:
    st_list.append("TS 2164 - KALORİFER TESİSATI PROJELENDİRME KURALLARI")
  if std_ts_3419:
    st_list.append(
        "TS 3419 – HAVALANDIRMA VE İKLİMLENDİRME TESİSLERİ PROJELENDİRME"
        " KURALLARI"
    )
  if std_ts_en_12056_2:
    st_list.append(
        "TS EN 12056-2 – CAZİBELİ DRENAJ SİSTEMLERİ -BİNA İÇİ- TASARIM VE"
        " HESAPLAMA"
    )
  if std_ts_en_12845:
    st_list.append("TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ")
  if std_mmo_84:
    st_list.append("MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI (Y.NO:84)")
  if std_mmo_122:
    st_list.append("MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI (Y.NO:122)")
  if std_mmo_133:
    st_list.append("MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI (Y.NO:133)")
  if std_deprem:
    st_list.append(
        "Türkiye Bina Deprem Yönetmeliği (Mekanik Ekipman Askı ve Destekleri)"
    )
  for s in st_list:
    doc.add_paragraph(s, style="List Bullet")

  # 3. MEKANİK TESİSAT PROJE KAPSAMI
  doc.add_heading("3. MEKANİK TESİSAT PROJE KAPSAMI", level=1)
  kp_list = []
  if kapsam_isitma:
    kp_list.append("Isıtma tesisatı,")
  if kapsam_sogutma:
    kp_list.append("Soğutma tesisatı,")
  if kapsam_soguk_su:
    kp_list.append("Kullanma soğuk suyu tesisatı,")
  if kapsam_sicak_su:
    kp_list.append("Kullanma sıcak suyu tesisatı,")
  if kapsam_yangin_depo:
    kp_list.append("Yangın ve kullanma suyu depolaması ve dağıtımı,")
  if kapsam_atik_su:
    kp_list.append("Yapı içinde atık su tesisatı (Yapı çıkış rögarına),")
  if kapsam_havalandirma:
    kp_list.append("Havalandırma Tesisatı")
  for k in kp_list:
    doc.add_paragraph(k, style="List Bullet")

  # 4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI
  doc.add_heading("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI", level=1)
  if chk_kalorifer:
    doc.add_paragraph(
        f"Kalorifer tesisatında {rej_kalorifer} °C sıcak su.", style="List Bullet"
    )
  if chk_fco_sog:
    doc.add_paragraph(
        f"Fan-Coil Soğutma tesisatında {rej_fco_sog} °C soğuk su.",
        style="List Bullet",
    )
  if chk_k_sicak:
    doc.add_paragraph(
        f"Kullanma sıcak suyunda {rej_k_sicak} °C sıcak su.", style="List Bullet"
    )

  # 5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ
  doc.add_heading("5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ", level=1)
  doc.add_heading("5.1 DIŞ HAVA TASARIM KRİTERLERİ", level=2)
  doc.add_paragraph(
      f"Yapının inşa edileceği ''{secilen_il}'' için kabul edilen dış hava"
      " koşulları aşağıdaki gibidir:"
  )
  doc.add_paragraph(
      f"• KIŞ: {iklim_veri['kis_kt']} °C KT , {iklim_veri['kis_yt']} °C YT",
      style="List Bullet",
  )
  doc.add_paragraph(
      f"• YAZ: {iklim_veri['yaz_kt']} °C KT , {iklim_veri['yaz_yt']} °C YT",
      style="List Bullet",
  )
  doc.add_paragraph(f"• Enlem: {iklim_veri['enlem']}", style="List Bullet")
  doc.add_paragraph(f"• Boylam: {iklim_veri['boylam']}", style="List Bullet")
  doc.add_paragraph(
      f"• Rakım: {iklim_veri['rakim']} m.", style="List Bullet"
  )
  doc.add_paragraph(
      f"• Günlük Sıcaklık Farkı (GSF): {iklim_veri['gsf']} °C",
      style="List Bullet",
  )

  # 6. SIHHİ TESİSAT & 6.1 ÖN BİLGİLER
  doc.add_heading("6. SIHHİ TESİSAT", level=1)
  doc.add_heading("6.1 SIHHİ TESİSAT ÖN BİLGİLER", level=2)

  sihhi_maddeler = []
  if sih_sec_1:
    sihhi_maddeler.append(
        "Bütün tesisin kullanma soğuk su ihtiyacı şehir şebekesinden"
        " sağlanacaktır."
    )
  if sih_sec_2:
    sihhi_maddeler.append(
        "Bütün tesisin kullanma soğuk su ihtiyacı kampüs içi su deposu dağıtım"
        " hattından sağlanacaktır."
    )
  if sih_sec_3:
    sihhi_maddeler.append(
        "Temiz su boru çapları yükleme birimine verilmiştir. 3/8” ’lik bir"
        " musluğun su verimi olan 0.25 lt/sn yükleme birimi olarak alınacaktır."
        " Diğer bütün sarfiyatlar bu birime tamamlanacaktır."
    )
  if sih_sec_4:
    sihhi_maddeler.append(
        f"Bütün binanın kullanma soğuk su ihtiyacı {sih_depo_konum.lower()} soğuk"
        " su deposundan sağlanacaktır. Basıncın yetersizliği ve su"
        " kesilmelerine karşın depo hidrofor sistemi uygulanmıştır. TS 1258 ve"
        " ilgili standartlar esas alınacaktır."
    )
  if sih_sec_depo_tipi:
    sihhi_maddeler.append(
        "Binada kullanım soğuk su depolaması için "
        f"{sih_depo_tipi.lower()} tipinde su deposu kullanılmıştır."
    )
  if sih_sec_5:
    sihhi_maddeler.append(
        "Binada kullanılacak sıhhi tesisat elemanları birinci sınıf beyaz"
        " vitrifiye seramik olacaktır."
    )
  if sih_sec_6:
    sihhi_maddeler.append(
        "Tesisatta kullanılacak malzemeler ekstra sınıf olacak ve mimari"
        " projede belirtilen yerlere techiz edilecektir."
    )
  if sih_sec_7:
    sihhi_maddeler.append(
        "Kullanma Sıcak suyu üretimi ısı merkezindeki "
        f"{sih_sicak_su_yontemi.lower()} vasıtasıyla yapılacaktır."
    )
  if sih_sec_8:
    sihhi_maddeler.append(
        "Sıhhi tesisat işlerinde ana dağıtım boruları galvaniz çelik, mahal içi"
        " dağıtım boruları PPRC tipte seçilecektir."
    )
  if sih_sec_9:
    sihhi_maddeler.append(
        "Çamaşırhane, laboratuvar, mutfak mahallerinde yumuşak su"
        " kullanılacaktır."
    )
  if sih_sec_10:
    sihhi_maddeler.append(
        "Kullanım sıcak suyu hazırlanması için ısıtma kazanı ve güneş enerjisi"
        " sistemi kullanılacaktır."
    )
  if sih_sec_11:
    sihhi_maddeler.append(
        "Kullanım sıcak suyu hazırlanması için ısıtma kazanı kullanılacaktır."
    )
  if sih_sec_12:
    sihhi_maddeler.append(
        "Yağmur suyu toplama yönetmeliğine göre 2 bin metrekareden büyük"
        " parsellerde inşa edilecek tüm binaların çatılarında toplanan yağmur"
        " sularının, bahçe sulama veya arıtılarak bina ihtiyacında kullanılmak"
        " üzere bahçe zemini altında bir depoda toplaması amacıyla 'yağmur suyu"
        " toplama sistemi' yapılması zorunluluğu getirildiği için yağmur hasadı"
        " tesisatı yapılmıştır."
    )

  if ek_sihhi_on_bilgi.strip():
    for es in ek_sihhi_on_bilgi.split("\n"):
      if es.strip():
        sihhi_maddeler.append(es.strip())

  for sm in sihhi_maddeler:
    doc.add_paragraph(sm, style="List Bullet")

  # Kaydetme ve İndirme
  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)

  st.success("Tüm bölümler eksiksiz olarak güncellendi ve rapor hazırlandı!")
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
