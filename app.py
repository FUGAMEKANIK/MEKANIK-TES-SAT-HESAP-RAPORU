from datetime import datetime
import io
import json
import math
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import matplotlib.pyplot as plt
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


def _toplu_checkbox_ayarla(anahtarlar, durum):
  for anahtar in anahtarlar:
    st.session_state[anahtar] = durum


def _toplu_secim_butonlari(
    anahtarlar, kolon_basliklari=("☑ TÜMÜNÜ SEÇ", "☐ TÜMÜNÜ KALDIR")
):
  c1, c2 = st.columns(2)
  with c1:
    st.button(
        kolon_basliklari[0],
        key=f"toplu_sec_{anahtarlar[0]}",
        on_click=_toplu_checkbox_ayarla,
        args=(anahtarlar, True),
        use_container_width=True,
    )
  with c2:
    st.button(
        kolon_basliklari[1],
        key=f"toplu_kaldir_{anahtarlar[0]}",
        on_click=_toplu_checkbox_ayarla,
        args=(anahtarlar, False),
        use_container_width=True,
    )


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

std_keys = [
    "std_ts_825",
    "std_yangin",
    "std_bep_2008_2010",
    "std_ts_1258",
    "std_ts_826",
    "std_ts_2164",
    "std_ts_3419",
    "std_ts_en_12056_2",
    "std_ts_en_12845",
    "std_mmo_84",
    "std_mmo_352_5",
    "std_mmo_122",
    "std_mmo_133",
    "std_mmo_155",
    "std_ashrae",
    "std_su",
    "std_klima",
    "std_tesisat",
    "std_kanal",
    "std_asansor",
    "std_deprem",
    "std_akustik",
    "std_isg",
]
_toplu_secim_butonlari(std_keys)

std_ts_825 = st.checkbox(
    "TS 825 - BİNALARDA ISI YALITIM KURALLARI", key="std_ts_825", value=True
)
std_yangin = st.checkbox(
    '09 Eylül 2009 tarih ve 27344 numaralı sayısında yayımlanan " BİNALARIN'
    ' YANGINDAN KORUNMASI HAKKINDA YÖNETMELİK"',
    key="std_yangin",
    value=True,
)
std_bep_2008_2010 = st.checkbox(
    "5 Aralık 2008 tarih, 27075 sayılı resmi gazetede yayımlanan “BİNALARDA"
    " ENERJİ PERFORMANSI YÖNETMELİĞİ” ve 1 Nisan 2010 tarih, 27539 sayılı resmi"
    " gazetede yayımlanan “BİNALARDA ENERJİ PERFORMANSI YÖNETMELİĞİ”",
    key="std_bep_2008_2010",
    value=True,
)
std_ts_1258 = st.checkbox(
    "TS 1258 – TEMİZSU TESİSATI HESAP KURALLARI",
    key="std_ts_1258",
    value=True,
)
std_ts_826 = st.checkbox(
    "TS 826 – BİNALARDA PİSSU TESİSATI HESAPLAMA KURALLARI",
    key="std_ts_826",
    value=True,
)
std_ts_2164 = st.checkbox(
    "TS 2164 - KALORİFER TESİSATI PROJELENDİRME KURALLARI",
    key="std_ts_2164",
    value=True,
)
std_ts_3419 = st.checkbox(
    "TS 3419 – HAVALANDIRMA VE İKLİMLENDİRME TESİSLERİ PROJELENDİRME"
    " KURALLARI",
    key="std_ts_3419",
    value=True,
)
std_ts_en_12056_2 = st.checkbox(
    "TS EN 12056-2 – CAZİBELİ DRENAJ SİSTEMLERİ -BİNA İÇİ- TASARIM VE HESAPLAMA",
    key="std_ts_en_12056_2",
    value=True,
)
std_ts_en_12845 = st.checkbox(
    "TS EN 12845 – SABİT YANGIN SÖNDÜRME SİSTEMLERİ – OTOMATİK SPRİNKLER"
    " SİSTEMLERİ- TASARIM, MONTAJ VE BAKIM",
    key="std_ts_en_12845",
    value=True,
)
std_mmo_84 = st.checkbox(
    "MMO KALORİFER TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:84)",
    key="std_mmo_84",
    value=True,
)
std_mmo_352_5 = st.checkbox(
    "MMO KALORİFER TESİSATI (Y.NO:352/5)", key="std_mmo_352_5", value=True
)
std_mmo_122 = st.checkbox(
    "MMO SIHHİ TESİSAT PROJE HAZIRLAMA ESASLARI(Y.NO:122)",
    key="std_mmo_122",
    value=True,
)
std_mmo_133 = st.checkbox(
    "MMO GAZ TESİSATI PROJE HAZIRLAMA ESASLARI(Y.NO:133)",
    key="std_mmo_133",
    value=True,
)
std_mmo_155 = st.checkbox(
    "MMO KAZAN VE BACA(Y.NO:155)", key="std_mmo_155", value=True
)

std_ashrae = st.checkbox("ASHRAE Standartları", key="std_ashrae", value=True)
std_su = st.checkbox(
    "İçmesuyu Temizleme ve Dağıtım Sistemleri Standartları",
    key="std_su",
    value=True,
)
std_klima = st.checkbox(
    "Klima ve Havalandırma Tesisatı Yönetmelikleri",
    key="std_klima",
    value=True,
)
std_tesisat = st.checkbox(
    "Merkezi Isıtma ve Sıhhi Sıcak Su Sistemlerinde Isı Maliyetlerinin"
    " Paylaştırılmasına İlişkin Yönetmelik",
    key="std_tesisat",
    value=False,
)
std_kanal = st.checkbox(
    "Kanalizasyon Şebekesi Olmayan Yerlerde Yapılacak Çukurlar",
    key="std_kanal",
    value=False,
)
std_asansor = st.checkbox(
    "Asansör Yönetmeliği ve İlgili Standartlar", key="std_asansor", value=False
)
std_deprem = st.checkbox(
    "Türkiye Bina Deprem Yönetmeliği (Mekanik Ekipman Askı ve Destekleri)",
    key="std_deprem",
    value=True,
)
std_akustik = st.checkbox(
    "Binaların Gürültüye Karşı Korunması Yönetmeliği",
    key="std_akustik",
    value=False,
)
std_isg = st.checkbox(
    "İş Sağlığı ve Güvenliği Kanunu ve İlgili Yönetmelikler",
    key="std_isg",
    value=True,
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

kapsam_keys = [
    "kapsam_isitma",
    "kapsam_sogutma",
    "kapsam_soguk_su",
    "kapsam_sicak_su",
    "kapsam_yangin_depo",
    "kapsam_atik_su",
    "kapsam_yangin_dagitim",
    "kapsam_kazan_dairesi",
    "kapsam_havalandirma",
    "kapsam_basinc_hava",
    "kapsam_medikal_gaz",
    "kapsam_otomatik",
]
_toplu_secim_butonlari(kapsam_keys)

kapsam_isitma = st.checkbox(
    "Isıtma tesisatı,", key="kapsam_isitma", value=True
)
kapsam_sogutma = st.checkbox(
    "Soğutma tesisatı,", key="kapsam_sogutma", value=True
)
kapsam_soguk_su = st.checkbox(
    "Kullanma soğuk suyu tesisatı,", key="kapsam_soguk_su", value=True
)
kapsam_sicak_su = st.checkbox(
    "Kullanma sıcak suyu tesisatı,", key="kapsam_sicak_su", value=True
)
kapsam_yangin_depo = st.checkbox(
    "Yangın ve kullanma suyu depolaması ve dağıtımı,",
    key="kapsam_yangin_depo",
    value=True,
)
kapsam_atik_su = st.checkbox(
    "Yapı içinde atık su tesisatı (Yapı çıkış rögarına),",
    key="kapsam_atik_su",
    value=True,
)
kapsam_yangin_dagitim = st.checkbox(
    "Yangın suyu iç ve dış dağıtım sistemleri,",
    key="kapsam_yangin_dagitim",
    value=True,
)
kapsam_kazan_dairesi = st.checkbox(
    "Merkezi ısıtma kazan dairesi ve tali teknik hacimler,",
    key="kapsam_kazan_dairesi",
    value=True,
)
kapsam_havalandirma = st.checkbox(
    "Havalandırma Tesisatı", key="kapsam_havalandirma", value=True
)
kapsam_basinc_hava = st.checkbox(
    "Basınçlı hava tesisatı,", key="kapsam_basinc_hava", value=False
)
kapsam_medikal_gaz = st.checkbox(
    "Medikal gaz tesisatı", key="kapsam_medikal_gaz", value=False
)
kapsam_otomatik = st.checkbox(
    "Otomatik kontrol sistemi kavramı tanımı,",
    key="kapsam_otomatik",
    value=True,
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

akiskan_keys = [
    "chk_kalorifer",
    "chk_fco_ist",
    "chk_fco_sog",
    "chk_ks_ist",
    "chk_buhar",
    "chk_ks_sog",
    "chk_boyler",
    "chk_k_sicak",
    "chk_doseme",
    "chk_kizgin",
]
_toplu_secim_butonlari(akiskan_keys)

col1, col2 = st.columns(2)

with col1:
  chk_kalorifer = st.checkbox(
      "1. Kalorifer tesisatı", key="chk_kalorifer", value=True
  )
  rej_kalorifer = st.selectbox(
      "Kalorifer Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_fco_ist = st.checkbox(
      "2. Fan-Coil ısıtma tesisatı", key="chk_fco_ist", value=True
  )
  rej_fco_ist = st.selectbox(
      "Fan-Coil Isıtma Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_fco_sog = st.checkbox(
      "3. Fan-Coil Soğutma tesisatı", key="chk_fco_sog", value=True
  )
  rej_fco_sog = st.selectbox(
      "Fan-Coil Soğutma Rejimi:", sicaklik_secenekleri, index=5
  )

  chk_ks_ist = st.checkbox(
      "4. Klima santrali ısıtma tesisatı", key="chk_ks_ist", value=True
  )
  rej_ks_ist = st.selectbox(
      "Klima Santrali Isıtma Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_buhar = st.checkbox("9. Buhar tesisatı", key="chk_buhar", value=False)
  rej_buhar = st.selectbox("Buhar Seçimi:", buhar_secenekleri, index=1)

with col2:
  chk_ks_sog = st.checkbox(
      "5. Klima santrali Soğutma tesisatı", key="chk_ks_sog", value=True
  )
  rej_ks_sog = st.selectbox(
      "Klima Santrali Soğutma Rejimi:", sicaklik_secenekleri, index=5
  )

  chk_boyler = st.checkbox(
      "6. Boyler ısıtma tesisatı", key="chk_boyler", value=True
  )
  rej_boyler = st.selectbox(
      "Boyler Isıtma Rejimi:", sicaklik_secenekleri, index=0
  )

  chk_k_sicak = st.checkbox(
      "7. Kullanma sıcak suyu", key="chk_k_sicak", value=True
  )
  rej_k_sicak = st.selectbox(
      "Kullanma Sıcak Suyu Rejimi:", sicaklik_secenekleri, index=7
  )

  chk_doseme = st.checkbox(
      "8. Döşemeden ısıtma tesisatı", key="chk_doseme", value=True
  )
  rej_doseme = st.selectbox(
      "Döşemeden Isıtma Rejimi:", sicaklik_secenekleri, index=4
  )

  chk_kizgin = st.checkbox(
      "10. Kızgın su tesisatı", key="chk_kizgin", value=False
  )
  rej_kizgin = st.selectbox(
      "Kızgın Su Rejimi:", kizgin_su_secenekleri, index=0
  )

# --- 5. BÖLÜM: İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ ---
st.header("5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ")
st.subheader("5.1 DIŞ HAVA TASARIM KRİTERLERİ")

iller_listesi = sorted(list(iklim_veritabani.keys()))
secilen_il = st.selectbox(
    "Yapının inşa edileceği ili seçin:", iller_listesi, index=0
)

ilceler_listesi = sorted(list(iklim_veritabani[secilen_il].keys()))
secilen_ilce = st.selectbox(
    "Yapının inşa edileceği ilçeyi seçin:", ilceler_listesi
)

iklim_veri = iklim_veritabani[secilen_il][secilen_ilce]

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
  st.markdown(f"• **Deniz seviyesinden yüksekliği**: `{iklim_veri['rakim']}` m.")

# --- 6. BÖLÜM: SIHHİ TESİSAT ---
st.header("6. SIHHİ TESİSAT")
st.subheader("6.1 SIHHİ TESİSAT ÖN BİLGİLER")
st.write(
    "Raporun 6.1 maddesinde yer almasını istediğiniz ön bilgi esaslarını"
    " seçin:"
)

sihhi_keys = [
    "sih_sec_1",
    "sih_sec_2",
    "sih_sec_3",
    "sih_sec_4",
    "sih_sec_depo_tipi",
    "sih_sec_5",
    "sih_sec_6",
    "sih_sec_7",
    "sih_sec_8",
    "sih_sec_9",
    "sih_sec_10",
    "sih_sec_12",
]
_toplu_secim_butonlari(sihhi_keys)

sih_sec_1 = st.checkbox(
    "Bütün tesisin kullanma soğuk su ihtiyacı şehir şebekesinden sağlanacaktır.",
    key="sih_sec_1",
    value=True,
)
sih_sec_2 = st.checkbox(
    "Bütün tesisin kullanma soğuk su ihtiyacı kampüs içi su deposu dağıtım"
    " hattından sağlanacaktır.",
    key="sih_sec_2",
    value=False,
)
sih_sec_3 = st.checkbox(
    "Temiz su boru çapları yükleme birimine verilmiştir. 3/8” ’lik bir"
    " musluğun su verimi olan 0.25 lt/sn yükleme birimi olarak alınacaktır."
    " Diğer bütün sarfiyatlar bu birime tamamlanacaktır.",
    key="sih_sec_3",
    value=True,
)
sih_sec_4 = st.checkbox(
    "Bütün binanın kullanma soğuk su ihtiyacı soğuk su deposundan sağlanacaktır."
    " Basıncın yetersizliği ve su kesilmelerine karşın depo hidrofor sistemi"
    " uygulanmıştır. TS 1258 ve ilgili standartlar esas alınacaktır.",
    key="sih_sec_4",
    value=True,
)
sih_depo_konumlari = st.multiselect(
    "Soğuk Su Deposu Konumu (Birden fazla seçebilirsiniz):",
    ["Bodrum kat", "Zemin kat", "1. kat", "2. kat", "Çatı katı"],
    default=["Bodrum kat"],
)

sih_sec_depo_tipi = st.checkbox(
    "Bina da kullanım soğuk su depolaması için belirtilen tipte su deposu"
    " kullanılmıştır.",
    key="sih_sec_depo_tipi",
    value=True,
)
sih_depo_tipleri = st.multiselect(
    "Kullanma Soğuk Su Deposu Tipi (Birden fazla seçebilirsiniz):",
    [
        "Paslanmaz Çelik Modüler su deposu",
        "Galvaniz Çelik Modüler su deposu",
        "GRP (Cam Takviyeli Polyester) Modüler su deposu",
        "Betonarme Su deposu",
        "Silindirik Plastik Su deposu",
    ],
    default=["Paslanmaz Çelik Modüler su deposu"],
)

sih_sec_5 = st.checkbox(
    "Binada kullanılacak sıhhi tesisat elemanları birinci sınıf beyaz vitrifiye"
    " seramik olacaktır.",
    key="sih_sec_5",
    value=True,
)
sih_sec_6 = st.checkbox(
    "Tesisatta kullanılacak malzemeler ekstra sınıf olacak ve mimari projede"
    " belirtilen yerlere techiz edilecektir.",
    key="sih_sec_6",
    value=True,
)
sih_sec_7 = st.checkbox(
    "Kullanma Sıcak suyu üretimi ısı merkezindeki sistem vasıtasıyla"
    " yapılacaktır.",
    key="sih_sec_7",
    value=True,
)
sih_sicak_su_yontemleri = st.multiselect(
    "Sıcak Su Üretim Sistemi / Yöntemleri (Birden fazla seçebilirsiniz):",
    [
        "Dik tip hijyenik tek serpantinli boyler",
        "Dik tip hijyenik çift serpantinli boyler",
        "Elektrikli Sıcak Su üreticisi",
        "Kombi",
        "Plakalı eşanjör akümülasyon tankı",
        "Isıtma kazanı",
    ],
    default=["Dik tip hijyenik tek serpantinli boyler"],
)
sih_sec_8 = st.checkbox(
    "Sıhhi tesisat işlerinde ana dağıtım boruları galvaniz çelik, mahal içi"
    " dağıtım boruları PPRC tipte seçilecektir.",
    key="sih_sec_8",
    value=True,
)

sih_sec_9 = st.checkbox(
    "Belirtilen mahallerde yumuşak su kullanılacaktır.",
    key="sih_sec_9",
    value=True,
)
sih_yumusak_su_mahalleri = st.multiselect(
    "Yumuşak Su Kullanılacak Mahaller (Birden fazla seçebilirsiniz):",
    ["Çamaşırhane", "Laboratuvar", "Mutfak", "Kuaför / Spa", "Diğer"],
    default=["Çamaşırhane", "Laboratuvar", "Mutfak"],
)
sih_yumusak_su_diger = st.text_input(
    "Diğer (Yumuşak su kullanılacak başka mahal varsa yazın):", ""
)

sih_sec_10 = st.checkbox(
    "Kullanım sıcak suyunun ısıtılması belirtilen sistemler vasıtasıyla"
    " yapılacaktır.",
    key="sih_sec_10",
    value=True,
)
sih_sicak_su_isitma_sistemleri = st.multiselect(
    "Kullanım Sıcak Suyu Isıtma Sistemleri (Birden fazla seçebilirsiniz):",
    ["Kazan", "Güneş enerjisi", "Elektrik"],
    default=["Kazan", "Güneş enerjisi"],
)

sih_sec_12 = st.checkbox(
    "Yağmur suyu toplama yönetmeliğine göre 2 bin metrekareden büyük"
    " parsellerde inşa edilecek tüm binaların çatılarında toplanan yağmur"
    " sularının, bahçe sulama veya arıtılarak bina ihtiyacında kullanılmak"
    " üzere bahçe zemini altında bir depoda toplaması amacıyla 'yağmur suyu"
    " toplama sistemi' yapılması zorunluluğu getirildiği için yağmur hasadı"
    " tesisatı yapılmıştır.",
    key="sih_sec_12",
    value=True,
)

ek_sihhi_on_bilgi = st.text_area(
    "İlave Sıhhi Tesisat Ön Bilgi Maddesi (Her satıra bir tane)", "", height=80
)

# --- 6.1.1 TEMİZ SU SARFİYAT YÜKLEME BİRİMLERİ VE ÇAP TAYİNİ ---
st.subheader("6.1.1 Temiz Su Sarfiyat Yükleme Birimleri ve Çap Tayini Girdileri")


# --- 6.2 PİS SU TESİSATI ---
st.subheader("6.2 PİS SU TESİSATI ESASLARI")
pissu_keys = [
    "pissu_sec_1",
    "pissu_sec_2",
    "pissu_sec_3",
    "pissu_sec_4",
    "pissu_sec_5",
    "pissu_sec_6",
    "pissu_sec_7",
]
_toplu_secim_butonlari(pissu_keys)

pissu_sec_1 = st.checkbox(
    "Yapının atık suları binanın pik kolonlarla toplanarak belirtilen kat ve"
    " geçiş yerlerinden rögarlara iletilecektir.",
    key="pissu_sec_1",
    value=True,
)
pis_su_konumlari = st.multiselect(
    "Atık su toplama konumu:",
    ["Bodrum kat", "Zemin kat", "1. kat", "Çatı katı"],
    default=["Bodrum kat"],
)
pis_su_gecisler = st.multiselect(
    "Atık su boru geçiş yeri:",
    ["döşemesinden", "tavanından", "asma tavan arasından"],
    default=["döşemesinden"],
)

pissu_sec_2 = st.checkbox(
    "Pis su kolonları üzerinde gerekli yerlere temizleme kapakları"
    " yerleştirilmiştir.",
    key="pissu_sec_2",
    value=True,
)
pissu_sec_3 = st.checkbox(
    "Tüm teknik hacimlerde, su tahliyesi için ızgaralı kanallar yapılacaktır.",
    key="pissu_sec_3",
    value=True,
)
pissu_sec_4 = st.checkbox(
    "Atık su boruları sessiz PVC boru gibi son teknoloji ürünü borular"
    " kullanılacaktır.",
    key="pissu_sec_4",
    value=True,
)
pissu_sec_5 = st.checkbox(
    "Pis su boru çapları yükleme birimi yöntemine göre belirlenmiştir.",
    key="pissu_sec_5",
    value=True,
)
pissu_sec_6 = st.checkbox(
    "Pis su akar kotunun kurtarmayan katları bodrum katta pis su çukurunda"
    " toplanıp, pompa vasıtasıyla yol kotundaki rögara aktarılacaktır.",
    key="pissu_sec_6",
    value=True,
)
pissu_sec_7 = st.checkbox(
    "Pis su vaziyette de görüldüğü gibi rögarlar vasıtasıyla yoldan geçen"
    " pis su kanalına bağlanacaktır.",
    key="pissu_sec_7",
    value=True,
)

ek_pissu_on_bilgi = st.text_area(
    "İlave Pis Su Tesisatı Maddesi (Her satıra bir tane)", "", height=80
)


# ---------------------------------------------------------------------------
# 6.1.1 & 6.2.1 TABLOLAR
# ---------------------------------------------------------------------------
st.subheader("6.1.1 Temiz Su Sarfiyat Yükleme Birimleri ve Çap Tayini")
t1_data = [
    ("DN", "PLASTİK", "ÇELİK", "Yükleme Birimi"),
    ("15", "Ø20", '1/2"', "(0-3.0)"),
    ("20", "Ø25", '3/4"', "(3.0-8.0)"),
    ("25", "Ø32", '1"', "(8.0-20.0)"),
    ("32", "Ø40", '1 1/4"', "(20.0-35.0)"),
    ("40", "Ø50", '1 1/2"', "(35.0-50.0)"),
    ("50", "Ø63", '2"', "(50.0-144.0)"),
    ("65", "Ø75", '2 1/2"', "(144.0-368.0)"),
    ("80", "Ø90", '3"', "(368.0-1156.0)"),
    ("100", "Ø125", '4"', "(1156-4900)"),
    ("125", "-", '5"', "(4.900-14.400)"),
    ("150", "-", '6"', "(14.400-40.000)"),
    ("200", "-", '8"', "(40.000-484.000)"),
    ("250", "-", '10"', "(484.000-518.400)"),
    ("300", "-", '12"', "(518.400-1.440.000)"),
]

st.subheader("6.2.1 Pis Su Sarfiyat Yükleme Birimleri ve Çap Tayini")
pissu_t1_data = [
    ("KULLANMA YERİ", "YÜKLEME BİRİMİ"),
    ("Hela / Klozet", "8"),
    ("Küvet / Duş", "7"),
    ("Lavabo / Bide", "2"),
    ("Eviye", "4"),
    ("Yer Süzgeci", "2"),
    ("Otopark Süzgeci", "6"),
    ("Basınçlı Yıkayıcı", "10"),
    ("Çamaşır-Bulaşık Makinası", "10"),
    ("Pisuar", "1"),
]
pissu_t2_data = [
    ("YÜKLEME BİRİMİ", "% 1 EĞİM", "BORU ÇAPI"),
    ("0-7", "", "50"),
    ("7-25", "", "70"),
    ("25-120", "", "100"),
    ("120-270", "", "125"),
    ("270-600", "", "150"),
    ("600-2400", "", "200"),
]


# ---------------------------------------------------------------------------
# 6.2.2 PİS SU TERFİ POMPALARI SEÇİMİ VE ÖZEL HESAP MODÜLÜ
# ---------------------------------------------------------------------------
st.subheader(
    "6.2.2 Her Bir Terfi Pompası İçin Özel Debi ve Güç Hesap Modülü"
)

POZ_POMPA_TABLOSU = [
    {
        "poz": "25.360.1301",
        "qmin": 5.0,
        "qmax": 10.0,
        "hmin": 5.0,
        "hmax": 10.0,
        "tanim": (
            "Debisi 5,0-10 m³/h, basıncı 5,0-10 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
    {
        "poz": "25.360.1302",
        "qmin": 5.0,
        "qmax": 10.0,
        "hmin": 10.0,
        "hmax": 15.0,
        "tanim": (
            "Debisi 5,0-10 m³/h, basıncı 10-15 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
    {
        "poz": "25.360.1303",
        "qmin": 5.0,
        "qmax": 10.0,
        "hmin": 15.0,
        "hmax": 20.0,
        "tanim": (
            "Debisi 5,0-10 m³/h, basıncı 15-20 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
    {
        "poz": "25.360.1304",
        "qmin": 10.0,
        "qmax": 15.0,
        "hmin": 5.0,
        "hmax": 10.0,
        "tanim": (
            "Debisi 10-15 m³/h, basıncı 5,0-10 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
    {
        "poz": "25.360.1305",
        "qmin": 10.0,
        "qmax": 15.0,
        "hmin": 10.0,
        "hmax": 15.0,
        "tanim": (
            "Debisi 10-15 m³/h, basıncı 10-15 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
    {
        "poz": "25.360.1306",
        "qmin": 15.0,
        "qmax": 20.0,
        "hmin": 5.0,
        "hmax": 10.0,
        "tanim": (
            "Debisi 15-20 m³/h, basıncı 5,0-10 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
    {
        "poz": "25.360.1307",
        "qmin": 15.0,
        "qmax": 20.0,
        "hmin": 10.0,
        "hmax": 15.0,
        "tanim": (
            "Debisi 15-20 m³/h, basıncı 10-15 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
    {
        "poz": "25.360.1308",
        "qmin": 15.0,
        "qmax": 20.0,
        "hmin": 15.0,
        "hmax": 20.0,
        "tanim": (
            "Debisi 15-20 m³/h, basıncı 15-20 mSS parçalayıcı bıçaklı dalgıç"
            " tip pis su pompası"
        ),
    },
]


def pompa_pozu_sec(q_m3h, h_mss):
  for kayit in POZ_POMPA_TABLOSU:
    if (kayit["qmin"] <= q_m3h <= kayit["qmax"]) and (
        kayit["hmin"] <= h_mss <= kayit["hmax"]
    ):
      return kayit["poz"], kayit["tanim"], "UYGUN"
  return (
      "SINIR DIŞI",
      (
          "Girilen tek pompa çalışma noktası 25.360.1301–1308 poz kapasite"
          " sınırları dışındadır!"
      ),
      "GECERSIZ",
  )


def pompa_hidrolik_hesap(q_m3h, h_mss, pompa_verimi=0.60, motor_verimi=0.90):
  rho = 1000.0
  g = 9.81
  q_m3s = q_m3h / 3600.0
  p_hid_kw = rho * g * q_m3s * h_mss / 1000.0
  p_mil_kw = p_hid_kw / pompa_verimi if pompa_verimi > 0 else 0.0
  p_elektrik_kw = p_mil_kw / motor_verimi if motor_verimi > 0 else 0.0

  motor_kademeleri = [
      0.75,
      1.1,
      1.5,
      2.2,
      3.0,
      4.0,
      5.5,
      7.5,
      11.0,
      15.0,
      18.5,
      22.0,
      30.0,
  ]
  motor_secim = next(
      (x for x in motor_kademeleri if x >= p_elektrik_kw), motor_kademeleri[-1]
  )

  return {
      "q_m3h": q_m3h,
      "h_mss": h_mss,
      "q_lps": q_m3h / 3.6,
      "p_hid_kw": p_hid_kw,
      "p_mil_kw": p_mil_kw,
      "p_elektrik_kw": p_elektrik_kw,
      "motor_secim_kw": motor_secim,
  }


URETICI_POMPA_VERITABANI = [
    {
        "marka": "Grundfos",
        "seri": "SEG",
        "model": "SEG.40.15.1",
        "q_min": 0.0, "q_max": 5.2, "h_max": 26.0,
        "p2_kw": 1.5,
        "curve": [(0.0,26.0),(1.0,24.0),(2.0,21.0),(3.0,17.0),(4.0,13.0),(5.0,9.0),(5.2,8.0)],
        "kaynak": "Grundfos SEG 50 Hz performans eğrisi / ISO 9906"
    },
    {
        "marka": "Grundfos",
        "seri": "SEG",
        "model": "SEG.40.26.3",
        "q_min": 0.0, "q_max": 5.2, "h_max": 40.0,
        "p2_kw": 2.6,
        "curve": [(0.0,40.0),(1.0,36.0),(2.0,31.0),(3.0,25.0),(4.0,18.0),(5.0,11.0),(5.2,9.0)],
        "kaynak": "Grundfos SEG 50 Hz performans eğrisi / ISO 9906"
    },
    {
        "marka": "Grundfos",
        "seri": "SEG",
        "model": "SEG.40.31.3",
        "q_min": 0.0, "q_max": 5.2, "h_max": 45.0,
        "p2_kw": 3.1,
        "curve": [(0.0,45.0),(1.0,41.0),(2.0,35.0),(3.0,29.0),(4.0,21.0),(5.0,13.0),(5.2,11.0)],
        "kaynak": "Grundfos SEG 50 Hz performans eğrisi / ISO 9906"
    },
    {
        "marka": "Wilo",
        "seri": "Rexa CUT",
        "model": "Rexa CUT GI03.20",
        "q_min": 0.0, "q_max": 20.0, "h_max": 20.0,
        "p2_kw": 1.1,
        "curve": [(0.0,20.0),(4.0,18.0),(8.0,15.0),(12.0,11.5),(16.0,7.5),(20.0,3.0)],
        "kaynak": "Wilo Rexa CUT katalog performans grafiği"
    },
    {
        "marka": "Wilo",
        "seri": "Rexa CUT",
        "model": "Rexa CUT GI03.25",
        "q_min": 0.0, "q_max": 20.0, "h_max": 25.0,
        "p2_kw": 2.5,
        "curve": [(0.0,25.0),(4.0,23.0),(8.0,20.0),(12.0,16.0),(16.0,11.0),(20.0,6.0)],
        "kaynak": "Wilo Rexa CUT katalog performans grafiği"
    },
    {
        "marka": "Wilo",
        "seri": "Rexa CUT",
        "model": "Rexa CUT GI03.29",
        "q_min": 0.0, "q_max": 20.0, "h_max": 29.0,
        "p2_kw": 1.5,
        "curve": [(0.0,29.0),(4.0,27.0),(8.0,24.0),(12.0,20.0),(16.0,14.0),(20.0,8.0)],
        "kaynak": "Wilo Rexa CUT katalog performans grafiği"
    },
    {
        "marka": "Wilo",
        "seri": "Rexa CUT",
        "model": "Rexa CUT GI03.34",
        "q_min": 0.0, "q_max": 20.0, "h_max": 34.0,
        "p2_kw": 2.5,
        "curve": [(0.0,34.0),(4.0,31.0),(8.0,27.0),(12.0,23.0),(16.0,17.0),(20.0,10.0)],
        "kaynak": "Wilo Rexa CUT katalog performans grafiği"
    },
    {
        "marka": "Wilo",
        "seri": "Rexa CUT",
        "model": "Rexa CUT GI03.41",
        "q_min": 0.0, "q_max": 20.0, "h_max": 41.0,
        "p2_kw": 3.9,
        "curve": [(0.0,41.0),(4.0,38.0),(8.0,34.0),(12.0,29.0),(16.0,22.0),(20.0,14.0)],
        "kaynak": "Wilo Rexa CUT katalog performans grafiği"
    },
]

def _egri_degeri(curve, q):
  if not curve or q < curve[0][0] or q > curve[-1][0]:
    return None
  for i in range(len(curve) - 1):
    q0, h0 = curve[i]
    q1, h1 = curve[i + 1]
    if q0 <= q <= q1:
      if q1 == q0:
        return h0
      oran = (q - q0) / (q1 - q0)
      return h0 + oran * (h1 - h0)
  return curve[-1][1]


def pompa_uretici_sec(q_m3h, h_mss, marka_secimi="Otomatik (Wilo + Grundfos)"):
  adaylar = []
  for pompa in URETICI_POMPA_VERITABANI:
    if marka_secimi == "Wilo" and pompa["marka"] != "Wilo":
      continue
    if marka_secimi == "Grundfos" and pompa["marka"] != "Grundfos":
      continue
    h_egri = _egri_degeri(pompa["curve"], q_m3h)
    if h_egri is None or h_egri < h_mss:
      continue
    adaylar.append((h_egri - h_mss, pompa["p2_kw"], pompa))

  if not adaylar:
    return None
  adaylar.sort(key=lambda x: (x[0], x[1]))
  secilen = dict(adaylar[0][2])
  secilen["h_calisma"] = _egri_degeri(secilen["curve"], q_m3h)
  return secilen


def pompa_secim_egrisi(q_m3h, h_mss, marka_secimi="Otomatik (Wilo + Grundfos)"):
  poz, _, durum = pompa_pozu_sec(q_m3h, h_mss)
  if durum == "UYGUN":
    model = pompa_uretici_sec(q_m3h, h_mss, marka_secimi)
    if model:
      q_egrisi = [p[0] for p in model["curve"]]
      h_egrisi = [p[1] for p in model["curve"]]
      baslik = f"{model['marka']} {model['model']} - Üretici Q-H Eğrisi"
      return q_egrisi, h_egrisi, baslik, model

    kayit = next(x for x in POZ_POMPA_TABLOSU if x["poz"] == poz)
    q0, q1 = kayit["qmin"], kayit["qmax"]
    h0, h1 = kayit["hmax"], kayit["hmin"]
    q_egrisi = [q0 + (q1 - q0) * i / 100 for i in range(101)]
    h_egrisi = [h0 + (h1 - h0) * ((q - q0) / (q1 - q0)) for q in q_egrisi]
    return q_egrisi, h_egrisi, f"Poz {poz} Sınır Zarfı - Üretici modeli bulunamadı", None

  q_egrisi = [2, 22]
  h_egrisi = [22, 2]
  return q_egrisi, h_egrisi, "Sınır Dışı Çalışma Noktası!", None


def pompa_grafigi_png(q_egrisi, h_egrisi, q_calisma, h_calisma, baslik, anonim=False):
  fig, ax = plt.subplots(figsize=(7.0, 3.8))
  ax.plot(q_egrisi, h_egrisi, linewidth=2.0, label=("Pompa Performans Eğrisi" if anonim else baslik))
  ax.scatter([q_calisma], [h_calisma], s=65, zorder=5, label=f"Çalışma Noktası ({q_calisma:.2f} m³/h, {h_calisma:.2f} mSS)")
  ax.set_xlabel("Debi Q [m³/h]")
  ax.set_ylabel("Basma Yüksekliği H [mSS]")
  ax.set_title("Pompa Performans Eğrisi" if anonim else baslik, fontsize=11)
  ax.grid(True, alpha=0.3)
  ax.legend(fontsize=8)
  fig.tight_layout()
  buf = io.BytesIO()
  fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
  plt.close(fig)
  buf.seek(0)
  return buf


secilen_psp_listesi = st.multiselect(
    "Projede yer alacak Pis Su Terfi Pompalarını seçin:",
    [f"PSP-{i:02d}" for i in range(1, 11)],
    default=["PSP-01"],
    key="secilen_psp_listesi",
)

poz_rapora_eklensin_mi = st.checkbox(
    "Cihaz Poz Numarasını Rapora Aktar", value=True, key="poz_aktar_chk"
)

psp_parametreleri = {}

pompa_marka_secimi = st.selectbox(
    "Pompa üreticisi / seçim modu",
    ["Otomatik (Wilo + Grundfos)", "Wilo", "Grundfos"],
    index=0,
    key="pompa_marka_secimi",
)

if secilen_psp_listesi:
  st.write(
      "Her bir terfi pompası çukuru için bina tipi, armatürler, emniyet"
      " faktörleri ve asıl/yedek adetlerini ayrı ayrı girin:"
  )

  for psp in secilen_psp_listesi:
    with st.expander(
        f"⚙️ {psp} Özel Debi ve Güç Hesap Modülü", expanded=True
    ):
      bina_tipi = st.selectbox(
          f"{psp} Bina Kullanım Türü",
          [
              (
                  "Evler, restoranlar, misafir evleri, oteller, ofis binaları"
                  " (Düzensiz kullanım) [k=0.5]"
              ),
              (
                  "Hastaneler, geniş gıda tesisleri, oteller vb. [k=0.7]"
              ),
              (
                  "Okullar, çamaşırhaneler, umumi tuvaletler ve duşlar"
                  " (Düzenli/Sık kullanım) [k=1.0]"
              ),
              (
                  "Endüstriyel laboratuvarlar vb. özel kullanım tesisleri"
                  " [k=1.2]"
              ),
          ],
          key=f"{psp}_bina",
      )

      if "Düzensiz" in bina_tipi:
        k_varsayilan = 0.5
      elif "Hastaneler" in bina_tipi:
        k_varsayilan = 0.7
      elif "Okullar" in bina_tipi:
        k_varsayilan = 1.0
      else:
        k_varsayilan = 1.2

      k_key = f"{psp}_k"
      if k_key not in st.session_state or st.session_state.get(
          f"{psp}_bina_eski"
      ) != bina_tipi:
        st.session_state[k_key] = k_varsayilan
        st.session_state[f"{psp}_bina_eski"] = bina_tipi

      k_katsayisi = st.number_input(
          f"{psp} Eşzamanlık Katsayısı (k)",
          min_value=0.1,
          max_value=2.0,
          value=st.session_state[k_key],
          step=0.05,
          key=k_key,
      )

      st.write(f"**{psp} için Armatür Adetleri:**")
      ac1, ac2, ac3 = st.columns(3)
      with ac1:
        adet_hela = st.number_input(
            f"{psp} Hela/Klozet (8 Y.B.)", min_value=0, value=4, key=f"{psp}_hela"
        )
        adet_lavabo = st.number_input(
            f"{psp} Lavabo/Bide (2 Y.B.)",
            min_value=0,
            value=6,
            key=f"{psp}_lav",
        )
      with ac2:
        adet_banyo = st.number_input(
            f"{psp} Küvet/Duş (7 Y.B.)", min_value=0, value=2, key=f"{psp}_ban"
        )
        adet_evye = st.number_input(
            f"{psp} Eviye (4 Y.B.)", min_value=0, value=1, key=f"{psp}_evy"
        )
      with ac3:
        adet_suzgec = st.number_input(
            f"{psp} Yer Süzgeci (2 Y.B.)", min_value=0, value=4, key=f"{psp}_suz"
        )
        adet_otopark = st.number_input(
            f"{psp} Otopark Süzgeci (6 Y.B.)",
            min_value=0,
            value=2,
            key=f"{psp}_oto",
        )

      ac4, ac5, ac6 = st.columns(3)
      with ac4:
        adet_basincli = st.number_input(
            f"{psp} Basınçlı Yıkayıcı (10 Y.B.)",
            min_value=0,
            value=0,
            key=f"{psp}_bas",
        )
      with ac5:
        adet_camasir = st.number_input(
            f"{psp} Çamaşır/Bulaşık (10 Y.B.)",
            min_value=0,
            value=1,
            key=f"{psp}_cam",
        )
      with ac6:
        adet_pisuvar = st.number_input(
            f"{psp} Pisuar (1 Y.B.)", min_value=0, value=0, key=f"{psp}_pis"
        )

      yb_hela = adet_hela * 8
      yb_lavabo = adet_lavabo * 2
      yb_banyo = adet_banyo * 7
      yb_evye = adet_evye * 4
      yb_suzgec = adet_suzgec * 2
      yb_otopark = adet_otopark * 6
      yb_basincli = adet_basincli * 10
      yb_camasir = adet_camasir * 10
      yb_pisuvar = adet_pisuvar * 1

      toplam_yb = (
          yb_hela
          + yb_lavabo
          + yb_banyo
          + yb_evye
          + yb_suzgec
          + yb_otopark
          + yb_basincli
          + yb_camasir
          + yb_pisuvar
      )

      net_q_lps = (
          k_katsayisi * math.sqrt(toplam_yb) if toplam_yb > 0 else 0.0
      )
      net_q_m3h = round(net_q_lps * 3.6, 2)

      st.markdown("---")

      emniyet_secenekleri = {
          "Emniyet Ekleme (1.0)": 1.0,
          "%10 Emniyet (1.10)": 1.10,
          "%15 Emniyet (1.15)": 1.15,
          "%20 Emniyet (1.20)": 1.20,
          "%25 Emniyet (1.25)": 1.25,
          "%30 Emniyet (1.30)": 1.30,
          "%40 Emniyet (1.40)": 1.40,
          "%50 Emniyet (1.50)": 1.50,
      }
      emniyet_etiket = st.selectbox(
          f"{psp} Debi Emniyet Oranı",
          list(emniyet_secenekleri.keys()),
          key=f"{psp}_emniyet_secim",
      )
      emniyet_katsayisi = emniyet_secenekleri[emniyet_etiket]

      emniyetli_q_m3h = round(net_q_m3h * emniyet_katsayisi, 2)

      v_key = f"{psp}_v_num"
      if v_key not in st.session_state or st.session_state.get(
          f"{psp}_yb_eski"
      ) != toplam_yb or st.session_state.get(f"{psp}_k_eski") != k_katsayisi or st.session_state.get(f"{psp}_emniyet_eski") != emniyet_katsayisi:
        st.session_state[v_key] = max(1.0, emniyetli_q_m3h)
        st.session_state[f"{psp}_yb_eski"] = toplam_yb
        st.session_state[f"{psp}_k_eski"] = k_katsayisi
        st.session_state[f"{psp}_emniyet_eski"] = emniyet_katsayisi

      toplam_v_val = st.number_input(
          f"{psp} Toplam Debi (Q_toplam) [m³/h] (Emniyetli)",
          min_value=1.0,
          max_value=100.0,
          value=st.session_state[v_key],
          step=0.5,
          key=v_key,
      )

      h_val = st.number_input(
          f"{psp} Basma Yüksekliği (H) [mSS]",
          min_value=1.0,
          max_value=30.0,
          value=12.0,
          step=0.5,
          key=f"{psp}_h_num",
      )

      asil_adedi = st.selectbox(
          f"{psp} Asıl Pompa Adedi", [1, 2, 3], index=0, key=f"{psp}_asil"
      )
      yedek_adedi = st.selectbox(
          f"{psp} Yedek Pompa Adedi", [1, 2], index=0, key=f"{psp}_yedek"
      )

      pompa_basina_v = toplam_v_val / asil_adedi

      hesap = pompa_hidrolik_hesap(pompa_basina_v, h_val, 0.60, 0.90)
      hesaplanan_poz, hesaplanan_tanim, poz_durumu = pompa_pozu_sec(
          pompa_basina_v, h_val
      )
      q_egrisi, h_egrisi, egrisi_basligi, secilen_uretici_pompa = pompa_secim_egrisi(
          pompa_basina_v, h_val, pompa_marka_secimi
      )

      st.metric(
          "Tek Pompa Debisi (Asıla Bölünen)", f"{pompa_basina_v:.2f} m³/h"
      )
      st.metric("Tek Pompa Elektrik Gücü", f"{hesap['motor_secim_kw']:.2f} kW")

      if poz_durumu == "UYGUN":
        st.success(f"✅ Uygun Poz: **{hesaplanan_poz}**")
      else:
        st.error("❌ HATA: Tek pompa debi/basıncı sınır dışındadır!")

      st.info(f"📌 **Poz Tanımı:** {hesaplanan_tanim}")

      program_grafik = pompa_grafigi_png(
          q_egrisi, h_egrisi, pompa_basina_v, h_val, egrisi_basligi, anonim=False
      )
      st.image(program_grafik, caption=egrisi_basligi, use_container_width=True)

      toplam_adet = asil_adedi + yedek_adedi
      adet_metin = (
          f"{toplam_adet} ({asil_adedi} Asıl"
          f" {'+ ' + str(yedek_adedi) + ' Yedek' if yedek_adedi > 0 else ''})"
      )
      tip_metin = (
          "Dalgıç Tip, Parçalayıcı Bıçaklı, Kesme Düzenekli Pis Su Terfi"
          " Pompası"
      )

      tablo_satirlari = []
      if adet_hela > 0:
        tablo_satirlari.append(("Hela / Klozet", 8, adet_hela, yb_hela))
      if adet_lavabo > 0:
        tablo_satirlari.append(("Lavabo / Bide", 2, adet_lavabo, yb_lavabo))
      if adet_banyo > 0:
        tablo_satirlari.append(("Küvet / Duş", 7, adet_banyo, yb_banyo))
      if adet_evye > 0:
        tablo_satirlari.append(("Eviye", 4, adet_evye, yb_evye))
      if adet_suzgec > 0:
        tablo_satirlari.append(("Yer Süzgeci", 2, adet_suzgec, yb_suzgec))
      if adet_otopark > 0:
        tablo_satirlari.append(("Otopark Süzgeci", 6, adet_otopark, yb_otopark))
      if adet_basincli > 0:
        tablo_satirlari.append(
            ("Basınçlı Yıkayıcı", 10, adet_basincli, yb_basincli)
        )
      if adet_camasir > 0:
        tablo_satirlari.append(
            ("Çamaşır/Bulaşık Makinası", 10, adet_camasir, yb_camasir)
        )
      if adet_pisuvar > 0:
        tablo_satirlari.append(("Pisuar", 1, adet_pisuvar, yb_pisuvar))

      psp_parametreleri[psp] = {
          "bina_tipi": bina_tipi,
          "k_katsayisi": k_katsayisi,
          "emniyet_etiket": emniyet_etiket,
          "emniyet_katsayisi": emniyet_katsayisi,
          "net_q_lps": net_q_lps,
          "net_q_m3h": net_q_m3h,
          "toplam_yb": toplam_yb,
          "q_lps_toplam": (toplam_v_val / 3.6),
          "v_toplam": toplam_v_val,
          "v_tek": pompa_basina_v,
          "q_lps_tek": hesap["q_lps"],
          "h": h_val,
          "guc": hesap["motor_secim_kw"],
          "asil_adedi": asil_adedi,
          "adet_str": adet_metin,
          "tip": tip_metin,
          "poz": hesaplanan_poz,
          "poz_durumu": poz_durumu,
          "poz_tanim": hesaplanan_tanim,
          "pompa_markasi": secilen_uretici_pompa["marka"] if secilen_uretici_pompa else "",
          "pompa_modeli": secilen_uretici_pompa["model"] if secilen_uretici_pompa else "",
          "pompa_curve": secilen_uretici_pompa["curve"] if secilen_uretici_pompa else list(zip(q_egrisi, h_egrisi)),
          "pompa_egrisi_basligi": egrisi_basligi,
          "pompa_h_egrisi_calisma": secilen_uretici_pompa["h_calisma"] if secilen_uretici_pompa else h_val,
          "pompa_kaynak": secilen_uretici_pompa["kaynak"] if secilen_uretici_pompa else "Poz sınır eğrisi",
          "tablo_satirlari": tablo_satirlari,
      }

st.subheader("Pis Su Terfi Pompası Genel Esasları ve Notlar")
terfi_keys = ["terfi_sec_1", "terfi_sec_2", "terfi_sec_3", "terfi_sec_4"]
_toplu_secim_butonlari(terfi_keys)

terfi_sec_1 = st.checkbox(
    "Kot kurtarmayan bodrum kat atık suları için paslanmaz gövdeli, parçalayıcı"
    " bıçaklı pis su atık su terfi pompaları seçilmiştir.",
    key="terfi_sec_1",
    value=True,
)
terfi_sec_2 = st.checkbox(
    "Pompalar yedekli çalışacak şekilde otomasyona bağlanacaktır.",
    key="terfi_sec_2",
    value=True,
)
terfi_sec_3 = st.checkbox(
    "Terfi çukurunda sıvı seviye şalterleri (şamandıra) bulunacak, su"
    " seviyesine göre pompalar otomatik devreye girip çıkacaktır.",
    key="terfi_sec_3",
    value=True,
)
terfi_sec_4 = st.checkbox(
    "Pompa basma hatlarında geri akışı önlemek için çekvalf ve bakım kolaylığı"
    " için sürgülü/kelebek vana kullanılacaktır.",
    key="terfi_sec_4",
    value=True,
)

ek_terfi_notu = st.text_area(
    "İlave Pis Su Terfi Pompası Genel Esasları (Her satıra bir tane)",
    "",
    height=80,
)


# --- 6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ ---
st.header("6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ")
st.subheader("6.3.1 KULLANMA SOĞUK SUYU DEPOSU SEÇİMİ")

poz_gosterilsin_mi = st.checkbox(
    "Poz numarasını göster",
    value=True,
    key="poz_gosterilsin_mi",
)

genel_bilgiler_tab = st.container()
with genel_bilgiler_tab:
    st.markdown("#### Genel Bilgiler")
    st.markdown("#### 6.3.1.1 KULLANMA SUYU İHTİYACININ BELİRLENMESİ")

    # SU TÜKETİMİ HESABI
    su_tuketim_secenekleri = {
        "Fabrikalar": ("Kişi", 45.0),
        "Görev başındakiler": ("Kişi", 45.0),
        "Hemşireler": ("Kişi", 135.0),
        "Duşlu oteller": ("Kişi", 90.0),
        "Küvetli oteller": ("Kişi", 150.0),
        "Bürolar": ("Kişi", 45.0),
        "Lokantalar": ("Kişi", 7.0),
        "Okullar - Gündüzlü": ("Kişi", 45.0),
        "Okullar - Yatılı": ("Kişi", 135.0),
        "Bahçe sulama": ("m²", 1.5),
        "Binek otosu": ("Adet", 100.0),
        "Askeri binalar - Yatılı": ("Kişi", 135.0),
        "Askeri binalar - Yatılı olmayan": ("Kişi", 45.0),
    }

    st.markdown("##### Su Tüketim Değerleri Tablosu")
    su_tuketim_tablosu = [
        {
            "Kullanım amacı": kategori,
            "Birim": birim,
            "Birim tüketim değeri": f"{deger:g} L/{birim}/gün",
        }
        for kategori, (birim, deger) in su_tuketim_secenekleri.items()
    ]

    # Kullanıcının paylaştığı ek kaynak tablosunda bulunan ve mevcut
    # hesaplama listesindeki tekil değerlerden farklı/eksik olan değerler.
    # Aralıklar (ör. 60 / 80) kaynak tablodaki haliyle gösterilir.
    ek_su_tuketim_tablosu = [
        {"Kullanım amacı": "Konutlar - Lavabolu", "Birim": "Kişi", "Birim tüketim değeri": "60 / 80 L/Kişi-gün"},
        {"Kullanım amacı": "Konutlar - Duşlu", "Birim": "Kişi", "Birim tüketim değeri": "80 / 115 L/Kişi-gün"},
        {"Kullanım amacı": "Konutlar - Küvetli", "Birim": "Kişi", "Birim tüketim değeri": "120 / 200 L/Kişi-gün"},
        {"Kullanım amacı": "Oteller - Duşlu", "Birim": "Kişi", "Birim tüketim değeri": "100 L/Kişi-gün"},
        {"Kullanım amacı": "Oteller - Küvetli", "Birim": "Kişi", "Birim tüketim değeri": "150 / 200 L/Kişi-gün"},
        {"Kullanım amacı": "Hastaneler", "Birim": "Kişi", "Birim tüketim değeri": "200 / 500 L/Kişi-gün"},
        {"Kullanım amacı": "Okullar", "Birim": "Kişi", "Birim tüketim değeri": "5 L/Kişi-gün"},
        {"Kullanım amacı": "Çocuk Yuvaları", "Birim": "Kişi", "Birim tüketim değeri": "80 / 150 L/Kişi-gün"},
        {"Kullanım amacı": "Kreşler", "Birim": "Kişi", "Birim tüketim değeri": "100 / 150 L/Kişi-gün"},
        {"Kullanım amacı": "Kışlalar", "Birim": "Kişi", "Birim tüketim değeri": "60 / 80 L/Kişi-gün"},
        {"Kullanım amacı": "Lokantalar (kaynak tablosu)", "Birim": "Kişi", "Birim tüketim değeri": "20 / 150 L/Kişi-gün"},
        {"Kullanım amacı": "Bahçe Sulama Bir Seferde", "Birim": "m²", "Birim tüketim değeri": "1,5 L/m²"},
        {"Kullanım amacı": "Oto Yıkama - Temizlik", "Birim": "Gün", "Birim tüketim değeri": "100 L/Gün"},
    ]
    su_tuketim_tablosu.extend(ek_su_tuketim_tablosu)
    st.table(su_tuketim_tablosu)
    st.caption(
        "Not: Birim tüketim değerleri, kullanıcı tarafından yüklenen "
        "dokümanda belirtilen TS-1258 kaynaklı değerler esas alınarak "
        "uygulamaya aktarılmıştır."
    )

    st.markdown("##### Su İhtiyacı Hesabı")
    secilen_su_kategorileri = st.multiselect(
        "Kullanım amacı / su tüketim kategorileri (birden fazla seçilebilir)",
        options=list(su_tuketim_secenekleri.keys()),
        key="secilen_su_kategorileri",
    )

    su_hesap_detaylari = []
    su_gunluk_ihtiyac_litre = 0.0
    if secilen_su_kategorileri:
        st.markdown("**Seçilen kategoriler için miktarları giriniz:**")
        for sira, kategori in enumerate(secilen_su_kategorileri):
            su_birim, su_birim_degeri = su_tuketim_secenekleri[kategori]
            su_miktari = st.number_input(
                f"{kategori} miktarı ({su_birim})",
                min_value=0.0,
                value=1.0,
                step=1.0,
                key=f"su_miktari_{sira}",
            )
            kategori_ihtiyaci_litre = su_miktari * su_birim_degeri
            su_gunluk_ihtiyac_litre += kategori_ihtiyaci_litre
            su_hesap_detaylari.append({
                "kategori": kategori, "birim": su_birim,
                "birim_degeri": su_birim_degeri, "miktar": su_miktari,
                "ihtiyac_litre": kategori_ihtiyaci_litre,
            })
        su_gunluk_ihtiyac_m3 = su_gunluk_ihtiyac_litre / 1000.0
        st.metric(
            "Günlük toplam su ihtiyacı",
            f"{su_gunluk_ihtiyac_litre:,.2f} L/gün".replace(",", "X").replace(".", ",").replace("X", "."),
        )
        st.caption(
            f"Seçilen {len(secilen_su_kategorileri)} kategori için toplam: "
            f"{su_gunluk_ihtiyac_litre:g} L/gün = {su_gunluk_ihtiyac_m3:g} m³/gün"
        )
    else:
        su_gunluk_ihtiyac_m3 = 0.0
        st.info("Hesaplama yapmak için en az bir su tüketim kategorisi seçiniz.")

    st.markdown("##### Su Deposu Depolama Süresi ve Gerekli Hacim")
    depo_sure_gun = st.number_input(
        "Kaç günlük su ihtiyacı için depo seçilecek?", min_value=1.0,
        value=1.0, step=1.0, key="depo_sure_gun",
        help="Depo hacmi, günlük toplam su ihtiyacı ile seçilen gün sayısının çarpımıyla hesaplanır.",
    )
    depo_gerekli_hacim_m3 = su_gunluk_ihtiyac_m3 * depo_sure_gun
    depo_gerekli_hacim_litre = su_gunluk_ihtiyac_litre * depo_sure_gun
    secilen_depo_tipi_metni = ", ".join(sih_depo_tipleri) if sih_sec_depo_tipi and sih_depo_tipleri else ""
    depo_hacmi_basligi = (
        f'Yapının kullanım soğuk suyu ihtiyacını karşılamak için seçilen "{secilen_depo_tipi_metni}" hacmi'
        if secilen_depo_tipi_metni
        else "Yapının kullanım soğuk suyu ihtiyacını karşılamak için seçilen su deposu hacmi"
    )
    # Kapasiteye göre otomatik poz seçimi. Hesaplanan hacme en yakın standart kapasite seçilir.
    # Böylece kapasite her zaman zorunlu olarak bir üst değere yuvarlanmaz; alt kapasite
    # hesaplanan değere daha yakınsa alt kapasite seçilebilir. Son karar kullanıcıdadır.
    depo_poz_kapasiteleri = {
        "Paslanmaz Çelik Modüler su deposu": [
            (1.25, "25.150.1201"), (2.50, "25.150.1202"), (3.75, "25.150.1203"),
            (5.00, "25.150.1204"), (6.25, "25.150.1205"), (7.50, "25.150.1206"),
            (10.0, "25.150.1207"), (12.5, "25.150.1208"), (15.0, "25.150.1209"),
            (20.0, "25.150.1210"), (22.5, "25.150.1211"), (25.0, "25.150.1212"),
            (30.0, "25.150.1213"), (37.5, "25.150.1214"), (40.0, "25.150.1215"),
            (45.0, "25.150.1216"), (50.0, "25.150.1217"), (56.0, "25.150.1218"),
            (59.6, "25.150.1219"), (62.0, "25.150.1220"), (75.0, "25.150.1221"),
            (90.0, "25.150.1222"), (93.2, "25.150.1223"), (104.2, "25.150.1224"),
            (112.0, "25.150.1225"), (121.5, "25.150.1226"),
        ],
        "Galvaniz Çelik Modüler su deposu": [
            (1.25, "25.150.1301"), (2.50, "25.150.1302"), (3.75, "25.150.1303"),
            (5.00, "25.150.1304"), (6.25, "25.150.1305"), (7.50, "25.150.1306"),
            (10.0, "25.150.1307"), (12.5, "25.150.1308"), (15.0, "25.150.1309"),
            (20.0, "25.150.1310"), (22.5, "25.150.1311"), (25.0, "25.150.1312"),
            (30.0, "25.150.1313"), (37.5, "25.150.1314"), (40.0, "25.150.1315"),
            (45.0, "25.150.1316"), (50.0, "25.150.1317"), (56.0, "25.150.1318"),
            (59.6, "25.150.1319"), (62.0, "25.150.1320"), (75.0, "25.150.1321"),
        ],
        "GRP (Cam Takviyeli Polyester) Modüler su deposu": [
            (1.0, "25.150.1601"), (3.0, "25.150.1602"), (5.0, "25.150.1603"),
            (10.0, "25.150.1604"), (15.0, "25.150.1605"), (20.0, "25.150.1606"),
            (30.0, "25.150.1607"), (40.0, "25.150.1608"), (50.0, "25.150.1609"),
            (60.0, "25.150.1610"), (70.0, "25.150.1611"), (80.0, "25.150.1612"),
            (90.0, "25.150.1613"), (100.0, "25.150.1614"), (120.0, "25.150.1615"),
            (150.0, "25.150.1616"), (180.0, "25.150.1617"), (200.0, "25.150.1618"),
            (240.0, "25.150.1619"), (270.0, "25.150.1620"), (300.0, "25.150.1621"),
        ],
    }

    otomatik_poz_kayitlari = []

    def en_yakin_kapasite_kaydi(kayitlar, hedef_m3):
        if not kayitlar:
            return None
        return min(kayitlar, key=lambda kayit: abs(kayit[0] - hedef_m3))

    # Her depo tipi için otomatik başlangıç kapasitesi gösterilir. Kullanıcı isterse
    # kapasiteyi değiştirebilir; son kullanıcı değeri rapora aktarılır.
    for depo_index, depo_tipi in enumerate(sih_depo_tipleri if sih_sec_depo_tipi else []):
        kayitlar = depo_poz_kapasiteleri.get(depo_tipi, [])
        uygun = en_yakin_kapasite_kaydi(kayitlar, depo_gerekli_hacim_m3)
        if uygun:
            otomatik_kapasite, otomatik_poz = uygun
            kapasite_key = f"manuel_depo_kapasitesi_{depo_index}"
            onceki_otomatik_key = f"onceki_otomatik_depo_kapasitesi_{depo_index}"
            poz_key = f"manuel_depo_pozu_{depo_index}"
            poz_elle_key = f"depo_pozunu_elle_duzenle_{depo_index}"

            # Hesap sonucu değiştiğinde, kullanıcı daha önce elle müdahale etmediyse
            # giriş alanı yeni otomatik kapasiteyle güncellenir. Elle değiştirilmiş
            # değerler korunur.
            mevcut_kapasite = st.session_state.get(kapasite_key)
            onceki_otomatik = st.session_state.get(onceki_otomatik_key)
            if mevcut_kapasite is None or mevcut_kapasite == onceki_otomatik:
                st.session_state[kapasite_key] = float(otomatik_kapasite)
            st.session_state[onceki_otomatik_key] = float(otomatik_kapasite)

            st.caption(
                f"Otomatik hesaplanan en yakın standart kapasite: {otomatik_kapasite:g} m³ "
                f"(hesaplanan ihtiyaç: {depo_gerekli_hacim_m3:g} m³)"
            )
            manuel_kapasite = st.number_input(
                f"{depo_tipi} için depo kapasitesi (m³) — otomatik gelir, elle değiştirilebilir",
                min_value=0.001,
                step=0.5,
                key=kapasite_key,
                help="Alan başlangıçta otomatik seçilen en yakın standart kapasiteyle doldurulur. İsterseniz son onay olarak elle değiştirebilirsiniz.",
            )

            kapasiteye_uygun_kayit = en_yakin_kapasite_kaydi(kayitlar, manuel_kapasite)
            kapasiteye_uygun_poz = kapasiteye_uygun_kayit[1] if kapasiteye_uygun_kayit else ""

            poz_elle_duzenle = st.checkbox(
                f"{depo_tipi} poz numarasını elle düzenle",
                value=False,
                key=poz_elle_key,
            )
            if poz_elle_duzenle:
                if poz_key not in st.session_state:
                    st.session_state[poz_key] = kapasiteye_uygun_poz
                manuel_poz = st.text_input(
                    f"{depo_tipi} için seçilen poz numarası",
                    key=poz_key,
                )
                kullanilacak_poz = manuel_poz.strip()
            else:
                # Elle düzenleme kapalıyken poz, elle girilen kapasiteye en yakın
                # standart kapasiteye göre otomatik olarak yeniden belirlenir.
                kullanilacak_poz = kapasiteye_uygun_poz
                st.caption(f"Kapasiteye en yakın otomatik seçilen poz: {kullanilacak_poz}")

            # Ekranda gösterilecek kapasite, kullanıcı girişindeki serbest değerden değil,
            # seçilen poz numarasının gerçek kapasite karşılığından alınır.
            poz_kapasite_kaydi = next(
                (kayit for kayit in kayitlar if kayit[1] == kullanilacak_poz),
                None,
            )
            poz_karsiligi_kapasite = (
                poz_kapasite_kaydi[0]
                if poz_kapasite_kaydi is not None
                else (kapasiteye_uygun_kayit[0] if kapasiteye_uygun_kayit else manuel_kapasite)
            )
            otomatik_poz_kayitlari.append((depo_tipi, poz_karsiligi_kapasite, kullanilacak_poz))
        else:
            st.warning(f"{depo_tipi} için kapasite listesi bulunamadı.")

    if poz_gosterilsin_mi and otomatik_poz_kayitlari:
        for depo_tipi, secilen_kapasite, secilen_poz in otomatik_poz_kayitlari:
            st.markdown(
                f'<div style="color:#000000;"><strong>Seçilen poz numarası:</strong> {secilen_poz} ' 
                f'<strong>(Kapasite: {secilen_kapasite:g} m³)</strong></div>',
                unsafe_allow_html=True,
            )
    elif poz_gosterilsin_mi and sih_depo_tipleri:
        st.warning("Hesaplanan hacim, tanımlı kapasite listesinin üzerindedir.")

    rapor_gosterilecek_kapasite_m3 = otomatik_poz_kayitlari[0][1] if otomatik_poz_kayitlari else depo_gerekli_hacim_m3
    rapor_gosterilecek_kapasite_litre = rapor_gosterilecek_kapasite_m3 * 1000.0
    depo_hacmi_degeri = (f"{rapor_gosterilecek_kapasite_litre:,.0f} L ({rapor_gosterilecek_kapasite_m3:,.3f} m³)").replace(",", "X").replace(".", ",").replace("X", ".")
    st.markdown(
        f'<div style="font-size:1.05rem; color:#000000;">{depo_hacmi_basligi}: '
        f'<strong style="color:#000000;">{depo_hacmi_degeri}</strong>&#39;dir.</div>',
        unsafe_allow_html=True,
    )

    st.caption(f"Hesap: {su_gunluk_ihtiyac_m3:g} m³/gün × {depo_sure_gun:g} gün = {depo_gerekli_hacim_m3:g} m³ ({depo_gerekli_hacim_litre:g} L)")
    su_tuketim_tipi = ", ".join(secilen_su_kategorileri) if secilen_su_kategorileri else "Seçim yapılmadı"
    su_birim = "-"; su_birim_degeri = 0.0; su_miktari = 0.0

poz_numarasi = ", ".join(k[2] for k in otomatik_poz_kayitlari) if poz_gosterilsin_mi else ""
poz_tipi = ", ".join(k[0] for k in otomatik_poz_kayitlari) if poz_gosterilsin_mi else ""

depo_keys = ["depo_sec_1", "depo_sec_2", "depo_sec_3"]
_toplu_secim_butonlari(depo_keys)

depo_sec_1 = st.checkbox(
    "Kullanma soğuk suyu deposu hacmi; binanın kullanım amacı, kullanıcı "
    "sayısı, kişi başına günlük su tüketimi, kullanım sürekliliği ve ihtiyaç "
    "duyulan su rezervi dikkate alınarak belirlenecektir. Depo kapasitesi, "
    "binanın günlük su ihtiyacını karşılayacak ve işletme koşullarında yeterli "
    "su rezervi sağlayacak şekilde tasarlanacaktır.",
    key="depo_sec_1",
    value=True,
)
depo_sec_2 = st.checkbox(
    "Su deposu içerisinde su kalitesinin korunması ve ölü hacim oluşumunun"
    " önlenmesi için bölme perdeleri yer alacaktır.",
    key="depo_sec_2",
    value=True,
)
depo_sec_3 = st.checkbox(
    "Su deposunda taşma, deşarj, havalandırma boruları ile bakım ve temizlik"
    " için adam geçiş kapağı (manhole) bulunacaktır.",
    key="depo_sec_3",
    value=True,
)

ek_depo_notu = st.text_area(
    "İlave Kullanma Soğuk Suyu Deposu Seçim Maddesi (Her satıra bir tane)",
    "",
    height=80,
)


def rapor_word_stillerini_uygula(doc):
  for style_name in [
      "Normal",
      "Body Text",
      "List Paragraph",
      "List Bullet",
      "List Number",
  ]:
    try:
      stl = doc.styles[style_name]
      stl.font.name = "Times New Roman"
      stl._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
      stl.font.size = Pt(12)
    except KeyError:
      pass

  h1 = doc.styles["Heading 1"]
  h1.font.name = "Times New Roman"
  h1._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
  h1.font.size = Pt(14)
  h1.font.bold = True
  h1.paragraph_format.page_break_before = True

  for level in [2, 3, 4]:
    h = doc.styles[f"Heading {level}"]
    h.font.name = "Times New Roman"
    h._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    h.font.size = Pt(12)
    h.font.bold = True

  for paragraph in doc.paragraphs:
    for run in paragraph.runs:
      run.font.name = "Times New Roman"
      run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
      if paragraph.style and paragraph.style.name.startswith("Heading 1"):
        run.font.size = Pt(14)
        run.font.bold = True
      elif paragraph.style and paragraph.style.name.startswith("Heading"):
        run.font.size = Pt(12)
        run.font.bold = True
      elif run.font.size is None:
        run.font.size = Pt(12)

  for table in doc.tables:
    for row in table.rows:
      for cell in row.cells:
        for paragraph in cell.paragraphs:
          for run in paragraph.runs:
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
            run.font.size = Pt(12)


# Rapor Oluştur Butonu
if st.button("Raporu Oluştur (.docx)"):
  gecersiz_var = any(
      p["poz_durumu"] != "UYGUN" for p in psp_parametreleri.values()
  )
  if gecersiz_var:
    st.error(
        "❌ Rapor oluşturulamadı! Seçilen pompalardan biri veya daha fazlasının"
        " tek pompa debisi poz sınırları dışındadır."
    )
  else:
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

    cover_section = doc.sections[0]
    cover_section.top_margin = Inches(1.5)
    cover_section.bottom_margin = Inches(1.5)
    cover_section.left_margin = Inches(1.2)
    cover_section.right_margin = Inches(1.2)

    # 1. SAYFA: KAPAK SAYFASI
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

    # 2. SAYFA: İÇİNDEKİLER
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

    # 3. SAYFA: GÖVDE
    body_section = doc.add_section()
    body_section.top_margin = Inches(1.2)
    body_section.bottom_margin = Inches(1.2)
    body_section.left_margin = Inches(1.2)
    body_section.right_margin = Inches(1.2)

    # --- 1. GENEL BİLGİLER ---
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
    doc.add_paragraph(
        "Yapılarda aşağıdaki mekanik tesisat sistemleri uygulanacaktır."
    )

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
    doc.add_paragraph(
        f"Yapının inşa edileceği ''{secilen_il}'' için kabul edilen dış hava"
        " koşulları aşağıdaki gibidir:"
    )
    doc.add_paragraph(
        f"• KIŞ: {iklim_veri['kis_kt']} °C Kuru Termometre (KT) ,"
        f" {iklim_veri['kis_yt']} °C Yaş Termometre (YT)"
    )
    doc.add_paragraph(
        f"• YAZ: {iklim_veri['yaz_kt']} °C Kuru Termometre (KT) ,"
        f" {iklim_veri['yaz_yt']} °C Yaş Termometre (YT)"
    )
    doc.add_paragraph(f"• Enlem: {iklim_veri['enlem']}")
    doc.add_paragraph(f"• Boylam: {iklim_veri['boylam']}")
    doc.add_paragraph(
        f"• Deniz seviyesinden yüksekliği (Rakım): {iklim_veri['rakim']} m."
    )
    doc.add_paragraph(f"• Günlük Sıcaklık Farkı (GSF): {iklim_veri['gsf']} °C")

    # --- 6. SIHHİ TESİSAT ---
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
          " musluğun su verimi olan 0.25 lt/sn yükleme birimi olarak"
          " alınacaktır. Diğer bütün sarfiyatlar bu birime tamamlanacaktır."
      )

    if sih_sec_4 and sih_depo_konumlari:
      if len(sih_depo_konumlari) == 1:
        konum_str = sih_depo_konumlari[0].lower()
      elif len(sih_depo_konumlari) == 2:
        konum_str = (
            f"{sih_depo_konumlari[0].lower()} ve"
            f" {sih_depo_konumlari[1].lower()}"
        )
      else:
        ilkler = ", ".join([k.lower() for k in sih_depo_konumlari[:-1]])
        son = sih_depo_konumlari[-1].lower()
        konum_str = f"{ilkler} ve {son}"

      sihhi_maddeler.append(
          f"Bütün binanın kullanma soğuk su ihtiyacı {konum_str} soğuk"
          " su deposundan sağlanacaktır. Basıncın yetersizliği ve su"
          " kesilmelerine karşın depo hidrofor sistemi uygulanmıştır. TS 1258 ve"
          " ilgili standartlar esas alınacaktır."
      )

    if sih_sec_depo_tipi and sih_depo_tipleri:
      if len(sih_depo_tipleri) == 1:
        tip_str = sih_depo_tipleri[0].lower()
      elif len(sih_depo_tipleri) == 2:
        tip_str = (
            f"{sih_depo_tipleri[0].lower()} ve {sih_depo_tipleri[1].lower()}"
        )
      else:
        ilkler = ", ".join([t.lower() for t in sih_depo_tipleri[:-1]])
        son = sih_depo_tipleri[-1].lower()
        tip_str = f"{ilkler} ve {son}"

      sihhi_maddeler.append(
          "Binada kullanım soğuk su depolaması için "
          f"{tip_str} tipinde su deposu kullanılmıştır."
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

    if sih_sec_7 and sih_sicak_su_yontemleri:
      if len(sih_sicak_su_yontemleri) == 1:
        secilenler_str = sih_sicak_su_yontemleri[0].lower()
      elif len(sih_sicak_su_yontemleri) == 2:
        secilenler_str = (
            f"{sih_sicak_su_yontemleri[0].lower()} ve"
            f" {sih_sicak_su_yontemleri[1].lower()}"
        )
      else:
        ilkler = ", ".join([y.lower() for y in sih_sicak_su_yontemleri[:-1]])
        son = sih_sicak_su_yontemleri[-1].lower()
        secilenler_str = f"{ilkler} ve {son}"

      sihhi_maddeler.append(
          "Kullanma Sıcak suyu üretimi ısı merkezindeki "
          f"{secilenler_str} vasıtasıyla yapılacaktır."
      )

    if sih_sec_8:
      sihhi_maddeler.append(
          "Sıhhi tesisat işlerinde ana dağıtım boruları galvaniz çelik, mahal"
          " içi dağıtım boruları PPRC tipte seçilecektir."
      )

    if sih_sec_9:
      secilen_mahaller = [
          m.lower() for m in sih_yumusak_su_mahalleri if m != "Diğer"
      ]
      if sih_yumusak_su_diger.strip():
        secilen_mahaller.append(sih_yumusak_su_diger.strip().lower())

      if secilen_mahaller:
        if len(secilen_mahaller) == 1:
          mahal_str = secilen_mahaller[0]
        elif len(secilen_mahaller) == 2:
          mahal_str = f"{secilen_mahaller[0]} ve {secilen_mahaller[1]}"
        else:
          ilkler = ", ".join(secilen_mahaller[:-1])
          son = secilen_mahaller[-1]
          mahal_str = f"{ilkler} ve {son}"

        sihhi_maddeler.append(
            f"{mahal_str.capitalize()} mahallerinde yumuşak su kullanılacaktır."
        )

    if sih_sec_10 and sih_sicak_su_isitma_sistemleri:
      if len(sih_sicak_su_isitma_sistemleri) == 1:
        isitma_str = sih_sicak_su_isitma_sistemleri[0].lower()
      elif len(sih_sicak_su_isitma_sistemleri) == 2:
        isitma_str = (
            f"{sih_sicak_su_isitma_sistemleri[0].lower()} ve"
            f" {sih_sicak_su_isitma_sistemleri[1].lower()}"
        )
      else:
        ilkler = ", ".join(
            [s.lower() for s in sih_sicak_su_isitma_sistemleri[:-1]]
        )
        son = sih_sicak_su_isitma_sistemleri[-1].lower()
        isitma_str = f"{ilkler} ve {son}"

      sihhi_maddeler.append(
          f"Kullanım sıcak suyunun ısıtılması {isitma_str} vasıtasıyla"
          " yapılacaktır."
      )

    if sih_sec_12:
      sihhi_maddeler.append(
          "Yağmur suyu toplama yönetmeliğine göre 2 bin metrekareden büyük"
          " parsellerde inşa edilecek tüm binaların çatılarında toplanan yağmur"
          " sularının, bahçe sulama veya arıtılarak bina ihtiyacında kullanılmak"
          " üzere bahçe zemini altında bir depoda toplaması amacıyla 'yağmur"
          " suyu toplama sistemi' yapılması zorunluluğu getirildiği için yağmur"
          " hasadı tesisatı yapılmıştır."
      )

    if ek_sihhi_on_bilgi.strip():
      for es in ek_sihhi_on_bilgi.split("\n"):
        if es.strip():
          sihhi_maddeler.append(es.strip())

    for sm in sihhi_maddeler:
      doc.add_paragraph(sm, style="List Bullet")

    # --- 6.1.1 TEMİZ SU SARFİYAT YÜKLEME BİRİMLERİ VE ÇAP TAYİNİ ---
    doc.add_heading(
        "6.1.1 Temiz Su Sarfiyat Yükleme Birimleri ve Çap Tayini", level=2
    )
    doc.add_paragraph(
        "Sıhhi tesisat boru çaplarının tespitinde ve kullanım yerlerine ait"
        " yükleme birimleri ile debi değerlerinde aşağıdaki tablolar esas"
        " alınmıştır."
    )

    t1_data = [
        ("DN", "PLASTİK", "ÇELİK", "Yükleme Birimi"),
        ("15", "Ø20", '1/2"', "(0-3.0)"),
        ("20", "Ø25", '3/4"', "(3.0-8.0)"),
        ("25", "Ø32", '1"', "(8.0-20.0)"),
        ("32", "Ø40", '1 1/4"', "(20.0-35.0)"),
        ("40", "Ø50", '1 1/2"', "(35.0-50.0)"),
        ("50", "Ø63", '2"', "(50.0-144.0)"),
        ("65", "Ø75", '2 1/2"', "(144.0-368.0)"),
        ("80", "Ø90", '3"', "(368.0-1156.0)"),
        ("100", "Ø125", '4"', "(1156-4900)"),
        ("125", "-", '5"', "(4.900-14.400)"),
        ("150", "-", '6"', "(14.400-40.000)"),
        ("200", "-", '8"', "(40.000-484.000)"),
        ("250", "-", '10"', "(484.000-518.400)"),
        ("300", "-", '12"', "(518.400-1.440.000)"),
    ]
    t1 = doc.add_table(rows=len(t1_data), cols=4)
    t1.style = "Table Grid"
    for r_idx, row in enumerate(t1_data):
      for c_idx, val in enumerate(row):
        t1.cell(r_idx, c_idx).text = val

    doc.add_paragraph()

    # --- 6.2 PİS SU TESİSATI ---
    doc.add_heading("6.2 PİS SU TESİSATI", level=2)

    pis_su_maddeleri = []

    if pissu_sec_1 and pis_su_konumlari and pis_su_gecisler:
      if len(pis_su_konumlari) == 1:
        konum_s = pis_su_konumlari[0].lower()
      else:
        konum_s = (
            f"{', '.join([k.lower() for k in pis_su_konumlari[:-1]])} ve"
            f" {pis_su_konumlari[-1].lower()}"
        )

      if len(pis_su_gecisler) == 1:
        gecis_s = pis_su_gecisler[0].lower()
      else:
        gecis_s = (
            f"{', '.join([g.lower() for g in pis_su_gecisler[:-1]])} ve"
            f" {pis_su_gecisler[-1].lower()}"
        )

      pis_su_maddeleri.append(
          f"Yapının atık suları binanın pik kolonlarla toplanarak {konum_s}"
          f" {gecis_s} rögarlara iletilecektir."
      )

    if pissu_sec_2:
      pis_su_maddeleri.append(
          "Pis su kolonları üzerinde gerekli yerlere temizleme kapakları"
          " yerleştirilmiştir."
      )

    if pissu_sec_3:
      pis_su_maddeleri.append(
          "Tüm teknik hacimlerde, su tahliyesi için ızgaralı kanallar"
          " yapılacaktır."
      )

    if pissu_sec_4:
      pis_su_maddeleri.append(
          "Atık su boruları sessiz PVC boru gibi son teknoloji ürünü borular"
          " kullanılacaktır."
      )

    if pissu_sec_5:
      pis_su_maddeleri.append(
          "Pis su boru çapları yükleme birimi yöntemine göre belirlenmiştir."
      )

    if pissu_sec_6:
      pis_su_maddeleri.append(
          "Pis su akar kotunun kurtarmayan katları bodrum katta pis su çukurunda"
          " toplanıp, pompa vasıtasıyla yol kotundaki rögara aktarılacaktır."
      )

    if pissu_sec_7:
      pis_su_maddeleri.append(
          "Pis su vaziyette de görüldüğü gibi rögarlar vasıtasıyla yoldan geçen"
          " pis su kanalına bağlanacaktır."
      )

    if ek_pissu_on_bilgi.strip():
      for ep in ek_pissu_on_bilgi.split("\n"):
        if ep.strip():
          pis_su_maddeleri.append(ep.strip())

    for psm in pis_su_maddeleri:
      doc.add_paragraph(psm, style="List Bullet")

    # --- 6.2.1 PİS SU SARFİYAT YÜKLEME BİRİMLERİ VE ÇAP TAYİNİ ---
    doc.add_heading("6.2.1 Pis Su Sarfiyat Yükleme Birimleri", level=2)
    doc.add_paragraph(
        "TS 826'ya göre pis su sarfiyat ve yükleme birimleri ile boru çapı"
        " tayinlerinde aşağıdaki tablolar esas alınmıştır."
    )

    doc.add_paragraph(
        "Tablo: TS 826'ya göre Pis Su Sarfiyat ve Yükleme Birimleri Cetveli"
    )
    t_pissu1 = doc.add_table(rows=len(pissu_t1_data), cols=2)
    t_pissu1.style = "Table Grid"
    for r_idx, row in enumerate(pissu_t1_data):
      for c_idx, val in enumerate(row):
        t_pissu1.cell(r_idx, c_idx).text = val

    doc.add_paragraph()
    doc.add_paragraph("Tablo: Yükleme Birimi ve Boru Çapı Esasları")
    t_pissu2 = doc.add_table(rows=len(pissu_t2_data), cols=3)
    t_pissu2.style = "Table Grid"
    for r_idx, row in enumerate(pissu_t2_data):
      for c_idx, val in enumerate(row):
        t_pissu2.cell(r_idx, c_idx).text = val

    doc.add_paragraph()
    doc.add_paragraph(
        "NOT: Lavabo yatay hatları Ø70, Lavabo inişleri Ø"
        " 50,her tuvalet çıkışı Ø100 olacaktır."
    )

    # --- 6.2.2 PİS SU TERFİ POMPALARI SEÇİM RAPORU ---
    if psp_parametreleri:
      doc.add_heading("6.2.2 PİS SU TERFİ POMPALARI SEÇİMİ", level=2)

      doc.add_paragraph(
          "Pis Su Terfi Pompası Genel Esasları ve Tasarım Kriterleri:"
      )
      terfi_maddeleri = []
      if terfi_sec_1:
        terfi_maddeleri.append(
            "Kot kurtarmayan bodrum kat atık suları için paslanmaz gövdeli,"
            " parçalayıcı bıçaklı pis su atık su terfi pompaları seçilmiştir."
        )
      if terfi_sec_2:
        terfi_maddeleri.append(
            "Pompalar yedekli çalışacak şekilde otomasyona bağlanacaktır."
        )
      if terfi_sec_3:
        terfi_maddeleri.append(
            "Terfi çukurunda sıvı seviye şalterleri (şamandıra) bulunacak, su"
            " seviyesine göre pompalar otomatik devreye girip çıkacaktır."
        )
      if terfi_sec_4:
        terfi_maddeleri.append(
            "Pompa basma hatlarında geri akışı önlemek için çekvalf ve bakım"
            " kolaylığı için sürgülü/kelebek vana kullanılacaktır."
        )

      if ek_terfi_notu.strip():
        for etn in ek_terfi_notu.split("\n"):
          if etn.strip():
            terfi_maddeleri.append(etn.strip())

      if terfi_maddeleri:
        for tm in terfi_maddeleri:
          doc.add_paragraph(tm, style="List Bullet")

      doc.add_paragraph(
          "Pis su terfi pompalarının çalışma noktaları; bina kullanım türü,"
          " armatür yükleme birimleri ve asıl pompa sayılarına göre ayrı ayrı"
          " hesaplanmıştır."
      )

      for idx, (psp, pp) in enumerate(psp_parametreleri.items(), start=1):
        doc.add_heading(f"6.2.2.{idx} {psp} TERFİ POMPASI SEÇİMİ", level=3)

        doc.add_paragraph(f"• Bina Kullanım Türü: {pp['bina_tipi']}")
        doc.add_paragraph(
            "• Seçilen Armatür Adetleri ve Yükleme Birimleri (Tablo):"
        )

        arm_tablo = doc.add_table(rows=1, cols=4)
        arm_tablo.style = "Table Grid"
        arm_tablo.rows[0].cells[0].text = "Armatür Cinsi"
        arm_tablo.rows[0].cells[1].text = "Yükleme Birimi (Y.B.)"
        arm_tablo.rows[0].cells[2].text = "Adet"
        arm_tablo.rows[0].cells[3].text = "Toplam Çarpım (Y.B.)"

        for satirlik in pp["tablo_satirlari"]:
          cinsi, birim_yb, adet_sayisi, carpim_yb = satirlik
          row_cells = arm_tablo.add_row().cells
          row_cells[0].text = cinsi
          row_cells[1].text = str(birim_yb)
          row_cells[2].text = str(adet_sayisi)
          row_cells[3].text = str(carpim_yb)

        doc.add_paragraph(
            f"• Toplam Yükleme Birimi (Y.B.) Toplamı = {pp['toplam_yb']} Y.B."
        )

        net_satir = (
            f"• Toplam Sistem Debisi Q_toplam = k * √Y.B ="
            f" {pp['k_katsayisi']} * √{pp['toplam_yb']} ="
            f" {pp['net_q_m3h']:.2f} m³/h ({pp['net_q_lps']:.2f} L/s)"
        )
        doc.add_paragraph(net_satir)

        if pp["emniyet_katsayisi"] > 1.0:
          yuzde_str = pp["emniyet_etiket"].split(" ")[0]
          emniyet_satir = (
              f"  (%{yuzde_str.replace('%','')} Emniyet Oranı Alınmıştır."
              f" Emniyetli Toplam Debi = {pp['v_toplam']:.2f} m³/h"
              f" [{pp['q_lps_toplam']:.2f} L/s])"
          )
          p_emn = doc.add_paragraph(emniyet_satir)
          p_emn.runs[0].font.italic = True

        doc.add_heading(f"-Seçilen Pompa: {psp}", level=4)
        doc.add_paragraph(
            f"V           = {pp['v_tek']:.2f} m3/h - {pp['q_lps_tek']:.2f} L/s"
        )
        doc.add_paragraph(f"H           = {pp['h']:.2f} mSS")
        doc.add_paragraph(f"Güç         = {pp['guc']:.2f} kW")
        doc.add_paragraph(f"Adet        = {pp['adet_str']}")
        doc.add_paragraph(
            f"Tip         = Dalgıç Tip, Parçalayıcı Bıçaklı, Kesme Düzenekli"
            " Pis Su Terfi Pompası"
        )
        if poz_rapora_eklensin_mi:
          doc.add_paragraph(f"Cihaz Poz No: {pp['poz']}")

        curve = pp.get("pompa_curve", [])
        if curve:
          q_curve = [p[0] for p in curve]
          h_curve = [p[1] for p in curve]
          grafik_buf = pompa_grafigi_png(
              q_curve, h_curve, pp["v_tek"], pp["h"],
              pp.get("pompa_egrisi_basligi", "Pompa Performans Eğrisi"),
              anonim=True,
          )
          doc.add_paragraph("Pompa Performans Eğrisi:")
          doc.add_picture(grafik_buf, width=Inches(6.2))
          doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- 6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ ---
    doc.add_heading("6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ", level=1)
    doc.add_heading("6.3.1 KULLANMA SOĞUK SUYU DEPOSU SEÇİMİ", level=2)
    doc.add_heading("Genel Bilgiler", level=3)

    # Dinamik Depo Tipi Metni Oluşturma (6.1'deki seçime bağlı)
    if sih_sec_depo_tipi and sih_depo_tipleri:
      if len(sih_depo_tipleri) == 1:
        dinamik_tip_str = sih_depo_tipleri[0].lower()
      elif len(sih_depo_tipleri) == 2:
        dinamik_tip_str = (
            f"{sih_depo_tipleri[0].lower()} ve {sih_depo_tipleri[1].lower()}"
        )
      else:
        ilkler = ", ".join([t.lower() for t in sih_depo_tipleri[:-1]])
        son = sih_depo_tipleri[-1].lower()
        dinamik_tip_str = f"{ilkler} ve {son}"
    else:
      dinamik_tip_str = "modüler su deposu"

    depo_maddeleri = []

    # Dinamik Cümle
    depo_maddeleri.append(
        "Binanın kullanma soğuk suyu ihtiyacının karşılanması ve kesintilere"
        f" karşı güvence altına alınması amacıyla {dinamik_tip_str}"
        " tasarlanmıştır."
    )

    if depo_sec_1:
      depo_maddeleri.append(
          "Kullanma soğuk suyu deposu hacmi; binanın kullanım amacı, kullanıcı "
          "sayısı, kişi başına günlük su tüketimi, kullanım sürekliliği ve ihtiyaç "
          "duyulan su rezervi dikkate alınarak belirlenecektir. Depo kapasitesi, "
          "binanın günlük su ihtiyacını karşılayacak ve işletme koşullarında yeterli "
          "su rezervi sağlayacak şekilde tasarlanacaktır."
      )
    if depo_sec_2:
      depo_maddeleri.append(
          "Su deposu içerisinde su kalitesinin korunması ve ölü hacim oluşumunun"
          " önlenmesi için bölme perdeleri yer alacaktır."
      )
    if depo_sec_3:
      depo_maddeleri.append(
          "Su deposunda taşma, deşarj, havalandırma boruları ile bakım ve temizlik"
          " için adam geçiş kapağı (manhole) bulunacaktır."
      )

    if ek_depo_notu.strip():
      for ed in ek_depo_notu.split("\n"):
        if ed.strip():
          depo_maddeleri.append(ed.strip())

    # Seçili depo notları Genel Bilgiler başlığı altında gösterilir.
    for dm in depo_maddeleri:
      doc.add_paragraph(dm, style="List Bullet")

    # Su ihtiyacı hesabı, seçili depo notlarından sonra ayrı alt başlık olarak verilir.
    doc.add_heading(
        "6.3.1.1 KULLANMA SUYU İHTİYACININ BELİRLENMESİ",
        level=3,
    )
    doc.add_heading("Su Tüketim Değerleri Tablosu", level=4)
    su_tuketim_word_tablosu = doc.add_table(rows=1, cols=3)
    su_tuketim_word_tablosu.style = "Table Grid"
    baslik_hucreleri = su_tuketim_word_tablosu.rows[0].cells
    baslik_hucreleri[0].text = "Kullanım amacı"
    baslik_hucreleri[1].text = "Birim"
    baslik_hucreleri[2].text = "Birim tüketim değeri"
    for kategori, (birim, deger) in su_tuketim_secenekleri.items():
        hucreler = su_tuketim_word_tablosu.add_row().cells
        hucreler[0].text = kategori
        hucreler[1].text = birim
        hucreler[2].text = f"{deger:g} L/{birim}/gün"

    doc.add_paragraph(
        "Not: Birim tüketim değerleri, kullanıcı tarafından yüklenen "
        "dokümanda belirtilen TS-1258 kaynaklı değerler esas alınarak "
        "uygulamaya aktarılmıştır."
    )

    doc.add_heading("Su İhtiyacı Hesabı", level=4)
    doc.add_paragraph(
        "Kullanım amacı / su tüketim kategorileri: "
        f"{su_tuketim_tipi}"
    )
    if su_hesap_detaylari:
        hesap_tablosu = doc.add_table(rows=1, cols=5)
        hesap_tablosu.style = "Table Grid"
        hesap_basliklari = hesap_tablosu.rows[0].cells
        hesap_basliklari[0].text = "Kategori"
        hesap_basliklari[1].text = "Miktar"
        hesap_basliklari[2].text = "Birim"
        hesap_basliklari[3].text = "Birim tüketimi"
        hesap_basliklari[4].text = "Günlük ihtiyaç"
        for detay in su_hesap_detaylari:
            hucreler = hesap_tablosu.add_row().cells
            hucreler[0].text = detay["kategori"]
            hucreler[1].text = f"{detay['miktar']:g}"
            hucreler[2].text = detay["birim"]
            hucreler[3].text = f"{detay['birim_degeri']:g} L/{detay['birim']}/gün"
            hucreler[4].text = f"{detay['ihtiyac_litre']:g} L/gün"
    else:
        doc.add_paragraph("Herhangi bir su tüketim kategorisi seçilmemiştir.")
    doc.add_paragraph(
        f"Günlük toplam su ihtiyacı: {su_gunluk_ihtiyac_litre:g} L/gün "
        f"({su_gunluk_ihtiyac_m3:g} m³/gün)"
    )
    doc.add_heading("Su Deposu Depolama Süresi ve Gerekli Hacim", level=4)
    doc.add_paragraph(f"Seçilen depolama süresi: {depo_sure_gun:g} gün")
    depo_hacmi_paragrafi = doc.add_paragraph()
    depo_hacmi_paragrafi.add_run(
        f'Yapının kullanım soğuk suyu ihtiyacını karşılamak için seçilen "{secilen_depo_tipi_metni}" hacmi: '
        if secilen_depo_tipi_metni
        else "Yapının kullanım soğuk suyu ihtiyacını karşılamak için seçilen su deposu hacmi: "
    )
    rapor_kapasite_m3 = otomatik_poz_kayitlari[0][1] if otomatik_poz_kayitlari else depo_gerekli_hacim_m3
    rapor_kapasite_litre = rapor_kapasite_m3 * 1000.0
    depo_hacmi_kalin = depo_hacmi_paragrafi.add_run(
        f"{rapor_kapasite_litre:g} L ({rapor_kapasite_m3:g} m³)"
    )
    depo_hacmi_kalin.bold = True
    depo_hacmi_paragrafi.add_run("'dir.")
    for run in depo_hacmi_paragrafi.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)
    # Word raporunda yalnızca otomatik/elle düzenlenmiş poz numarası gösterilir.
    # Kapasite bilgisi rapora yazılmaz; kapasite yalnızca program ekranında gösterilir.
    if poz_gosterilsin_mi and poz_numarasi:
        poz_paragrafi = doc.add_paragraph()
        poz_paragrafi.add_run("Cihaz Poz No: ")
        poz_kalin = poz_paragrafi.add_run(poz_numarasi.strip())
        poz_kalin.bold = True
        for run in poz_paragrafi.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)
    rapor_word_stillerini_uygula(doc)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("6.3.1 Kullanma Soğuk Suyu Deposu raporu başarıyla hazırlandı!")

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
