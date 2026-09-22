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
# 6.1.1 TEMİZ SU SARFİYAT YÜKLEME BİRİMLERİ VE ÇAP TAYİNİ (ORİJİNAL TABLOLAR)
# ---------------------------------------------------------------------------
st.subheader("6.1.1 Temiz Su Sarfiyat Yükleme Birimleri ve Çap Tayini")
st.write(
    "Sıhhi tesisat boru çaplarının tespitinde ve kullanım yerlerine ait yükleme"
    " birimleri ile debi değerlerinde aşağıdaki tablolar esas alınmıştır."
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

t2_data = [
    ("KULLANMA YERİ", "DEBİ (lt/sn)", "YÜKLEME BİRİMİ"),
    ("Küvetli Banyo (DN15 mm)", "0.40", "2.50"),
    ("Banyo, Jakuzili (DN15)", "1.00", "16.00"),
    ("Bide Rezervuarı", "0.13", "0.25"),
    ("Bulaşık Makinası", "0.40", "2.50"),
    ("Çamaşır Makinası", "0.40", "2.50"),
    ("Duş", "0.40", "2.50"),
    ("1 Gözlü Eviye", "0.25", "1.00"),
    ("2 Gözlü Eviye", "0.31", "1.50"),
    ("Hela Rezervuarı", "0.13", "0.25"),
    ("Basınçlı Hela Yıkayıcısı (DN15 mm)", "0.61", "6.00"),
    ("Basınçlı Hela Yıkayıcısı (DN20 mm)", "0.83", "11.00"),
    ("Basınçlı Hela Yıkayıcısı (DN25 mm)", "1.30", "27.00"),
    ("Kurna", "0.40", "2.50"),
    ("Lavabo", "0.18", "0.50"),
    ("DN15 mm musluk", "0.31", "1.50"),
    ("DN20 mm musluk", "0.71", "8.00"),
    ("DN25 mm musluk", "1.06", "18.00"),
    ("Pisuvar", "0.13", "0.25"),
    ("Şofben (10lt/dk)", "0.18", "0.50"),
    ("Şofben (16lt/dk)", "0.25", "1.00"),
    ("Şofben (26lt/dk)", "0.43", "3.00"),
    ("Taharet Musluğu", "0.13", "0.25"),
    ("Termosifon", "0.40", "2.50"),
]


# ---------------------------------------------------------------------------
# 6.2.1 PİS SU SARFİYAT YÜKLEME BİRİMLERİ VE ÇAP TAYİNİ (TS 826 TABLOLARI)
# ---------------------------------------------------------------------------
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


def pompa_secim_egrisi(q_m3h, h_mss):
  poz, _, durum = pompa_pozu_sec(q_m3h, h_mss)
  if durum == "UYGUN":
    kayit = next(x for x in POZ_POMPA_TABLOSU if x["poz"] == poz)
    q0, q1 = kayit["qmin"], kayit["qmax"]
    h0, h1 = kayit["hmax"], kayit["hmin"]
    q_egrisi = [q0 + (q1 - q0) * i / 100 for i in range(101)]
    h_egrisi = [h0 + (h1 - h0) * ((q - q0) / (q1 - q0)) for q in q_egrisi]
    baslik = f"Poz {poz} Sınır Zarfı (Uygun)"
  else:
    q_egrisi = [2, 22]
    h_egrisi = [22, 2]
    baslik = "Sınır Dışı Çalışma Noktası!"
  return q_egrisi, h_egrisi, baslik


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

if secilen_psp_listesi:
  st.write(
      "Her bir terfi pompası çukuru için bina tipi, armatürler, k katsayısı ve"
      " asıl/yedek adetlerini ayrı ayrı girin:"
  )

  for psp in secilen_psp_listesi:
    with st.expander(
        f"⚙️ {psp} Özel Debi ve Güç Hesap Modülü", expanded=True
    ):
      sc1, sc2 = st.columns(2)
      with sc1:
        bina_tipi = st.selectbox(
            f"{psp} Bina Kullanım Türü",
            [
                "Konut / Ev",
                "İş Yeri / Ofis",
                "Otel / Konaklama",
                "Hastane / Sağlık",
                "Okul / Eğitim",
            ],
            key=f"{psp}_bina",
        )
      with sc2:
        k_varsayilan = (
            0.5 if bina_tipi in ["Konut / Ev", "Okul / Eğitim"] else 0.7
        )
        k_katsayisi = st.number_input(
            f"{psp} Eşzamanlık Katsayısı (k)",
            min_value=0.1,
            max_value=1.0,
            value=k_varsayilan,
            step=0.05,
            key=f"{psp}_k",
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

      ac4, ac5 = st.columns(2)
      with ac4:
        adet_basincli = st.number_input(
            f"{psp} Basınçlı Yıkayıcı (10 Y.B.)",
            min_value=0,
            value=0,
            key=f"{psp}_bas",
        )
        adet_camasir = st.number_input(
            f"{psp} Çamaşır/Bulaşık (10 Y.B.)",
            min_value=0,
            value=1,
            key=f"{psp}_cam",
        )
      with ac5:
        adet_pisuvar = st.number_input(
            f"{psp} Pisuar (1 Y.B.)", min_value=0, value=0, key=f"{psp}_pis"
        )
        h_val = st.number_input(
            f"{psp} Basma Yüksekliği (H) [mSS]",
            min_value=1.0,
            max_value=30.0,
            value=12.0,
            step=0.5,
            key=f"{psp}_h_num",
        )

      bc1, bc2 = st.columns(2)
      with bc1:
        asil_adedi = st.selectbox(
            f"{psp} Asıl Pompa Adedi", [1, 2, 3], index=0, key=f"{psp}_asil"
        )
      with bc2:
        yedek_adedi = st.selectbox(
            f"{psp} Yedek Pompa Adedi", [1, 2], index=0, key=f"{psp}_yedek"
        )

      # Yükleme Birimi ve Eşzamanlı Debi Formülasyonu
      toplam_yb = (
          (adet_hela * 8)
          + (adet_lavabo * 2)
          + (adet_banyo * 7)
          + (adet_evye * 4)
          + (adet_suzgec * 2)
          + (adet_otopark * 6)
          + (adet_basincli * 10)
          + (adet_camasir * 10)
          + (adet_pisuvar * 1)
      )

      hesaplanan_q_lps = (
          k_katsayisi * math.sqrt(toplam_yb) if toplam_yb > 0 else 0.0
      )
      hesaplanan_q_m3h = round(hesaplanan_q_lps * 3.6, 2)

      # Kullanıcının otomatik gelen debiyi isterse manuel düzenleyebilmesi/yuvarlayabilmesi için number_input
      toplam_v_val = st.number_input(
          f"{psp} Toplam Debi (Q_toplam) [m³/h] (Hesaplanan: {hesaplanan_q_m3h})",
          min_value=1.0,
          max_value=100.0,
          value=max(1.0, hesaplanan_q_m3h),
          step=0.5,
          key=f"{psp}_v_num",
      )

      # Pompa başına düşen debi (Asıl pompa sayısına bölünür)
      pompa_basina_v = toplam_v_val / asil_adedi

      hesap = pompa_hidrolik_hesap(pompa_basina_v, h_val, 0.60, 0.90)
      hesaplanan_poz, hesaplanan_tanim, poz_durumu = pompa_pozu_sec(
          pompa_basina_v, h_val
      )
      q_egrisi, h_egrisi, egrisi_basligi = pompa_secim_egrisi(
          pompa_basina_v, h_val
      )

      with c2:
        st.metric(
            "Tek Pompa Debisi (Asıla Bölünen)", f"{pompa_basina_v:.2f} m³/h"
        )
        st.metric(
            "Tek Pompa Elektrik Gücü", f"{hesap['motor_secim_kw']:.2f} kW"
        )

        if poz_durumu == "UYGUN":
          st.success(
              f"✅ Uygun Poz: **{hesaplanan_poz}** (Eğri Sınırları İçinde)"
          )
        else:
          st.error(
              "❌ HATA: Tek pompa debi/basıncı 25.360.1301–1308 sınırları"
              " dışındadır! Asıl pompa sayısını artırın."
          )

      st.info(f"📌 **Poz Tanımı:** {hesaplanan_tanim}")

      # Yalnızca program içinde görünen seçim eğrisi grafiği
      fig, ax = plt.subplots(figsize=(6, 3))
      ax.plot(q_egrisi, h_egrisi, label=egrisi_basligi, color="blue")
      p_renk = "green" if poz_durumu == "UYGUN" else "red"
      ax.scatter(
          [pompa_basina_v],
          [h_val],
          s=70,
          color=p_renk,
          zorder=5,
          label=f"Çalışma Noktası ({pompa_basina_v:.1f} m³/h, {h_val} mSS)",
      )
      ax.set_xlim(2, 22)
      ax.set_ylim(2, 22)
      ax.set_xlabel("Debi Q [m³/h]")
      ax.set_ylabel("Basma Yüksekliği H [mSS]")
      ax.set_title(
          f"{psp} Program İçi Pompa Seçim ve Sınır Kontrolü", fontsize=9
      )
      ax.grid(True, alpha=0.3)
      ax.legend(fontsize=7)
      fig.tight_layout()
      st.pyplot(fig)
      plt.close(fig)

      toplam_adet = asil_adedi + yedek_adedi
      adet_metin = (
          f"{toplam_adet} ({asil_adedi} Asıl"
          f" {'+ ' + str(yedek_adedi) + ' Yedek' if yedek_adedi > 0 else ''})"
      )
      tip_metin = (
          "Dalgıç Tip, Parçalayıcı Bıçaklı, Kesme Düzenekli Pis Su Terfi"
          " Pompası"
      )

      # Seçilen armatür listesini kaydetme
      armatur_listesi = []
      if adet_hela > 0:
        armatur_listesi.append(f"Hela / Klozet: {adet_hela} Adet")
      if adet_lavabo > 0:
        armatur_listesi.append(f"Lavabo / Bide: {adet_lavabo} Adet")
      if adet_banyo > 0:
        armatur_listesi.append(f"Küvet / Duş: {adet_banyo} Adet")
      if adet_evye > 0:
        armatur_listesi.append(f"Eviye: {adet_evye} Adet")
      if adet_suzgec > 0:
        armatur_listesi.append(f"Yer Süzgeci: {adet_suzgec} Adet")
      if adet_otopark > 0:
        armatur_listesi.append(f"Otopark Süzgeci: {adet_otopark} Adet")
      if adet_basincli > 0:
        armatur_listesi.append(f"Basınçlı Yıkayıcı: {adet_basincli} Adet")
      if adet_camasir > 0:
        armatur_listesi.append(f"Çamaşır/Bulaşık Makinası: {adet_camasir} Adet")
      if adet_pisuvar > 0:
        armatur_listesi.append(f"Pisuar: {adet_pisuvar} Adet")

      psp_parametreleri[psp] = {
          "bina_tipi": bina_tipi,
          "k_katsayisi": k_katsayisi,
          "toplam_yb": toplam_yb,
          "q_lps_toplam": hesaplanan_q_lps,
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
          "armatur_listesi": armatur_listesi,
      }

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


# Rapor Oluştur Butonu
if st.button("Raporu Oluştur (.docx)"):
  gecersiz_var = any(
      p["poz_durumu"] != "UYGUN" for p in psp_parametreleri.values()
  )
  if gecersiz_var:
    st.error(
        "❌ Rapor oluşturulamadı! Seçilen pompalardan biri veya daha fazlasının"
        " tek pompa debisi poz sınırları dışındadır. Lütfen asıl pompa sayısını"
        " veya debiyi kontrol edin."
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
    # 3. SAYFA: GÖVDE VE RAPOR İÇERİĞİ
    # ==========================================
    doc.add_page_break()

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

    if ek_pissu_on_bilgi.strip():
      for es in ek_pissu_on_bilgi.split("\n"):
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
    doc.add_heading(
        "6.2.1 Pis Su Sarfiyat Yükleme Birimleri ve Çap Tayini", level=2
    )
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
          "Pis su terfi pompalarının çalışma noktaları; bina kullanım türü,"
          " armatür yükleme birimleri ve asıl pompa sayılarına göre ayrı ayrı"
          " hesaplanmıştır."
      )

      for idx, (psp, pp) in enumerate(psp_parametreleri.items(), start=1):
        doc.add_heading(
            f"6.2.2.{idx} {psp} PİS SU TERFİ POMPA SEÇİMİ", level=3
        )

        doc.add_paragraph(
            f"• Bina Kullanım Türü: {pp['bina_tipi']} (Eşzamanlık katsayısı k ="
            f" {pp['k_katsayisi']})"
        )
        doc.add_paragraph("• Seçilen Armatür Adetleri ve Yükleme Birimleri:")
        for arm in pp["armatur_listesi"]:
          doc.add_paragraph(f"  - {arm}", style="List Bullet")
        doc.add_paragraph(
            f"• Toplam Yükleme Birimi (Y.B.) = {pp['toplam_yb']} Y.B."
        )
        doc.add_paragraph(
            f"• Toplam Sistem Debisi Q_toplam = k * √Y.B. ="
            f" {pp['q_lps_toplam']:.2f} L/s ({pp['v_toplam']:.2f} m³/h)"
        )

        doc.add_paragraph(
            f"V           = {pp['v_tek']:.2f} m3/h - {pp['q_lps_tek']:.2f} L/s"
        )
        doc.add_paragraph(f"H           = {pp['h']:.2f} mSS")
        doc.add_paragraph(f"Güç         = {pp['guc']:.2f} kW")
        doc.add_paragraph(f"Adet        = {pp['adet_str']}")
        doc.add_paragraph(f"Tip         = {pp['tip']}")
        if poz_rapora_eklensin_mi:
          doc.add_paragraph(f"Cihaz Poz No: {pp['poz']}")

    # Hafızada dosya oluşturma
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("Armatür listeleri ve hesap formülleri eklenerek rapor hazırlandı!")

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
