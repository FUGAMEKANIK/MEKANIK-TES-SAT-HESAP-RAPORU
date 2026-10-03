from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from datetime import datetime
import io
import json
import math
import os
import re
import subprocess
import tempfile
import shutil
import urllib.request
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from proje_yonetimi import ProjeYoneticisi, guvenli_dosya_adi


# ---------------------------------------------------------------------------
# MGM YAĞIŞ VERİSİ
# ---------------------------------------------------------------------------
@st.cache_data(ttl=86400, show_spinner=False)
def mgm_gunluk_en_yuksek_yagis_mm(il_adi):
    """MGM Resmi İklim İstatistikleri sayfasından ilin günlük toplam en
    yüksek yağış miktarını (mm) ve gerçekleşme tarihini çeker.

    MGM'nin bazı il sayfalarında bağlantı/HTML yapısı zaman zaman farklı
    dönebildiği için birden fazla URL biçimi, yeniden deneme ve iki farklı
    metin deseni kullanılır. Böylece rapordaki "Veri alınamadı" sayısı
    mümkün olduğunca azaltılır.
    """
    import re as _re
    import unicodedata as _unicodedata
    import html as _html_lib
    import time as _time

    il_adi = str(il_adi or "").strip()
    if not il_adi:
        return None, None, None

    _il_url = "".join(
        ch for ch in _unicodedata.normalize("NFKD", il_adi)
        if not _unicodedata.combining(ch)
    ).upper()
    _il_url = (
        _il_url.replace("Ç", "C").replace("Ğ", "G")
        .replace("İ", "I").replace("Ö", "O")
        .replace("Ş", "S").replace("Ü", "U")
    )

    # MGM'de aynı il istatistiğine ulaşabilen alternatif URL biçimleri.
    _base1 = "https://www.mgm.gov.tr/veridegerlendirme/il-ve-ilceler-istatistik.aspx"
    _base2 = "https://www.mgm.gov.tr/Veridegerlendirme/Il-Ve-Ilceler-Istatistik.Aspx"
    _urls = [
        f"{_base1}?k=undefined&m={_il_url}",
        f"{_base1}?m={_il_url}",
        f"{_base2}?m={_il_url}",
    ]

    _baslik = "Günlük Toplam En Yüksek Yağış Miktarı"
    _headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/140.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.7",
        "Connection": "close",
    }

    def _parse_html(_html, _url):
        # HTML entity'lerini gerçek karakterlere çevir; bazı MGM sayfalarında
        # başlık/değerler entity olarak gelebiliyor.
        _metin = _html_lib.unescape(_html or "")
        _metin = _re.sub(r"<script\b[^>]*>.*?</script>", " ", _metin, flags=_re.I | _re.S)
        _metin = _re.sub(r"<style\b[^>]*>.*?</style>", " ", _metin, flags=_re.I | _re.S)
        _metin = _re.sub(r"<[^>]+>", " ", _metin)
        _metin = _re.sub(r"\s+", " ", _metin).strip()

        _pos = _metin.lower().find(_baslik.lower())
        if _pos < 0:
            return None, None, _url

        _parca = _metin[_pos:_pos + 1200]
        _desenler = [
            # Standart MGM görünümü: 11.06.1997 88,9 mm
            r"(\d{1,2}\.\d{1,2}\.\d{4})\s*[:|]?\s*([0-9]+(?:[.,][0-9]+)?)\s*mm\b",
            # Bazı HTML varyasyonlarında değer önce gelebilir.
            r"([0-9]+(?:[.,][0-9]+)?)\s*mm\b\s*[:|]?\s*(\d{1,2}\.\d{1,2}\.\d{4})",
        ]
        for _i, _desen in enumerate(_desenler):
            _eslesme = _re.search(_desen, _parca, flags=_re.I)
            if not _eslesme:
                continue
            if _i == 0:
                _tarih = _eslesme.group(1)
                _deger = float(_eslesme.group(2).replace(",", "."))
            else:
                _deger = float(_eslesme.group(1).replace(",", "."))
                _tarih = _eslesme.group(2)
            return _deger, _tarih, _url
        return None, None, _url

    _son_url = _urls[0]
    for _url in _urls:
        _son_url = _url
        for _deneme in range(3):
            try:
                _req = urllib.request.Request(_url, headers=_headers)
                with urllib.request.urlopen(_req, timeout=20) as _response:
                    _html = _response.read().decode("utf-8", errors="ignore")
                _deger, _tarih, _parsed_url = _parse_html(_html, _url)
                if _deger is not None:
                    return _deger, _tarih, _parsed_url
            except Exception:
                pass
            if _deneme < 2:
                _time.sleep(0.7 * (_deneme + 1))

    return None, None, _son_url


# MGM'nin il seçimindeki resmi 81 il listesi.
# Rapor tablosu yalnızca Word raporu oluşturulurken kullanılır; arayüzde gösterilmez.
MGM_81_IL = [
    "Adana", "Adıyaman", "Afyonkarahisar", "Ağrı", "Aksaray", "Amasya",
    "Ankara", "Antalya", "Ardahan", "Artvin", "Aydın", "Balıkesir",
    "Bartın", "Batman", "Bayburt", "Bilecik", "Bingöl", "Bitlis",
    "Bolu", "Burdur", "Bursa", "Çanakkale", "Çankırı", "Çorum",
    "Denizli", "Diyarbakır", "Düzce", "Edirne", "Elazığ", "Erzincan",
    "Erzurum", "Eskişehir", "Gaziantep", "Giresun", "Gümüşhane",
    "Hakkari", "Hatay", "Iğdır", "Isparta", "İstanbul", "İzmir",
    "Kahramanmaraş", "Karabük", "Karaman", "Kars", "Kastamonu",
    "Kayseri", "Kırıkkale", "Kırklareli", "Kırşehir", "Kilis", "Kocaeli",
    "Konya", "Kütahya", "Malatya", "Manisa", "Mardin", "Mersin",
    "Muğla", "Muş", "Nevşehir", "Niğde", "Ordu", "Osmaniye", "Rize",
    "Sakarya", "Samsun", "Siirt", "Sinop", "Sivas", "Şanlıurfa",
    "Şırnak", "Tekirdağ", "Tokat", "Trabzon", "Tunceli", "Uşak", "Van",
    "Yalova", "Yozgat", "Zonguldak",
]

# MGM URL parametresinde kullanılan ve il adından farklı olan özel kodlar.
MGM_IL_URL_KODU = {"Mersin": "ICEL"}


@st.cache_data(ttl=86400, show_spinner=False)
def mgm_81_il_yagis_tablosu():
    """81 ilin MGM günlük maksimum yağış verisini rapor için toplar.

    Sonuç Word raporunda tablo olarak kullanılır; Streamlit arayüzüne
    herhangi bir tablo basılmaz. Paralel istekler rapor oluşturma süresini
    kısaltır.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    def _tek_il(il):
        _kod = MGM_IL_URL_KODU.get(il, il)
        deger, tarih, url = mgm_gunluk_en_yuksek_yagis_mm(_kod)
        # Yardımcı fonksiyon özel kodla çağrıldığında mgm_il alanının
        # raporda gerçek il adı olarak görünmesi için gerçek URL'yi üret.
        if _kod != il:
            url = (
                "https://www.mgm.gov.tr/veridegerlendirme/il-ve-ilceler-istatistik.aspx"
                f"?k=undefined&m={_kod}"
            )
        return il, deger, tarih, url

    sonuc = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        gelecekler = {executor.submit(_tek_il, il): il for il in MGM_81_IL}
        for gelecek in as_completed(gelecekler):
            il = gelecekler[gelecek]
            try:
                sonuc.append(gelecek.result())
            except Exception:
                try:
                    _kod = MGM_IL_URL_KODU.get(il, il)
                    _d, _t, _u = mgm_gunluk_en_yuksek_yagis_mm(_kod)
                    sonuc.append((il, _d, _t, _u))
                except Exception:
                    sonuc.append((il, None, None, None))

    sirali = {il: (deger, tarih, url) for il, deger, tarih, url in sonuc}
    return [
        (il, *sirali.get(il, (None, None, None)))
        for il in MGM_81_IL
    ]


# ---------------------------------------------------------------------------
# PROGRAM BASLIKLARI - RENKLI VE DUZENLI GORUNUM
# ---------------------------------------------------------------------------
st.markdown('<style>\n#sayfa_basi, #sayfa_sonu { scroll-margin-top: 20px; }\n.hizli-navigasyon {\n    position: fixed;\n    right: 18px;\n    top: 50%;\n    transform: translateY(-50%);\n    z-index: 999999;\n    display: flex;\n    flex-direction: column;\n    gap: 7px;\n}\n.hizli-navigasyon a {\n    display: block;\n    min-width: 108px;\n    padding: 8px 11px;\n    text-align: center;\n    text-decoration: none !important;\n    border-radius: 8px;\n    border: 1px solid #B8C7D9;\n    background: rgba(255,255,255,0.96);\n    color: #0B3D91 !important;\n    font-weight: 700;\n    font-size: 13px;\n    box-shadow: 0 2px 8px rgba(0,0,0,0.16);\n}\n.hizli-navigasyon a:hover { background: #EAF2F8; }\n@media (max-width: 700px) {\n    .hizli-navigasyon { right: 8px; top: 50%; transform: translateY(-50%); }\n    .hizli-navigasyon a { min-width: 92px; padding: 7px 8px; font-size: 12px; }\n}\n</style>', unsafe_allow_html=True)

st.markdown("""
<style>
/* Ana bolum basliklari */
h1 {
    color: #0B3D91 !important;
    font-weight: 800 !important;
    border-bottom: 3px solid #0B3D91;
    padding-bottom: 8px;
    margin-top: 18px !important;
}

/* 2. seviye basliklar: ana bolumler */
h2 {
    color: #0B3D91 !important;
    background: #EAF2F8;
    border-left: 7px solid #0B3D91;
    border-radius: 6px;
    padding: 9px 14px !important;
    font-weight: 750 !important;
    margin-top: 18px !important;
    margin-bottom: 10px !important;
}

/* 3. seviye basliklar: alt bolumler */
h3 {
    color: #0F5B78 !important;
    background: #F2F8FA;
    border-left: 5px solid #1B8AAA;
    border-radius: 5px;
    padding: 7px 12px !important;
    font-weight: 700 !important;
    margin-top: 13px !important;
    margin-bottom: 8px !important;
}

/* 4. seviye basliklar: hesap/alt basliklar */
h4 {
    color: #7A4E00 !important;
    font-weight: 700 !important;
    border-bottom: 1px solid #E6C77A;
    padding-bottom: 4px !important;
}

/* Streamlit st.header/st.subheader icin daha temiz dikey bosluk */
div[data-testid="stHeading"] h1,
div[data-testid="stHeading"] h2,
div[data-testid="stHeading"] h3,
div[data-testid="stHeading"] h4 {
    letter-spacing: 0.1px;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    """
    <div id="sayfa_basi"></div>
    <div class="hizli-navigasyon">
        <a href="#sayfa_basi">⬆ Başa Git</a>
        <a href="#sayfa_sonu">⬇ Sona Git</a>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# PROJE YÖNETİMİ — KAYDET / FARKLI KAYDET / PROJE AÇ / SİL
# ---------------------------------------------------------------------------
_PROJE_YONETICISI = ProjeYoneticisi()
_PROJE_META_KEYS = {
    "aktif_proje_id", "aktif_proje_adi", "proje_yukleme_bildirimi",
    "proje_dis_dosya_hazirlandi", "proje_dis_dosya_yukle_bekliyor",
    "proje_son_kayit_zamani", "proje_kaynak", "_proje_adi_manuel",
}
_PROJE_WIDGET_KAYDETME_DISI_KEYS = {
    # Proje yönetimi widgetları
    "proje_kaydet_btn_v115", "proje_farkli_kaydet_btn_v115",
    "proje_ac_btn_v115", "proje_dis_dosya_ac_btn_v115",
    "proje_kaydet_btn_v116", "proje_ac_btn_v116",
    "proje_sil_btn_v116", "proje_basliktan_al_btn_v116",
    "proje_dis_dosya_indir_v116", "proje_dis_dosya_ac_btn_v116",
    "proje_ac_sec_v116", "proje_dis_dosya_yukle_v116",
    "proje_yeni_btn_v119", "yeni_proje_btn_v118",
    # Ana bölüm toplu seçim / rapor üretim butonları
    "rapor_tumunu_sec_v92", "rapor_tumunu_kaldir_v92",
    "rapor_word_indir_v101",
}


def _proje_widget_anahtari_mi(anahtar):
    """Bir session-state anahtarının Streamlit tarafından yönetilen geçici bir
    buton/aksiyon widget'ı olup olmadığını belirler.

    Eski proje dosyalarında buton anahtarları bulunabilir. Bunları yüklemeye
    çalışmak StreamlitValueAssignmentNotAllowedError üretir.
    """
    anahtar = str(anahtar)
    if anahtar in _PROJE_WIDGET_KAYDETME_DISI_KEYS:
        return True
    if "_btn_" in anahtar or anahtar.endswith("_btn"):
        return True
    if anahtar.startswith("rapor_tumunu_") or anahtar.startswith("rapor_word_indir"):
        return True

    # Eski proje dosyalarında bazı butonlar "_btn_" içermeyen anahtarlarla
    # kaydedilmiş olabilir. Bunların hiçbirini projeye geri yüklemiyoruz.
    _alt = anahtar.lower()
    _buton_anahtar_parcalari = (
        "sidebar_rapor_olustur",
        "rapor_olustur",
        "proje_yeni",
        "proje_kaydet_btn",
        "proje_farkli_kaydet",
        "proje_ac_btn",
        "proje_sil_btn",
        "proje_dis_dosya_ac_btn",
        "proje_dis_dosya_indir",
        "rapor_pdf_indir",
        "rapor_word_indir",
        "rapor_html_indir",
        "rapor_txt_indir",
        "rapor_indir",
    )
    if any(parca in _alt for parca in _buton_anahtar_parcalari):
        return True

    # Genel güvenlik: Streamlit aksiyon widget'ı isimlerinde sık kullanılan
    # buton/aksiyon ifadelerini proje verisi olarak geri yükleme.
    if _alt.endswith("_button") or "_button_" in _alt:
        return True
    # _toplu_secim_butonlari() dinamik olarak oluşturduğu st.button anahtarları
    # eski proje kayıtlarında bulunabilir. Bunlar session_state'e geri yazılırsa
    # StreamlitValueAssignmentNotAllowedError oluşur.
    if anahtar.startswith("alt_grup_sec_v55_") or anahtar.startswith("alt_grup_kaldir_v55_"):
        return True
    # data_editor widget anahtarları Streamlit tarafından yönetilir.
    if "editor" in anahtar.lower():
        return True

    # Streamlit st.button durumları projeye ait veri değildir. Bazı butonlarda
    # anahtar dinamik oluşturulduğu için "_btn_" filtresi tek başına yetmez.
    _alt = anahtar.lower()
    _button_prefixleri = (
        "boyler_raporda_kullan_",
        "rapor_sec_",
        "rapor_kullan_",
        "sistem_raporda_",
    )
    if _alt.startswith(_button_prefixleri):
        return True

    # İsimlendirmesi açıkça geçici seçim/aksiyon durumunu belirten anahtarlar.
    if _alt.endswith("_sec_key") or _alt.endswith("_button_key"):
        return True

    return False


def _proje_degerini_temizle(deger):
    """Session State içindeki değerleri proje JSON'una güvenli biçimde dönüştürür."""
    if deger is None or isinstance(deger, (str, int, float, bool)):
        return deger
    if isinstance(deger, pd.DataFrame):
        return {
            "__tip__": "pandas_dataframe",
            "columns": [str(c) for c in deger.columns],
            "records": deger.where(pd.notna(deger), None).to_dict(orient="records"),
        }
    if isinstance(deger, (list, tuple)):
        sonuc = []
        for x in deger:
            temiz = _proje_degerini_temizle(x)
            if temiz is not _GEcersiz_PROJE:
                sonuc.append(temiz)
        return sonuc
    if isinstance(deger, dict):
        sonuc = {}
        for k, v in deger.items():
            temiz = _proje_degerini_temizle(v)
            if temiz is not _GEcersiz_PROJE:
                sonuc[str(k)] = temiz
        return sonuc
    return _GEcersiz_PROJE


_GEcersiz_PROJE = object()


def _proje_ayarlarini_topla():
    ayarlar = {}
    for anahtar, deger in st.session_state.items():
        if (
            anahtar in _PROJE_META_KEYS
            or _proje_widget_anahtari_mi(anahtar)
            or anahtar.startswith("proje_")
        ):
            continue
        temiz = _proje_degerini_temizle(deger)
        if temiz is not _GEcersiz_PROJE:
            ayarlar[anahtar] = temiz
    return ayarlar


def _proje_degerini_geri_yukle(deger):
    """Proje JSON'undaki özel veri tiplerini uygulama değerlerine dönüştürür."""
    if isinstance(deger, dict) and deger.get("__tip__") == "pandas_dataframe":
        return pd.DataFrame(deger.get("records", []), columns=deger.get("columns"))
    if isinstance(deger, dict):
        return {str(k): _proje_degerini_geri_yukle(v) for k, v in deger.items()}
    if isinstance(deger, list):
        return [_proje_degerini_geri_yukle(v) for v in deger]
    return deger


def _proje_adi():
    """Kullanıcının proje adı alanındaki adı; boşsa kapak başlığını döndürür."""
    ad = str(st.session_state.get("proje_adi_giris", "")).strip()
    if not ad:
        ad = str(st.session_state.get("is_adi", "")).strip()
    return ad


def _proje_kisa_dosya_adi(proje_adi):
    """Dosya/klasör adında yalnızca proje adının ilk iki kelimesini kullanır."""
    kelimeler = [k for k in str(proje_adi or "").strip().split() if k]
    if not kelimeler:
        return ""
    return guvenli_dosya_adi(" ".join(kelimeler[:2]))

def _proje_verisini_hazirla(proje_adi=None):
    proje_adi = str(proje_adi or _proje_adi()).strip()
    return {
        "proje_dosya_suru": 2,
        "proje_adi": proje_adi,
        "proje_id": _proje_kisa_dosya_adi(proje_adi) if proje_adi else "",
        "olusturma_tarihi": datetime.now().isoformat(timespec="seconds"),
        "sirket_adi": str(st.session_state.get("sirket_adi", "")),
        "rapor_turu": str(st.session_state.get("rapor_turu", "")),
        "hazirlayan": str(st.session_state.get("hazirlayan", "")),
        "mmo_no": str(st.session_state.get("mmo_no", "")),
        "rapor_tarihi": str(st.session_state.get("tarih", "")),
        "session_state": _proje_ayarlarini_topla(),
    }


def _proje_olustur_veya_kaydet(proje_adi, farkli_kaydet=False):
    proje_adi = str(proje_adi or "").strip()
    if not proje_adi:
        st.session_state["proje_yukleme_bildirimi"] = (
            "Proje adı boş bırakılamaz. Proje Yönetimi bölümünden proje adını girin."
        )
        return False

    proje_id = _proje_kisa_dosya_adi(proje_adi)
    aktif = st.session_state.get("aktif_proje_id", "")
    try:
        if farkli_kaydet and _PROJE_YONETICISI.proje_var_mi(proje_id):
            st.session_state["proje_yukleme_bildirimi"] = (
                f"Bu isimde proje zaten var: {proje_adi}"
            )
            return False

        bilgiler = {
            "sirket_adi": str(st.session_state.get("sirket_adi", "")),
            "rapor_turu": str(st.session_state.get("rapor_turu", "")),
            "hazirlayan": str(st.session_state.get("hazirlayan", "")),
            "mmo_no": str(st.session_state.get("mmo_no", "")),
            "rapor_tarihi": str(st.session_state.get("tarih", "")),
            "genel": {"session_state": _proje_ayarlarini_topla()},
        }

        # Aktif proje adı değiştirilmişse Kaydet işlemi mevcut projeyi yeni ada taşır.
        if (not farkli_kaydet) and aktif and aktif != proje_id and _PROJE_YONETICISI.proje_var_mi(proje_id):
            st.session_state["proje_yukleme_bildirimi"] = (
                f"'{proje_adi}' adıyla başka bir proje zaten var. Farklı bir isim seçin."
            )
            return False

        if (not farkli_kaydet) and aktif and aktif != proje_id and _PROJE_YONETICISI.proje_var_mi(aktif):
            eski_klasor = Path(_PROJE_YONETICISI.ana_dizin) / aktif
            eski_veri = _PROJE_YONETICISI.ac(aktif)
            eski_veri.proje_adi = proje_adi
            eski_veri.proje_id = proje_id
            eski_veri.genel = bilgiler["genel"]
            eski_veri.sirket_adi = bilgiler["sirket_adi"]
            eski_veri.rapor_turu = bilgiler["rapor_turu"]
            eski_veri.hazirlayan = bilgiler["hazirlayan"]
            eski_veri.mmo_no = bilgiler["mmo_no"]
            eski_veri.rapor_tarihi = bilgiler["rapor_tarihi"]
            yeni_klasor = Path(_PROJE_YONETICISI.ana_dizin) / proje_id
            eski_klasor.rename(yeni_klasor)
            _PROJE_YONETICISI.kaydet(eski_veri)
            proje = eski_veri
        elif farkli_kaydet or not _PROJE_YONETICISI.proje_var_mi(proje_id):
            proje = _PROJE_YONETICISI.proje_olustur(
                proje_adi, proje_id=proje_id, **bilgiler
            )
        else:
            proje = _PROJE_YONETICISI.ac(proje_id)
            proje.proje_adi = proje_adi
            proje.genel = bilgiler["genel"]
            proje.sirket_adi = bilgiler["sirket_adi"]
            proje.rapor_turu = bilgiler["rapor_turu"]
            proje.hazirlayan = bilgiler["hazirlayan"]
            proje.mmo_no = bilgiler["mmo_no"]
            proje.rapor_tarihi = bilgiler["rapor_tarihi"]
            _PROJE_YONETICISI.kaydet(proje)

        st.session_state["aktif_proje_id"] = proje.proje_id
        st.session_state["aktif_proje_adi"] = proje.proje_adi
        st.session_state["proje_kaynak"] = "yerel"
        st.session_state["proje_son_kayit_zamani"] = datetime.now().strftime("%H:%M:%S")
        st.session_state["proje_yukleme_bildirimi"] = f"'{proje.proje_adi}' kaydedildi."
        return True
    except Exception as hata:
        st.session_state["proje_yukleme_bildirimi"] = f"Proje kaydedilemedi: {hata}"
        return False


def _proje_kaydet_callback():
    _proje_olustur_veya_kaydet(_proje_adi(), farkli_kaydet=False)


def _yeni_proje_baslat():
    """Yeni proje isteğini işaretler; gerçek session temizliği bir sonraki rerunun
    başında, widgetlar yeniden oluşturulmadan önce yapılır.

    Bu yaklaşım StreamlitValueAssignmentNotAllowedError oluşmasını önler.
    """
    st.session_state["_yeni_proje_sifirlama_bekliyor"] = True



def _proje_sil(proje_id):
    try:
        if not proje_id:
            return
        klasor = Path(_PROJE_YONETICISI.ana_dizin) / proje_id
        if not klasor.exists():
            st.session_state["proje_yukleme_bildirimi"] = "Silinecek proje bulunamadı."
            return
        shutil.rmtree(klasor)
        if st.session_state.get("aktif_proje_id") == proje_id:
            st.session_state.pop("aktif_proje_id", None)
            st.session_state.pop("aktif_proje_adi", None)
            st.session_state["proje_kaynak"] = ""
            st.session_state["proje_son_kayit_zamani"] = ""
        st.session_state["proje_yukleme_bildirimi"] = f"'{proje_id}' projesi silindi."
    except Exception as hata:
        st.session_state["proje_yukleme_bildirimi"] = f"Proje silinemedi: {hata}"


def _proje_yeni():
    """Mevcut aktif çalışmayı kapatıp temiz bir yeni proje oturumu başlatır."""
    try:
        # Projeye ait kaydedilebilir tüm kullanıcı verilerini temizle.
        # Buton/aksiyon widgetları özellikle korunur; aksi halde Streamlit
        # ValueAssignment hatası oluşabilir.
        mevcut_ayarlar = _proje_ayarlarini_topla()
        for anahtar in list(mevcut_ayarlar.keys()):
            if anahtar not in _PROJE_META_KEYS and not _proje_widget_anahtari_mi(anahtar):
                st.session_state.pop(anahtar, None)

        # Proje yönetimi durumunu sıfırla.
        for anahtar in (
            "aktif_proje_id", "aktif_proje_adi", "proje_son_kayit_zamani",
            "proje_kaynak", "proje_dis_dosya_hazirlandi",
            "proje_dis_dosya_yukle_bekliyor",
        ):
            st.session_state.pop(anahtar, None)

        st.session_state["proje_adi_giris"] = ""
        st.session_state["_proje_adi_manuel"] = False
        st.session_state["proje_yukleme_bildirimi"] = "Yeni proje açıldı. Yeni proje bilgilerini girebilirsiniz."
    except Exception as hata:
        st.session_state["proje_yukleme_bildirimi"] = f"Yeni proje açılamadı: {hata}"


def _proje_basliktan_al():
    ad = str(st.session_state.get("is_adi", "")).strip()
    st.session_state["proje_adi_giris"] = ad
    st.session_state["_proje_adi_manuel"] = False
    st.session_state["proje_yukleme_bildirimi"] = "Proje adı kapak başlığından alındı."


def _proje_adi_manuel_degisti():
    st.session_state["_proje_adi_manuel"] = True


def _is_adi_degisti():
    # Kullanıcı proje adını elle değiştirmediyse proje adı kapak başlığıyla eşit tutulur.
    if not st.session_state.get("_proje_adi_manuel", False):
        st.session_state["proje_adi_giris"] = str(st.session_state.get("is_adi", ""))


def _proje_ac(proje_id):
    try:
        # Mevcut geçici widget durumlarını temizle.
        for _k in list(st.session_state.keys()):
            if _proje_widget_anahtari_mi(_k):
                st.session_state.pop(_k, None)

        proje = _PROJE_YONETICISI.ac(proje_id)
        ayarlar = proje.genel.get("session_state", {}) if isinstance(proje.genel, dict) else {}
        for anahtar, deger in ayarlar.items():
            if _proje_widget_anahtari_mi(anahtar) or anahtar.startswith("proje_"):
                continue
            st.session_state[anahtar] = _proje_degerini_geri_yukle(deger)
        st.session_state["aktif_proje_id"] = proje.proje_id
        st.session_state["aktif_proje_adi"] = proje.proje_adi
        st.session_state["proje_adi_giris"] = proje.proje_adi
        st.session_state["_proje_adi_manuel"] = True
        st.session_state["proje_kaynak"] = "yerel"
        st.session_state["proje_son_kayit_zamani"] = ""
        st.session_state["proje_yukleme_bildirimi"] = f"'{proje.proje_adi}' açıldı."
    except Exception as hata:
        st.session_state["proje_yukleme_bildirimi"] = f"Proje açılamadı: {hata}"


def _proje_otomatik_kaydet():
    """Aktif yerel projeyi her Streamlit rerun'unda günceller.
    Dışarıdan yüklenen projeler tarayıcı güvenliği nedeniyle kaynak dosyaya
    sessizce geri yazılamaz; onlar için Farklı Kaydet tekrar kullanılmalıdır.
    """
    aktif = st.session_state.get("aktif_proje_id", "")
    kaynak = st.session_state.get("proje_kaynak", "")
    ad = _proje_adi()
    if not aktif or kaynak != "yerel" or not ad:
        return
    if aktif != _proje_kisa_dosya_adi(ad):
        return
    try:
        _proje_olustur_veya_kaydet(ad, farkli_kaydet=False)
    except Exception:
        pass


def _proje_dis_dosyayi_yukle_callback():
    """Bilgisayardan seçilen .proje.json dosyasını gerçekten oku ve bir sonraki
    rerunda uygulanmak üzere beklet.

    ÖNEMLİ: file_uploader'ın session-state anahtarını burada silmiyoruz.
    Önceki sürümde bu anahtar _proje_widget_anahtari_mi() filtresine takıldığı
    için UploadedFile nesnesi callback çalışmadan önce kayboluyor ve dosya
    ekranda görünmesine rağmen "açılmıyor" gibi davranıyordu.
    """
    yuklenen = st.session_state.get("proje_dis_dosya_yukle_v116")
    if yuklenen is None:
        st.session_state["proje_yukleme_bildirimi"] = "Lütfen bir .proje.json dosyası seçin."
        return

    try:
        ham_veri = yuklenen.getvalue()
        if not ham_veri:
            raise ValueError("Seçilen proje dosyası boş.")

        veri = json.loads(ham_veri.decode("utf-8-sig"))
        if not isinstance(veri, dict):
            raise ValueError("Proje dosyasının ana yapısı geçersiz.")

        ayarlar = veri.get("session_state", {})
        if not isinstance(ayarlar, dict):
            raise ValueError("Proje dosyasında geçerli session_state bulunamadı.")

        # Eski/yeni format uyumluluğu: bazı dış proje dosyalarında bilgiler
        # doğrudan 'genel.session_state' altında olabilir.
        if not ayarlar and isinstance(veri.get("genel"), dict):
            ayarlar = veri["genel"].get("session_state", {})
            if not isinstance(ayarlar, dict):
                ayarlar = {}

        # Yüklenen dosyanın gerçek içeriğini bir sonraki reruna taşıyoruz.
        st.session_state["proje_dis_dosya_yukle_bekliyor"] = ayarlar
        st.session_state["aktif_proje_id"] = str(veri.get("proje_id", ""))
        st.session_state["aktif_proje_adi"] = str(veri.get("proje_adi", ""))
        st.session_state["proje_adi_giris"] = str(veri.get("proje_adi", ""))
        st.session_state["_proje_adi_manuel"] = True
        st.session_state["proje_kaynak"] = "harici"
        st.session_state["proje_son_kayit_zamani"] = ""
        st.session_state["proje_yukleme_bildirimi"] = (
            f"'{veri.get('proje_adi', 'Proje')}' dosyası okundu ve yükleniyor..."
        )

        # Widget callback'inden güvenli şekilde çıkıp yeni session state ile
        # tam bir Streamlit rerun başlat.
        st.rerun()
    except Exception as hata:
        st.session_state["proje_yukleme_bildirimi"] = f"Proje dosyası açılamadı: {hata}"


_bekleyen_dis_proje = st.session_state.pop("proje_dis_dosya_yukle_bekliyor", None)
if isinstance(_bekleyen_dis_proje, dict):
    for _k in list(st.session_state.keys()):
        if _proje_widget_anahtari_mi(_k):
            st.session_state.pop(_k, None)
    for _anahtar, _deger in _bekleyen_dis_proje.items():
        if _proje_widget_anahtari_mi(_anahtar) or _anahtar.startswith("proje_"):
            continue
        st.session_state[_anahtar] = _proje_degerini_geri_yukle(_deger)

# ---------------------------------------------------------------------------
# YENİ PROJE İSTEĞİNİ, WIDGETLAR OLUŞMADAN ÖNCE UYGULA
# ---------------------------------------------------------------------------
if st.session_state.pop("_yeni_proje_sifirlama_bekliyor", False):
    # Yeni proje geçişinde korunacak kurumsal bilgiler.
    _korunacak_kurumsal = {}
    for _k in list(st.session_state.keys()):
        _ks = str(_k).lower()
        _kurumsal = (
            _ks in {
                "sirket_adi", "firma_adi", "kurulus_adi",
                "isveren", "isveren_adi", "isveren_bilgileri",
                "adres", "proje_adresi", "firma_adresi",
            }
            or "isveren" in _ks
            or "adres" in _ks
            or "sirket" in _ks
            or "firma" in _ks
            or "kurulus" in _ks
        )
        if _kurumsal:
            _korunacak_kurumsal[_k] = st.session_state.get(_k)

    # Eski projenin tüm widget/proje değerlerini kaldır.
    st.session_state.clear()

    # Kurumsal bilgiler yeni projeye aynen taşınır.
    for _k, _v in _korunacak_kurumsal.items():
        st.session_state[_k] = _v

    st.session_state["aktif_proje_id"] = ""
    st.session_state["aktif_proje_adi"] = ""
    st.session_state["proje_adi_giris"] = ""
    st.session_state["proje_kaynak"] = ""
    st.session_state["proje_son_kayit_zamani"] = ""
    st.session_state["proje_dis_dosya_hazirlandi"] = False
    st.session_state["proje_dis_dosya_yukle_bekliyor"] = None
    st.session_state["_proje_adi_manuel"] = False
    st.session_state["proje_yukleme_bildirimi"] = (
        "🆕 Yeni proje başlatıldı. Proje bilgileri temizlendi; "
        "şirket/işveren/adres bilgileri korundu."
    )


def _rapor_docx_pdf_donustur(docx_bytes):
    """
    DOCX raporunu PDF'e dönüştürür.
    1) LibreOffice/soffice varsa orijinal Word düzenine en yakın PDF'i üretir.
    2) LibreOffice yoksa matplotlib'in yerleşik PDF backend'i ile PDF üretir.
       Böylece ayrıca ReportLab kurulumu zorunlu değildir.
    """
    _hatalar = []

    # ---------------------------------------------------------------
    # 1. Tercih: LibreOffice / soffice
    # ---------------------------------------------------------------
    with tempfile.TemporaryDirectory() as _tmp:
        _docx = Path(_tmp) / "rapor.docx"
        _docx.write_bytes(docx_bytes)

        for _soffice in ("libreoffice", "soffice"):
            try:
                _sonuc = subprocess.run(
                    [
                        _soffice,
                        "--headless",
                        "--convert-to", "pdf",
                        "--outdir", _tmp,
                        str(_docx),
                    ],
                    capture_output=True,
                    text=True,
                    timeout=90,
                )
                _pdf = Path(_tmp) / "rapor.pdf"
                if _sonuc.returncode == 0 and _pdf.exists() and _pdf.stat().st_size > 100:
                    return _pdf.read_bytes()

                _hatalar.append(
                    f"{_soffice}: {_sonuc.stderr.strip() or _sonuc.stdout.strip()}"
                )
            except Exception as _e:
                _hatalar.append(f"{_soffice}: {_e}")

    # ---------------------------------------------------------------
    # 2. Güvenli fallback: matplotlib PdfPages
    # ---------------------------------------------------------------
    try:
        from matplotlib.backends.backend_pdf import PdfPages
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        import textwrap

        _doc = Document(io.BytesIO(docx_bytes))

        # Türkçe karakterleri destekleyen font.
        _font_name = "DejaVu Sans"

        _elements = []

        # Paragraflar
        for _p in _doc.paragraphs:
            _txt = _p.text.strip()
            if not _txt:
                continue

            _style = (_p.style.name or "").lower()
            if "heading 1" in _style:
                _kind = "h1"
            elif "heading 2" in _style:
                _kind = "h2"
            elif "heading 3" in _style:
                _kind = "h3"
            else:
                _kind = "p"

            _elements.append((_kind, _txt))

        # Tabloları metinsel ama okunabilir biçimde ekle.
        for _table in _doc.tables:
            _elements.append(("table", ""))
            for _row in _table.rows:
                _elements.append(
                    (
                        "table_row",
                        "  |  ".join(_cell.text.replace("\n", " / ") for _cell in _row.cells)
                    )
                )
            _elements.append(("table_end", ""))

        if not _elements:
            return None

        _tmp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        _tmp_pdf.close()
        _pdf_path = Path(_tmp_pdf.name)

        try:
            with PdfPages(str(_pdf_path)) as _pdf:
                _fig = None
                _ax = None
                _y = 0.96
                _table_mode = False

                def _new_page():
                    nonlocal _fig, _ax, _y, _table_mode
                    if _fig is not None:
                        _pdf.savefig(_fig, bbox_inches="tight")
                        plt.close(_fig)
                    _fig = Figure(figsize=(8.27, 11.69))
                    FigureCanvasAgg(_fig)
                    _ax = _fig.add_axes([0.075, 0.055, 0.86, 0.89])
                    _ax.axis("off")
                    _ax.set_xlim(0, 1)
                    _ax.set_ylim(0, 1)
                    _y = 0.98
                    _table_mode = False

                _new_page()

                for _kind, _txt in _elements:
                    if _kind == "table":
                        _table_mode = True
                        _y -= 0.012
                        continue

                    if _kind == "table_end":
                        _table_mode = False
                        _y -= 0.012
                        continue

                    if _kind == "table_row":
                        _lines = textwrap.wrap(
                            _txt, width=105, replace_whitespace=False
                        ) or [""]
                        for _line in _lines:
                            if _y < 0.055:
                                _new_page()
                            _ax.text(
                                0.015, _y, _line,
                                fontsize=7.2,
                                fontname=_font_name,
                                va="top",
                            )
                            _y -= 0.022
                        continue

                    if _kind == "h1":
                        _fs, _weight, _space = 17, "bold", 0.055
                    elif _kind == "h2":
                        _fs, _weight, _space = 13, "bold", 0.042
                    elif _kind == "h3":
                        _fs, _weight, _space = 11, "bold", 0.035
                    else:
                        _fs, _weight, _space = 9.2, "normal", 0.025

                    if _y < 0.09:
                        _new_page()

                    _lines = textwrap.wrap(
                        _txt, width=100 if _kind == "p" else 85,
                        replace_whitespace=False
                    ) or [""]

                    for _idx, _line in enumerate(_lines):
                        if _y < 0.055:
                            _new_page()
                        _ax.text(
                            0.0, _y, _line,
                            fontsize=_fs,
                            fontname=_font_name,
                            fontweight=_weight,
                            color=(
                                "#0B3D91" if _kind == "h1"
                                else "#0F5B78" if _kind == "h3"
                                else "#111111"
                            ),
                            va="top",
                        )
                        _y -= 0.024 if _kind == "p" else 0.030

                    _y -= _space

                if _fig is not None:
                    _pdf.savefig(_fig, bbox_inches="tight")
                    plt.close(_fig)

            _data = _pdf_path.read_bytes()
            if _data and len(_data) > 100:
                return _data

        finally:
            try:
                _pdf_path.unlink(missing_ok=True)
            except Exception:
                pass

    except Exception as _e:
        _hatalar.append(f"matplotlib PDF: {_e}")

    return None


def _rapor_docx_html(docx_bytes):
    """DOCX'teki temel paragraf ve tabloları bağımsız HTML çıktısına çevirir."""
    _doc = Document(io.BytesIO(docx_bytes))
    _html = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        "<title>Mekanik Uygulama Raporu</title>",
        "<style>body{font-family:Arial,sans-serif;margin:40px;line-height:1.45}"
        "table{border-collapse:collapse;width:100%;margin:12px 0}"
        "td,th{border:1px solid #999;padding:6px}"
        "h1{color:#0B3D91}h2{color:#0B3D91}h3{color:#155E75}"
        "</style></head><body>"
    ]
    for _p in _doc.paragraphs:
        _txt = _p.text.strip()
        if not _txt:
            continue
        _style = (_p.style.name or "").lower()
        if "heading 1" in _style:
            _html.append(f"<h1>{_txt}</h1>")
        elif "heading 2" in _style:
            _html.append(f"<h2>{_txt}</h2>")
        elif "heading 3" in _style:
            _html.append(f"<h3>{_txt}</h3>")
        else:
            _html.append(f"<p>{_txt}</p>")
    for _table in _doc.tables:
        _html.append("<table>")
        for _ri, _row in enumerate(_table.rows):
            _html.append("<tr>")
            for _cell in _row.cells:
                _tag = "th" if _ri == 0 else "td"
                _html.append(f"<{_tag}>{_cell.text}</{_tag}>")
            _html.append("</tr>")
        _html.append("</table>")
    _html.append("</body></html>")
    return "\n".join(_html).encode("utf-8")


def _rapor_docx_txt(docx_bytes):
    _doc = Document(io.BytesIO(docx_bytes))
    _satirlar = []
    for _p in _doc.paragraphs:
        if _p.text.strip():
            _satirlar.append(_p.text)
    for _table in _doc.tables:
        for _row in _table.rows:
            _satirlar.append(" | ".join(_cell.text for _cell in _row.cells))
    return ("\n".join(_satirlar) + "\n").encode("utf-8")

# Opsiyonel pis su pompası bölümü kapalı olsa bile rapor oluşturma kontrolünde
# bu değişkenin güvenli biçimde bulunması gerekir.
psp_parametreleri = globals().get("psp_parametreleri", {})

# ---------------------------------------------------------------------------
# PROJE KONTROL PANELİ
# ---------------------------------------------------------------------------
with st.sidebar.expander("💾 PROJE YÖNETİMİ", expanded=True):
    st.button(
        "🆕 Yeni Proje",
        key="proje_yeni_btn_v119",
        use_container_width=True,
        on_click=_yeni_proje_baslat,
        help="Mevcut çalışmayı kapatır ve temiz bir yeni proje başlatır. Kayıtlı projeler silinmez.",
    )

    st.markdown("---")
    st.markdown("### 📄 RAPOR OLUŞTURMA")

    _rapor_format = st.selectbox(
        "Çıktı formatı",
        ["Word (.docx)", "PDF (.pdf)", "HTML (.html)", "Metin (.txt)"],
        key="rapor_cikti_format_v134",
    )

    # Bu bir aksiyon widget'ıdır; proje JSON'una hiçbir şekilde kaydedilmez.
    if st.button(
        "📄 RAPORU OLUŞTUR",
        key="sidebar_rapor_olustur_v134",
        use_container_width=True,
        help="Hesap raporunu seçtiğiniz çıktı formatında hazırlar.",
    ):
        st.session_state["_rapor_olustur_istegi_v134"] = True
        st.rerun()

    if st.session_state.get("_rapor_hazir_docx_v134"):
        _rapor_docx_veri = st.session_state["_rapor_hazir_docx_v134"]
        _rapor_ad = st.session_state.get(
            "_rapor_hazir_adi_v134", "Mekanik_Uygulama_Raporu"
        )

        if _rapor_format == "Word (.docx)":
            st.download_button(
                "📥 Word'u İndir",
                data=_rapor_docx_veri,
                file_name=f"{_rapor_ad}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="rapor_word_indir_v134",
                use_container_width=True,
            )
        elif _rapor_format == "PDF (.pdf)":
            _pdf_veri = _rapor_docx_pdf_donustur(_rapor_docx_veri)
            if _pdf_veri:
                st.download_button(
                    "📥 PDF'yi İndir",
                    data=_pdf_veri,
                    file_name=f"{_rapor_ad}.pdf",
                    mime="application/pdf",
                    key="rapor_pdf_indir_v134",
                    use_container_width=True,
                )
            else:
                st.error(
                    "PDF oluşturulamadı. Word/HTML çıktısı kullanılabilir; "
                    "PDF motoru bu çalışma ortamında başlatılamadı."
                )
        elif _rapor_format == "HTML (.html)":
            st.download_button(
                "📥 HTML'yi İndir",
                data=_rapor_docx_html(_rapor_docx_veri),
                file_name=f"{_rapor_ad}.html",
                mime="text/html",
                key="rapor_html_indir_v134",
                use_container_width=True,
            )
        else:
            st.download_button(
                "📥 Metin Dosyasını İndir",
                data=_rapor_docx_txt(_rapor_docx_veri),
                file_name=f"{_rapor_ad}.txt",
                mime="text/plain",
                key="rapor_txt_indir_v134",
                use_container_width=True,
            )

    _mevcut_projeler = []
    try:
        _mevcut_projeler = sorted([
            p.name for p in Path(_PROJE_YONETICISI.ana_dizin).iterdir()
            if p.is_dir() and (p / "proje.json").exists()
        ])
    except Exception:
        _mevcut_projeler = []

    st.caption("Yeni proje açmadan önce mevcut çalışmayı Kaydet veya Farklı Kaydet ile saklayabilirsiniz.")

    st.text_input(
        "Proje Adı",
        key="proje_adi_giris",
        value=st.session_state.get("proje_adi_giris", st.session_state.get("is_adi", "")),
        on_change=_proje_adi_manuel_degisti,
        help="İlk değer İşin Adı / Proje Başlığı alanından gelir. Buradan elle değiştirebilirsiniz.",
    )
    st.caption("İlk değer otomatik olarak kapaktaki proje başlığından alınır; isterseniz burada değiştirebilirsiniz.")
    st.button("↻ Başlıktan tekrar al", key="proje_basliktan_al_btn_v116", use_container_width=True, on_click=_proje_basliktan_al)

    _pc1, _pc2 = st.columns(2)
    with _pc1:
        st.button("💾 Kaydet", key="proje_kaydet_btn_v116", use_container_width=True, on_click=_proje_kaydet_callback)
    with _pc2:
        _dis_ad_now = _proje_adi()
        _dis_veri_now = _proje_verisini_hazirla(_dis_ad_now) if _dis_ad_now else {"session_state": {}}
        _dis_json_now = json.dumps(_dis_veri_now, ensure_ascii=False, indent=2)
        _dis_dosya_adi_now = f"{_proje_kisa_dosya_adi(_dis_ad_now)}.proje.json" if _dis_ad_now else "Proje.proje.json"
        st.download_button(
            "📑 Farklı Kaydet",
            data=_dis_json_now.encode("utf-8"),
            file_name=_dis_dosya_adi_now,
            mime="application/json",
            key="proje_dis_dosya_indir_v116",
            use_container_width=True,
            help="Tarayıcınızın indirme ayarına göre dosya konumunu seçebilirsiniz.",
        )

    if _mevcut_projeler:
        _secili_proje = st.selectbox(
            "Uygulama İçindeki Kayıtlı Proje",
            _mevcut_projeler,
            index=(_mevcut_projeler.index(st.session_state.get("aktif_proje_id"))
                   if st.session_state.get("aktif_proje_id") in _mevcut_projeler else 0),
            key="proje_ac_sec_v116",
        )
        _pc3, _pc4 = st.columns(2)
        with _pc3:
            st.button(
                "📂 Aç", key="proje_ac_btn_v116", use_container_width=True,
                on_click=lambda: _proje_ac(st.session_state.get("proje_ac_sec_v116", "")),
            )
        with _pc4:
            st.button(
                "🗑️ Sil", key="proje_sil_btn_v116", use_container_width=True,
                on_click=lambda: _proje_sil(st.session_state.get("proje_ac_sec_v116", "")),
            )
    else:
        st.caption("Henüz uygulama içinde kayıtlı proje yok.")

    st.markdown("**📂 Bilgisayardan / Masaüstünden Proje Aç**")
    st.file_uploader(
        "Proje dosyasını seçin (.proje.json)",
        type=["json"],
        key="proje_dis_dosya_yukle_v116",
        help="Daha önce Farklı Kaydet ile oluşturduğunuz .proje.json dosyasını seçin.",
    )
    st.button(
        "📂 Seçilen Projeyi Yükle", key="proje_dis_dosya_ac_btn_v116",
        use_container_width=True, on_click=_proje_dis_dosyayi_yukle_callback,
    )

    if st.session_state.get("aktif_proje_adi"):
        st.success(f"💾 **Aktif proje:** {st.session_state['aktif_proje_adi']}")
    if st.session_state.get("proje_son_kayit_zamani"):
        st.caption(f"Son kayıt: {st.session_state['proje_son_kayit_zamani']}")
    if st.session_state.get("proje_yukleme_bildirimi"):
        st.info(st.session_state["proje_yukleme_bildirimi"])

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
def enable_update_fields_on_open(doc):
    """Word belgesini açarken alanların (özellikle İçindekiler) güncellenmesini ister."""
    settings = doc.settings.element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")


def add_toc(paragraph):
  """Word TOC alanı; başlıklar ve sayfa numaraları Word tarafından güncellenir."""
  run = paragraph.add_run()
  fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin"); fld_begin.set(qn("w:dirty"), "true")
  instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = 'TOC \\o "1-3" \\h \\z \\u'
  fld_sep = OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
  txt = OxmlElement("w:t"); txt.text = "İçindekiler güncelleniyor..."
  fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end"); fld_end.set(qn("w:dirty"), "true")
  run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_sep); run._r.append(txt); run._r.append(fld_end)

# 6.3 alt bölümlerinin dinamik numaralandırma kaynağı.
# Bu liste, bölüm seçimi kapatılıp açıldığında numaranın otomatik yeniden
# sıralanmasını sağlar.
_BOLUM_63_COCUKLARI = [
    ("rapor_bolum_631", "SU DEPOSU KAPASİTE HESAPLAMALARI"),
    ("rapor_bolum_632", "KULLANMA SOĞUK SUYU HİDROFORU SEÇİMİ"),
    ("rapor_bolum_633", "KULLANMA SICAK SUYU İHTİYACI HESAPLARI"),
    ("rapor_bolum_634", "KULLANMA SICAK SU TESİSATI RE-SİRKULASYON POMPASI SEÇİMİ"),
]

# Sol menüdeki "Tümünü Seç / Tümünü Kaldır" işlemlerinin kullandığı anahtarlar.
BOLUM_SECIM_ANAHTARLARI = [
    "rapor_bolum_1", "rapor_bolum_2", "rapor_bolum_3", "rapor_bolum_4",
    "rapor_bolum_5", "rapor_bolum_51", "rapor_bolum_6", "rapor_bolum_61",
    "rapor_bolum_611", "rapor_bolum_62", "rapor_bolum_621", "rapor_bolum_622",
    "rapor_bolum_63", "rapor_bolum_631", "rapor_bolum_631_1", "rapor_bolum_631_2",
    "rapor_bolum_631_2_1", "rapor_bolum_631_2_2", "rapor_bolum_631_2_3",
    "rapor_bolum_631_2_4", "rapor_bolum_631_2_5", "rapor_bolum_631_2_6",
    "rapor_bolum_631_2_7", "rapor_bolum_631_2_8", "rapor_bolum_632",
    "rapor_bolum_633", "rapor_bolum_634",
]
def _63_dinamik_no(anahtar):
    """Aktif 6.3 alt bölümleri içindeki sıralı numarayı döndürür."""
    aktifler = [k for k, _ in _BOLUM_63_COCUKLARI if bool(st.session_state.get(k, False))]
    try:
        return aktifler.index(anahtar) + 1
    except ValueError:
        # Bölüm kapalıysa, açıldığı anda doğal sırasını korusun.
        return [k for k, _ in _BOLUM_63_COCUKLARI].index(anahtar) + 1

def _63_dinamik_baslik(anahtar):
    ad = dict(_BOLUM_63_COCUKLARI)[anahtar]
    return f"6.3.{_63_dinamik_no(anahtar)} {ad.upper()}"

def _63_sidebar_baslik(anahtar, varsayilan_baslik):
    """Sol menüde 6.3 alt bölüm numarasını aktif seçimlere göre dinamik gösterir."""
    if anahtar in dict(_BOLUM_63_COCUKLARI):
        ad = dict(_BOLUM_63_COCUKLARI)[anahtar]
        return f"6.3.{_63_dinamik_no(anahtar)} {ad.upper()}"
    return varsayilan_baslik

def _tum_bolumleri_sec():
    for _anahtar in BOLUM_SECIM_ANAHTARLARI:
        st.session_state[_anahtar] = True

def _tum_bolumleri_kaldir():
    for _anahtar in BOLUM_SECIM_ANAHTARLARI:
        st.session_state[_anahtar] = False

# ---------------------------------------------------------------------------
# SOL MENÜ / HİYERARŞİK BÖLÜM NAVİGASYONU
# ---------------------------------------------------------------------------
# Ana tesisat grupları kendi alt bölümlerini içerir. Böylece sol menüde
# örneğin 6. SIHHİ TESİSAT açıldığında yalnızca 6.x bölümleri,
# 7. YANGIN TESİSATI açıldığında ise 7.x bölümleri görülür.
_BOLUM_NAV_GENEL = [
    ("1. Kapak Bilgileri", "bolum_1", "rapor_bolum_1"),
    ("2. Uygulanacak Standart ve Yönetmelikler", "bolum_2", "rapor_bolum_2"),
    ("3. Mekanik Tesisat Proje Kapsamı", "bolum_3", "rapor_bolum_3"),
    ("4. Tesiste Kullanılacak Isı İletim Akışkanları", "bolum_4", "rapor_bolum_4"),
    ("5. İklim, Konfor Şartları ve Tasarım Kriterleri", "bolum_5", "rapor_bolum_5"),
    ("5.1 Dış Hava Tasarım Kriterleri", "bolum_51", "rapor_bolum_51"),
]

_BOLUM_NAV_SIHHI = [
    ("6. Sıhhi Tesisat", "bolum_6", "rapor_bolum_6"),
    ("6.1 Sıhhi Tesisat Ön Bilgiler", "bolum_61", "rapor_bolum_61"),
    ("6.1.1 Temiz Su Hesabı", "bolum_611", "rapor_bolum_611"),
    ("6.2 Pis Su Tesisatı Esasları", "bolum_62", "rapor_bolum_62"),
    ("6.2.1 Pis Su Hesabı", "bolum_621", "rapor_bolum_621"),
    ("6.2.2 Pis Su Terfi Pompaları", "bolum_622", "rapor_bolum_622"),
    ("6.3 Sıhhi Tesisat Cihaz Seçimleri", "bolum_63", "rapor_bolum_63"),
    ("6.3.1 SU DEPOSU KAPASİTE HESAPLAMALARI", "bolum_631", "rapor_bolum_631"),
    ("6.3.1.1 Kullanma Suyu Deposu Seçimi", "bolum_631_1", "rapor_bolum_631_1"),
    ("6.3.1.2 Yağmur Suyu Deposu Seçimi", "bolum_631_2", "rapor_bolum_631_2"),
    ("6.3.1.2.1 Yağmur Suyu Toplama Hesabı", "bolum_631_2_1", "rapor_bolum_631_2_1"),
    ("6.3.1.2.2 Yağmur Suyu Filtresi Seçimi", "bolum_631_2_2", "rapor_bolum_631_2_2"),
    ("6.3.1.2.3 İlk Yağış Ayırıcı Seçimi", "bolum_631_2_3", "rapor_bolum_631_2_3"),
    ("6.3.1.2.4 Yağmur Suyu Deposu Hacim Hesabı", "bolum_631_2_4", "rapor_bolum_631_2_4"),
    ("6.3.1.2.5 Taşma Hattı Hesabı", "bolum_631_2_5", "rapor_bolum_631_2_5"),
    ("6.3.1.2.6 Taşma Sifonu / Koku Kapanı", "bolum_631_2_6", "rapor_bolum_631_2_6"),
    ("6.3.1.2.7 Depo Girişi / Sakin Giriş", "bolum_631_2_7", "rapor_bolum_631_2_7"),
    ("6.3.1.2.8 Havalandırma ve Haşere Koruması", "bolum_631_2_8", "rapor_bolum_631_2_8"),
    ("6.3.2 KULLANMA SOĞUK SUYU HİDROFORU SEÇİMİ", "bolum_632", "rapor_bolum_632"),
    ("6.3.3 KULLANMA SICAK SUYU İHTİYACI HESAPLARI", "bolum_633", "rapor_bolum_633"),
    ("6.3.4 KULLANMA SICAK SU TESİSATI RE-SİRKULASYON POMPASI SEÇİMİ", "bolum_634", "rapor_bolum_634"),
]

_BOLUM_NAV_YANGIN = [
    ("7. YANGIN TESİSATI", "bolum_7", "rapor_bolum_7"),
    ("7.1 YANGIN TESİSATI GENEL ESASLARI", "bolum_71", "rapor_bolum_71"),
    ("7.2 YANGIN TEHLİKE SINIFI VE TASARIM KRİTERLERİ", "bolum_72", "rapor_bolum_72"),
    ("7.2.1 Bina Kullanım Amacı", "bolum_721", "rapor_bolum_721"),
    ("7.2.2 Yangın Tehlike Sınıfı", "bolum_722", "rapor_bolum_722"),
    ("7.2.3 Yangın Bölmeleri", "bolum_723", "rapor_bolum_723"),
    ("7.2.4 Tasarım Kriterleri", "bolum_724", "rapor_bolum_724"),
    ("7.2.5 Tasarım Debisi", "bolum_725", "rapor_bolum_725"),
    ("7.3 YANGIN SUYU DEPOSU HESABI", "bolum_73", "rapor_bolum_73"),
    ("7.3.1 Gerekli Yangın Suyu Hacmi", "bolum_731", "rapor_bolum_731"),
    ("7.3.2 Yangın Suyu Deposu Seçimi", "bolum_732", "rapor_bolum_732"),
    ("7.3.3 Depo Hacmi Kontrolü", "bolum_733", "rapor_bolum_733"),
    ("7.4 YANGIN POMPA GRUBU SEÇİMİ", "bolum_74", "rapor_bolum_74"),
    ("7.4.1 Ana Yangın Pompası", "bolum_741", "rapor_bolum_741"),
    ("7.4.2 Yedek Yangın Pompası", "bolum_742", "rapor_bolum_742"),
    ("7.4.3 Jokey Pompa", "bolum_743", "rapor_bolum_743"),
    ("7.4.4 Pompa Basma Yüksekliği", "bolum_744", "rapor_bolum_744"),
    ("7.5 YANGIN DOLABI / HİDRANT TESİSATI", "bolum_75", "rapor_bolum_75"),
    ("7.5.1 Yangın Dolabı", "bolum_751", "rapor_bolum_751"),
    ("7.5.2 Hidrant", "bolum_752", "rapor_bolum_752"),
    ("7.5.3 Basınç Kontrolü", "bolum_753", "rapor_bolum_753"),
    ("7.6 SPRİNKLER TESİSATI", "bolum_76", "rapor_bolum_76"),
    ("7.6.1 Tehlike Sınıfı", "bolum_761", "rapor_bolum_761"),
    ("7.6.2 Tasarım Alanı", "bolum_762", "rapor_bolum_762"),
    ("7.6.3 Debi Hesabı", "bolum_763", "rapor_bolum_763"),
    ("7.6.4 Basınç Hesabı", "bolum_764", "rapor_bolum_764"),
    ("7.6.5 Hidrolik Hesap", "bolum_765", "rapor_bolum_765"),
    ("7.7 YANGIN TESİSATI HİDROLİK HESAPLARI", "bolum_77", "rapor_bolum_77"),
    ("7.8 YANGIN TESİSATI EKİPMAN SEÇİMLERİ", "bolum_78", "rapor_bolum_78"),
    ("7.9 YANGIN TESİSATI SONUÇ TABLOSU", "bolum_79", "rapor_bolum_79"),
]

_BOLUM_NAV_DIGER = [
    ("8. ISITMA TESİSATI", "bolum_8", "rapor_bolum_8"),
    ("9. SOĞUTMA TESİSATI", "bolum_9", "rapor_bolum_9"),
    ("10. HAVALANDIRMA TESİSATI", "bolum_10", "rapor_bolum_10"),
]

def _nav_satiri(_baslik, _anchor, _key, _seviye=0):
    if _key not in st.session_state:
        st.session_state[_key] = True
    if _anchor.startswith("bolum_63"):
        _baslik = _63_sidebar_baslik(_key, _baslik)
    _pad = "" if _seviye == 0 else ("&nbsp;" * (4 * _seviye))
    _c_nav, _c_chk = st.columns([8.5, 1.5], vertical_alignment="center")
    with _c_nav:
        st.markdown(
            f'<a class="proje-nav-tab" style="padding-left:{8 + _seviye*16}px;" href="#{_anchor}">{_pad}▸ {_baslik}</a>',
            unsafe_allow_html=True,
        )
    with _c_chk:
        st.checkbox("", key=_key, label_visibility="collapsed")

with st.sidebar:
    st.markdown("## 📑 PROJE BÖLÜMLERİ")
    st.caption("Ana tesisat grubunu açın; alt bölümler kendi grubunun altında gösterilir.")

    c1, c2 = st.columns(2)
    with c1:
        st.button("☑ TÜMÜNÜ SEÇ", key="rapor_tumunu_sec_v92", on_click=_tum_bolumleri_sec, use_container_width=True)
    with c2:
        st.button("☐ TÜMÜNÜ KALDIR", key="rapor_tumunu_kaldir_v92", on_click=_tum_bolumleri_kaldir, use_container_width=True)

    st.markdown("### Bölümler")

    with st.expander("📘 1–5 GENEL BİLGİLER", expanded=True):
        for _item in _BOLUM_NAV_GENEL:
            _nav_satiri(*_item, _seviye=0)

    with st.expander("🚰 6. SIHHİ TESİSAT", expanded=False):
        for _item in _BOLUM_NAV_SIHHI:
            _nav_satiri(*_item, _seviye=0)

    with st.expander("🔥 7. YANGIN TESİSATI", expanded=True):
        for _item in _BOLUM_NAV_YANGIN:
            _nav_satiri(*_item, _seviye=0)

    with st.expander("♨ 8. ISITMA TESİSATI", expanded=False):
        _nav_satiri(*_BOLUM_NAV_DIGER[0], _seviye=0)

    with st.expander("❄ 9. SOĞUTMA TESİSATI", expanded=False):
        _nav_satiri(*_BOLUM_NAV_DIGER[1], _seviye=0)

    with st.expander("💨 10. HAVALANDIRMA TESİSATI", expanded=False):
        _nav_satiri(*_BOLUM_NAV_DIGER[2], _seviye=0)

# Rapor Oluştur Butonu
_proje_otomatik_kaydet()

_rapor_olustur_sidebar = st.session_state.pop("_rapor_olustur_istegi_v134", False)
if _rapor_olustur_sidebar:

  _re_sirk_rapor_kontrol = st.session_state.get("re_sirkulasyon_pompa_sonucu_v99", {})
  gecersiz_var = any(
      p.get("poz_durumu") != "UYGUN" for p in (psp_parametreleri or {}).values()
  ) or (
      bolum_634_aktif
      and _re_sirk_rapor_kontrol
      and _re_sirk_rapor_kontrol.get("poz_durumu") != "UYGUN"
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

    def ana_baslik_ekle(metin):
        """Ana bölüm başlığını yeni sayfadan başlatır ve altındaki içerikle birlikte tutar."""
        p = doc.add_heading(metin, level=1)
        p.paragraph_format.page_break_before = True
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.keep_together = True
        p.paragraph_format.widow_control = True
        return p

    cover_section = doc.sections[0]
    cover_section.top_margin = Inches(1.15)
    cover_section.bottom_margin = Inches(1.0)
    cover_section.left_margin = Inches(1.0)
    cover_section.right_margin = Inches(1.0)
    cover_section.header_distance = Inches(0.25)
    cover_section.footer_distance = Inches(0.25)

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
      run_is.font.italic = False
      run_is.font.name = "Arial"
      run_is.font.color.rgb = RGBColor(0, 0, 0)

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

    # Word açılışında başlık alanları ve içindekiler otomatik güncellensin
    enable_update_fields_on_open(doc)

    # 2. SAYFA: İÇİNDEKİLER
    doc.add_page_break()
    doc.add_heading("İÇİNDEKİLER", level=1)
    p_toc = doc.add_paragraph()
    add_toc(p_toc)


    # 3. SAYFA: GÖVDE
    doc.add_page_break()
    body_section = doc.add_section()
    body_section.top_margin = Inches(0.85)
    body_section.bottom_margin = Inches(0.75)
    body_section.left_margin = Inches(1.0)
    body_section.right_margin = Inches(1.0)
    body_section.header_distance = Inches(0.25)
    body_section.footer_distance = Inches(0.25)

    # --- 1. GENEL BİLGİLER ---
    ana_baslik_ekle("1. GENEL BİLGİLER")
    # Genel Bilgiler içindeki proje ifadesi: tırnaksız, siyah, kalın ve italik.
    p_giris = doc.add_paragraph()
    p_giris.add_run("Bu raporda ")
    if aktif_is:
        _genel_proje_ifade = str(aktif_is).strip().strip("'‘’").strip()
        _r_proje = p_giris.add_run(_genel_proje_ifade)
        _r_proje.font.name = "Times New Roman"
        _r_proje._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        _r_proje.font.bold = True
        _r_proje.font.italic = True
        _r_proje.font.color.rgb = RGBColor(0, 0, 0)
    else:
        p_giris.add_run("ilgili proje")
    p_giris.add_run(
        " için tasarlanan mekanik tesisatlar açıklanmış ve tüm uygulama ve "
        "detay projelerine esas teşkil eden tasarım kriterleri ve mekanik "
        "tesisat sistem çözümleri tespit edilmiştir."
    )
    yapi_metni = (
        f"Yapı {secilen_il} ili {secilen_ilce} ilçesinde inşa edilecektir."
    )
    doc.add_paragraph(yapi_metni)

    if bolum_2_aktif:
            # --- 2. UYGULANACAK STANDART VE YÖNETMELİKLER ---
            ana_baslik_ekle("2. UYGULANACAK STANDART VE YÖNETMELİKLER")
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

    if bolum_3_aktif:
            # --- 3. MEKANİK TESİSAT PROJE KAPSAMI ---
            ana_baslik_ekle("3. MEKANİK TESİSAT PROJE KAPSAMI")
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

    if bolum_4_aktif:
            # --- 4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI ---
            ana_baslik_ekle("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI")
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

    if bolum_5_aktif:
            # --- 5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ ---
            ana_baslik_ekle("5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ")
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

    if bolum_6_aktif:
      # --- 6. SIHHİ TESİSAT ---
      ana_baslik_ekle("6. SIHHİ TESİSAT")
      if bolum_61_aktif:
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

        if sih_sec_yagmur_depo_tipi and sih_yagmur_depo_tipleri:
          if len(sih_yagmur_depo_tipleri) == 1:
            yagmur_tip_str = sih_yagmur_depo_tipleri[0].lower()
          elif len(sih_yagmur_depo_tipleri) == 2:
            yagmur_tip_str = (
                f"{sih_yagmur_depo_tipleri[0].lower()} ve "
                f"{sih_yagmur_depo_tipleri[1].lower()}"
            )
          else:
            ilkler = ", ".join([t.lower() for t in sih_yagmur_depo_tipleri[:-1]])
            son = sih_yagmur_depo_tipleri[-1].lower()
            yagmur_tip_str = f"{ilkler} ve {son}"

          sihhi_maddeler.append(
              "Binada yağmur suyu depolaması için "
              f"{yagmur_tip_str} tipinde su deposu kullanılmıştır."
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
      if bolum_62_aktif:
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

          # --- PİS SU POMPALARI GENEL ÇALIŞMA NOTU ---
          # Tüm pompa seçimleri rapora aktarıldıktan sonra yalnızca BİR KEZ eklenir.
          if psp_parametreleri:
            not_basligi = doc.add_paragraph()
            not_basligi.paragraph_format.space_before = Pt(8)
            not_basligi.paragraph_format.space_after = Pt(3)
            run = not_basligi.add_run(
                "Pis su pompaları ilgili genel çalışma notu:"
            )
            run.bold = True

            doc.add_paragraph(
                "İki pompalı terfi istasyonunda pompalardan birincisi, normal "
                "tüketim zamanında çalışacak, ikinci pompa yedek. üç pompalı "
                "terfi istasyonunda pompalardan birincisi ve ikincisi, normal "
                "tüketim zamanında çalışacak, üçüncü pompa yedek konumunda "
                "olacaktır. Pompaların yedeklemesi otomatik olarak münavebe ile "
                "sağlanacaktır. Pompaların elektrik panosu bu işlevleri "
                "sağlayacak şekilde imal ve monte edilecektir."
            )

        # --- 6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ ---
      if bolum_63_aktif:
        doc.add_heading("6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ", level=1)
        if bolum_631_aktif:
          doc.add_heading(_63_dinamik_baslik("rapor_bolum_631"), level=2)
          doc.add_heading("6.3.1.1 KULLANMA SUYU DEPOSU SEÇİMİ:", level=3)
          doc.add_heading("Genel Bilgiler", level=4)

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
              f"6.3.{_63_dinamik_no("rapor_bolum_631")}.1 KULLANMA SUYU İHTİYACININ BELİRLENMESİ",
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
              konut_raporu_mu = hesap_modu == "Konutlar"
              hastane_raporu_mu = hesap_modu == "Hastaneler"
              hesap_tablosu = doc.add_table(
                  rows=1,
                  cols=8 if konut_raporu_mu else (6 if hastane_raporu_mu else 5)
              )
              hesap_tablosu.style = "Table Grid"
              hesap_basliklari = hesap_tablosu.rows[0].cells
              if konut_raporu_mu:
                  basliklar = [
                      "Kategori", "Hane başına kişi", "Hane sayısı", "Toplam kişi",
                      "Birim", "Birim tüketimi", "Günlük ihtiyaç", "Açıklama"
                  ]
              elif hastane_raporu_mu:
                  basliklar = [
                      "CİHAZ", "Yatak sayısı", "Katsayı", "Toplam kişi sayısı",
                      "Tüketim [L/kişi-gün]", "Toplam [L/gün]"
                  ]
              else:
                  basliklar = ["Kategori", "Miktar", "Birim", "Birim tüketimi", "Günlük ihtiyaç"]
              for i, baslik in enumerate(basliklar):
                  hesap_basliklari[i].text = baslik

              for detay in su_hesap_detaylari:
                  hucreler = hesap_tablosu.add_row().cells
                  hucreler[0].text = detay["kategori"]
                  if konut_raporu_mu:
                      hucreler[1].text = f"{hane_kisi_sayisi:g}"
                      hucreler[2].text = f"{toplam_hane_sayisi:g}"
                      hucreler[3].text = f"{toplam_kisi_sayisi:g}"
                      hucreler[4].text = detay["birim"]
                      hucreler[5].text = f"{detay['birim_degeri']:g} L/{detay['birim']}/gün"
                      hucreler[6].text = f"{detay['ihtiyac_litre']:g} L/gün"
                      hucreler[7].text = f"{toplam_hane_sayisi:g} hane × {hane_kisi_sayisi:g} kişi"
                  elif hastane_raporu_mu:
                      hucreler[1].text = f"{detay['yatak_sayisi']:g}"
                      hucreler[2].text = f"{detay['katsayi']:g}"
                      hucreler[3].text = f"{detay['miktar']:g}"
                      hucreler[4].text = f"{detay['birim_degeri']:g} L/kişi-gün"
                      hucreler[5].text = f"{detay['ihtiyac_litre']:g} L/gün"
                  else:
                      hucreler[1].text = f"{detay['miktar']:g}"
                      hucreler[2].text = detay["birim"]
                      hucreler[3].text = f"{detay['birim_degeri']:g} L/{detay['birim']}/gün"
                      hucreler[4].text = f"{detay['ihtiyac_litre']:g} L/gün"

              if hastane_raporu_mu:
                  toplam_hucreler = hesap_tablosu.add_row().cells
                  toplam_hucreler[0].text = "GENEL TOPLAM"
                  toplam_hucreler[1].text = ""
                  toplam_hucreler[2].text = ""
                  toplam_hucreler[3].text = f"{sum(detay['miktar'] for detay in su_hesap_detaylari):g}"
                  toplam_hucreler[4].text = ""
                  toplam_hucreler[5].text = f"{su_gunluk_ihtiyac_litre:g} L/gün"
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
          if poz_gosterilsin_mi and poz_numarasi:
              poz_paragrafi = doc.add_paragraph()
              poz_paragrafi.add_run("Cihaz Poz No: ")
              poz_kalin = poz_paragrafi.add_run(
                  f"{poz_numarasi}"
              )
              poz_kalin.bold = True
              for run in poz_paragrafi.runs:
                  run.font.color.rgb = RGBColor(0, 0, 0)

        # 6.3.1.2 Yağmur suyu deposu seçimi
        _yagmur_rapor = locals().get("yagmur_hesap", {}) or st.session_state.get("yagmur_hesap", {})
        _yagmur_aktif_rapor = bool(st.session_state.get("yagmur_suyu_aktif", False)) and bool(_yagmur_rapor)
        if _yagmur_aktif_rapor:
          doc.add_heading("6.3.1.2 YAĞMUR SUYU DEPOSU SEÇİMİ:", level=3)
          _yr = _yagmur_rapor
          doc.add_heading("6.3.1.2.1 YAĞMUR SUYU TOPLAMA HESABI", level=4)
          if _yr.get("mgm_yagis_mm") is not None:
              doc.add_paragraph(
                  f"MGM verisi: {_yr.get('mgm_il', '')} ili için Günlük Toplam En Yüksek "
                  f"Yağış Miktarı = {_yr.get('mgm_yagis_mm', 0):.1f} mm "
                  f"({_yr.get('mgm_yagis_tarih', '')})."
              )

          # MGM'nin 81 il için yayımladığı günlük toplam en yüksek yağış
          # değerleri rapora eklenir. Bu tablo Streamlit arayüzünde gösterilmez.
          doc.add_heading("MGM İLLER BAZINDA GÜNLÜK TOPLAM EN YÜKSEK YAĞIŞ MİKTARLARI", level=5)
          doc.add_paragraph(
              "Aşağıdaki değerler Meteoroloji Genel Müdürlüğü (MGM) Resmi İklim "
              "İstatistikleri sayfalarında yayımlanan 'Günlük Toplam En Yüksek Yağış "
              "Miktarı' verileridir. Proje ili için hesapta kullanılan değer, ilgili "
              "satırda gösterilmektedir."
          )
          try:
              _mgm_81 = mgm_81_il_yagis_tablosu()
          except Exception:
              _mgm_81 = [(il, None, None, None) for il in MGM_81_IL]

          _mgm_tbl = doc.add_table(rows=1, cols=3)
          _mgm_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
          _mgm_tbl.autofit = True
          _hdr = _mgm_tbl.rows[0].cells
          _hdr[0].text = "İL"
          _hdr[1].text = "GÜNLÜK TOPLAM EN YÜKSEK YAĞIŞ (mm)"
          _hdr[2].text = "TARİH"

          # Başlık satırı biçimi.
          for _cell in _hdr:
              _cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
              _tcPr = _cell._tc.get_or_add_tcPr()
              _shd = OxmlElement("w:shd")
              _shd.set(qn("w:fill"), "D9E2F3")
              _tcPr.append(_shd)
              for _p in _cell.paragraphs:
                  _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                  for _r in _p.runs:
                      _r.bold = True
                      _r.font.size = Pt(8.5)

          _secili_mgm_il = str(_yr.get("mgm_il", "")).strip()

          # Seçilen ili güvenilir biçimde eşleştir:
          # Türkçe büyük/küçük harf ve olası boşluk farklarından etkilenmesin.
          def _il_karsilastirma_adi(_metin):
              _x = str(_metin or "").strip().replace("İ", "I").replace("ı", "i")
              return _x.casefold()

          _secili_mgm_il_karsilastirma = _il_karsilastirma_adi(_secili_mgm_il)

          for _il, _deger, _tarih, _url in _mgm_81:
              _cells = _mgm_tbl.add_row().cells
              _is_secili_il = (
                  _il_karsilastirma_adi(_il) == _secili_mgm_il_karsilastirma
                  and bool(_secili_mgm_il_karsilastirma)
              )

              _cells[0].text = _il
              _cells[1].text = f"{_deger:.1f}" if _deger is not None else "Veri alınamadı"
              _cells[2].text = _tarih or "-"

              for _cell in _cells:
                  _cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                  for _p in _cell.paragraphs:
                      _p.alignment = (
                          WD_ALIGN_PARAGRAPH.CENTER
                          if _cell is not _cells[0]
                          else WD_ALIGN_PARAGRAPH.LEFT
                      )
                      for _r in _p.runs:
                          _r.font.size = Pt(8.5)

                  # Projede seçilen il satırı sarı renkle vurgulanır.
                  if _is_secili_il:
                      _tcPr = _cell._tc.get_or_add_tcPr()
                      _shd = _tcPr.find(qn("w:shd"))
                      if _shd is None:
                          _shd = OxmlElement("w:shd")
                          _tcPr.append(_shd)
                      _shd.set(qn("w:fill"), "FFF2CC")

                      # Seçilen ilin okunabilirliği için satır yazıları kalın.
                      for _p in _cell.paragraphs:
                          for _r in _p.runs:
                              _r.bold = True

              # Seçilen ilin yanına raporda açık bir işaret de koy.
              if _is_secili_il:
                  _cells[0].text = f"{_il}  ← SEÇİLEN İL"
                  for _p in _cells[0].paragraphs:
                      _p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                      for _r in _p.runs:
                          _r.bold = True

          doc.add_paragraph(
              "Kaynak: Meteoroloji Genel Müdürlüğü (MGM), Resmi İklim İstatistikleri – "
              "İllerimize Ait Genel İstatistiki Veriler. "
              "MGM verilerinin ölçüm periyotları illere göre farklılık gösterebilir."
          )

          # Hesap girdileri, MGM kaynak bilgisinin hemen altında gösterilir.
          doc.add_paragraph(
              f"Toplama alanı: A = {_yr.get('cati_alani', 0):.2f} m²; "
              f"tasarım yağış yüksekliği: P = {_yr.get('yagis', 0):.2f} mm; "
              f"akış katsayısı: C = {_yr.get('akis_katsayisi', 0):.2f}"
          )

          # Önce ham yağış hacmi, ardından ilk yağış ayırma ve iki yeni katsayı
          # uygulanarak sarnıca alınacak net hacim gösterilir.
          _yr_A = float(_yr.get('cati_alani', 0) or 0)
          _yr_P = float(_yr.get('yagis', 0) or 0)
          _yr_C = float(_yr.get('akis_katsayisi', 0) or 0)
          _yr_ilk = float(_yr.get('ilk_yagis_hacmi', 0) or 0)
          _yr_sarnic_orani = float(_yr.get('sarnic_orani', 80) or 0)
          _yr_filtre_etkinlik = float(_yr.get('filtre_etkinlik', 90) or 0)
          _yr_ham_V = float(_yr.get('ham_toplanabilir_m3', _yr_A * _yr_P * _yr_C / 1000.0) or 0)
          _yr_ilk_haric_V = max(0.0, _yr_ham_V - _yr_ilk)
          _yr_V = float(_yr.get('toplanabilir_m3', _yr_ilk_haric_V * _yr_sarnic_orani / 100.0 * _yr_filtre_etkinlik / 100.0) or 0)
          doc.add_paragraph("Toplanabilir yağmur suyu ve sarnıca alınacak net su hesabı:")
          doc.add_paragraph("V_ham = A × P × C / 1000")
          doc.add_paragraph(
              f"V_ham = {_yr_A:.2f} × {_yr_P:.2f} × {_yr_C:.2f} / 1000 = {_yr_ham_V:.2f} m³"
          )
          doc.add_paragraph(
              f"V_sarnıç = (V_ham − V_ilk yağış) × {_yr_sarnic_orani:.0f}/100 × "
              f"{_yr_filtre_etkinlik:.0f}/100"
          )
          doc.add_paragraph(
              f"V_sarnıç = ({_yr_ham_V:.2f} − {_yr_ilk:.2f}) × {_yr_sarnic_orani:.0f}/100 × "
              f"{_yr_filtre_etkinlik:.0f}/100 = {_yr_V:.2f} m³"
          )
          doc.add_paragraph(
              f"İlk yağış hariç sarnıca alınacak yağmur suyu oranı: %{_yr_sarnic_orani:.0f}; "
              f"filtre etkinlik katsayısı: %{_yr_filtre_etkinlik:.0f}."
          )

          doc.add_heading("6.3.1.2.2 YAĞMUR SUYU FİLTRESİ SEÇİMİ", level=4)
          doc.add_paragraph(f"Filtre tipi: {_yr.get('filtre_tipi', '')}")
          doc.add_paragraph(f"Hesaplanan yağış debisi: {_yr.get('debi_m3h', 0):.2f} m³/h")
          doc.add_paragraph(
              f"Filtre seçim debisi: {_yr.get('filtre_debisi', 0):.2f} m³/h "
              f"(emniyet: %{_yr.get('filtre_emniyet', 0):.0f})"
          )

          doc.add_heading("6.3.1.2.3 İLK YAĞIŞ AYIRICI SEÇİMİ", level=4)
          doc.add_paragraph(
              f"İlk yağış ayırma miktarı: {_yr.get('ilk_yagis_l_m2', 0):.2f} L/m²; "
              f"hesaplanan ayırıcı hacmi: {_yr.get('ilk_yagis_hacmi', 0):.2f} m³"
          )

          doc.add_heading("6.3.1.2.4 YAĞMUR SUYU DEPOSU HACİM HESABI", level=4)
          doc.add_paragraph(
              f"Günlük kullanım ihtiyacı: {_yr.get('kullanim_gunluk', 0):.2f} m³/gün; "
              f"depolama süresi: {_yr.get('depolama_gun', 0):.0f} gün"
          )
          doc.add_paragraph(f"Gerekli depo hacmi: {_yr.get('gerekli_depo', 0):.2f} m³")
          doc.add_paragraph(f"Seçilen yağmur suyu deposu hacmi: {_yr.get('secilen_depo', 0):.2f} m³")

          doc.add_heading("6.3.1.2.5 TAŞMA HATTI HESABI", level=4)
          doc.add_paragraph(
              f"Taşma tasarım debisi: {_yr.get('tasma_debisi', 0):.2f} m³/h "
              f"(emniyet: %{_yr.get('tasma_emniyet', 0):.0f}); "
              f"seçilen taşma hattı: DN {_yr.get('tasma_cap', 0)}"
          )

          doc.add_heading("6.3.1.2.6 TAŞMA SİFONU / KOKU KAPANI", level=4)
          doc.add_paragraph(
              "Taşma hattında sifon/koku kapanı kullanılacaktır." if _yr.get("sifon") else
              "Taşma hattında sifon/koku kapanı öngörülmemiştir."
          )
          if _yr.get("kanal_baglanti"):
              doc.add_paragraph("Taşma hattı kanalizasyona bağlanacaktır; geri tepme koruması sağlanacaktır.")
          else:
              doc.add_paragraph("Taşma hattı kanalizasyona bağlanmayacaktır.")
          if _yr.get("geri_tepme"):
              doc.add_paragraph("Geri tepme önleyici düzenek öngörülmüştür.")

          doc.add_heading("6.3.1.2.7 DEPO GİRİŞİ / SAKİN GİRİŞ", level=4)
          doc.add_paragraph(
              "Depo girişinde sakin giriş düzeni kullanılacaktır."
              if _yr.get("sakin_giris") else
              "Depo girişinde ayrıca sakin giriş düzeni öngörülmemiştir."
          )

          doc.add_heading("6.3.1.2.8 HAVALANDIRMA VE HAŞERE KORUMASI", level=4)
          doc.add_paragraph(
              "Depo havalandırması yapılacaktır." if _yr.get("havalandirma") else
              "Depo havalandırması ayrıca belirtilmemiştir."
          )
          doc.add_paragraph(
              "Havalandırma açıklıkları haşere girişine karşı korunacaktır." if _yr.get("hasere") else
              "Havalandırma açıklıkları için ayrıca haşere koruması belirtilmemiştir."
          )

          doc.add_paragraph(
              "Not: Yağmur suyu pompası bu bölümde seçilmemiştir. Gerekli pompa/hidrofor seçimi "
              "ilgili hidrofor-pompa seçim modülünde yapılacaktır."
          )

        # 6.3.2 başlığı, 6.3.1 bölümünün tamamından sonra eklenir.

        # 6.3.2 başlığı, 6.3.1 bölümünün tamamından sonra eklenir.
        if bolum_632_aktif:
          doc.add_heading(_63_dinamik_baslik("rapor_bolum_632"), level=2)

          genel_bilgiler_basligi = doc.add_paragraph()
          genel_bilgiler_run = genel_bilgiler_basligi.add_run(
              "• Genel Bilgiler ve Hidrofor Seçim Esasları"
          )
          genel_bilgiler_run.bold = False
          genel_bilgiler_run.italic = True
          genel_bilgiler_run.font.size = Pt(12)
          genel_bilgiler_run.font.color.rgb = RGBColor(68, 114, 196)

          hidrofor_genel_secimler_rapor = locals().get(
              "hidrofor_genel_secimler", [True] * 15
          )
          hidrofor_genel_maddeleri_rapor = locals().get(
              "hidrofor_genel_maddeleri", []
          )
          for secili, madde in zip(
              hidrofor_genel_secimler_rapor, hidrofor_genel_maddeleri_rapor
          ):
              if secili:
                  doc.add_paragraph(madde, style="List Bullet")

          # Yalnızca seçilen hidroforlar, boşluk bırakmadan sıralanır.
          for hesap in locals().get("hidrofor_hesaplari", []):
              i = hesap["index"]
              ad = hesap["baslik"]
              doc.add_heading(f"6.3.{_63_dinamik_no("rapor_bolum_632")}.{i} {ad}", level=3)

                  # Alt başlık: Genel Bilgiler başlığı ile aynı görünüm
              if hesap.get("hesaplama_aktif", True):
                hidrofor_hesap_basligi = doc.add_paragraph()
                hidrofor_hesap_basligi.paragraph_format.space_before = Pt(4)
                hidrofor_hesap_basligi.paragraph_format.space_after = Pt(4)
                hidrofor_hesap_run = hidrofor_hesap_basligi.add_run(
                    "• Hidrofor Hesaplamaları"
                )
                hidrofor_hesap_run.bold = False
                hidrofor_hesap_run.italic = True
                hidrofor_hesap_run.font.name = "Arial"
                hidrofor_hesap_run.font.size = Pt(12)
                hidrofor_hesap_run.font.color.rgb = RGBColor(68, 114, 196)
                doc.add_paragraph(
                    f"Toplam yükleme birimi: Z = {hesap['toplam_yb']:.0f} YB"
                )
                doc.add_paragraph(
                    "Saatlik maksimum su tüketimi: Vm = 3600 × (0,25 × √Z)"
                )
                doc.add_paragraph(
                    f"Vm = 3600 × (0,25 × √{hesap['toplam_yb']:.0f}) = "
                    f"{hesap['vm_lph']:.0f} L/h = {hesap['vm_m3h']:.2f} m³/h"
                )
                doc.add_paragraph(
                    f"Pompa debisi: Vp = {hesap['vm_m3h']:.2f} × "
                    f"(1 + %{hesap['emniyet_orani']:.0f}) = {hesap['vp_m3h']:.2f} m³/h"
                )
                doc.add_paragraph(
                    f"Asıl pompa adedi: {hesap['asil_pompa']} adet; "
                    f"bir pompa debisi: {hesap['vp_pompa_m3h']:.2f} m³/h"
                )

                doc.add_paragraph("İşletme basınçları:")
                doc.add_paragraph(f"Kot Farkı: hp = {hesap['hp']:.0f} mSS")
                doc.add_paragraph(f"Akma Basıncı: ha = {hesap['ha']:.0f} mSS")
                doc.add_paragraph(f"Boru Kayıpları: hb = {hesap['hb']:.1f} mSS")
                doc.add_paragraph(f"Sayaç Kayıpları: hc = {hesap['hc']:.0f} mSS")
                doc.add_paragraph("İşletme alt basıncı: Pa = (hp + ha + hb+ hc )")
                doc.add_paragraph(
                    f"İşletme alt basıncı: Pa = ( {hesap['hp']:.0f} + {hesap['ha']:.0f} + "
                    f"{hesap['hb']:.1f} + {hesap['hc']:.0f}) = {hesap['p_alt_mss']:.1f} mSS"
                )
                doc.add_paragraph(
                    f"İşletme alt basıncı: Pa = {hesap['p_alt_atu']:.1f} atü seçildi."
                )
                doc.add_paragraph(
                    f"İşletme üst basıncı: Pu = {hesap['p_ust_atu']:.1f} atü seçildi."
                )
              if hesap.get("tank_hesabi_aktif", True):
                rapor_hidrofor_alt_basligi_ekle(doc, "Hidrofor Tankı Hesabı")
                doc.add_paragraph("VN : Hidrofor tankı nominal hacmi (m3)")
                doc.add_paragraph("QP : Bir pompanın PALT basınçta verdiği max debi miktarı (m3 / h)")
                doc.add_paragraph("S : Şalt sayısı (Motorun saatte devreye girip çıkma sayısı) 1/S")
                doc.add_paragraph("- 2 veya 3 kW lık motor güçlerine kadar şalt sayısı 40'a kadar çıkabilir.")
                doc.add_paragraph("- Büyük motorlarda şalt sayısı 20'ye çekildi.")
                doc.add_paragraph(
                    "VN = 0,33 × QP × (PÜST + 1) / ((PÜST − PALT) × S)"
                )
                doc.add_paragraph(
                    f"VN = 0,33 × {hesap['vp_pompa_m3h']:.2f} × "
                    f"({hesap['p_ust_atu']:.2f} + 1) / "
                    f"(({hesap['p_ust_atu']:.2f} − {hesap['p_alt_atu']:.2f}) × {hesap['schalt']:.0f}) = "
                    f"{hesap['vn_m3']:.3f} m³ = {hesap['vn_m3']*1000:.0f} L"
                )
                doc.add_paragraph(
                    f"Her tank için gerekli hacim: {hesap.get('tank_birim_gerekli_litre', 0):.0f} L"
                )
                doc.add_paragraph(
                    f"Seçilen tank: {hesap['tank_adedi']} × {hesap['tank_birim_litre']:.0f} L = "
                    f"{hesap['tank_toplam_litre']:.0f} L"
                )
                if hesap.get("tank_poz"):
                    doc.add_paragraph(f"Genleşme Tankı Cihaz Poz No: {hesap['tank_poz']}")
              if hesap.get("karakteristik_aktif", True):
                rapor_hidrofor_alt_basligi_ekle(doc, "Hidroforun Karakteristikleri")
                doc.add_paragraph(
                    f"Tank hacmi: Vt = {hesap['tank_adedi']} x {hesap['tank_birim_litre']:.0f} lt."
                )
                doc.add_paragraph(
                    f"Pompa debisi: Vp = {hesap['asil_pompa']} ad x {hesap['vp_pompa_m3h']:.0f} m3/h"
                )
                doc.add_paragraph(f"İşletme alt basıncı: H a = {hesap['p_alt_mss']:.0f} mSS")
                doc.add_paragraph(f"İşletme üst basıncı: H ü = {hesap['p_ust_mss']:.0f} mSS")
                doc.add_paragraph(
                    f"Pompa gücü: Np = {hesap['toplam_pompa']} ad x {hesap['guc']:.1f} KW."
                )
                tip_metni = hesap.get("poz_aciklama", "")
                if not tip_metni or tip_metni.startswith("Birim Poziyat") or tip_metni.startswith("Seçilen debi"):
                    tip_metni = HIDROFOR_TIP_METINLERI.get(hesap['toplam_pompa'], "Hidrofor")
                doc.add_paragraph(f"Tipi: {tip_metni}")
                if hesap.get("poz_rapora_eklensin_mi", True):
                    if hesap.get("poz"):
                        doc.add_paragraph(f"Cihaz Poz No: {hesap['poz']}")
                    else:
                        doc.add_paragraph(
                            "Cihaz Poz No: Birim Poziyat verisi bu pompa adedi için henüz tanımlanmamıştır."
                        )
              if hesap.get("egrisi_aktif", True):
                curve = hesap.get("pompa_curve", [])
                if curve:
                    q_curve = [p[0] for p in curve]
                    h_curve = [p[1] for p in curve]
                    grafik_buf = pompa_grafigi_png(
                        q_curve, h_curve, hesap["vp_pompa_m3h"], hesap["h"],
                        hesap.get("pompa_egrisi_basligi", "Pompa Performans Eğrisi"),
                        anonim=True,
                    )
                    rapor_hidrofor_alt_basligi_ekle(doc, "Pompa Performans Eğrisi:")
                    doc.add_picture(grafik_buf, width=Inches(6.2))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 6.3.3 Kullanma Sıcak Suyu İhtiyacı Hesapları
    if bolum_633_aktif:
        doc.add_heading(_63_dinamik_baslik("rapor_bolum_633"), level=2)
        if sicak_su_hesap_detaylari:
            yapi_tip_par = doc.add_paragraph()
            yapi_tip_run = yapi_tip_par.add_run(f"Yapı / kullanım tipi: {sicak_su_yapi_tipi}")
            yapi_tip_run.bold = True
            yapi_tip_run.font.color.rgb = RGBColor(0, 0, 0)
            t_sicak = doc.add_table(rows=1, cols=5)
            t_sicak.style = "Table Grid"
            basliklar = ["Kullanım yeri", "Standart aralığı [L]", "Hesapta kullanılan [L]", "Adet", "Toplam [L]"]
            for i, baslik in enumerate(basliklar):
                t_sicak.rows[0].cells[i].text = baslik
            for detay in sicak_su_hesap_detaylari:
                hucre = t_sicak.add_row().cells
                hucre[0].text = detay["kullanim"]
                hucre[1].text = str(detay["kaynak_aralik"])
                hucre[2].text = f"{detay['birim_degeri']:.0f}"
                hucre[3].text = f"{detay['miktar']:.0f}"
                hucre[4].text = f"{detay['toplam_litre']:.0f}"
            toplam = t_sicak.add_row().cells
            toplam[0].text = "GENEL TOPLAM"
            toplam[4].text = f"{sicak_su_gunluk_toplam_litre:,.0f} L/gün".replace(",", ".")

            # Tablonun bir satır altında Ortalama Ani Sıcak Su İhtiyacı başlığı
            doc.add_paragraph("")
            ortalama_ani_baslik = doc.add_paragraph()
            run = ortalama_ani_baslik.add_run("Ortalama Ani Sıcak Su İhtiyacı:")
            run.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)

            faktor_par = doc.add_paragraph()
            faktor_par.add_run("Kullanma Eş Zaman Faktörü = ")
            run = faktor_par.add_run(f"{kullanma_es_faktoru:.2f}")
            run.bold = True

            faktor_par = doc.add_paragraph()
            faktor_par.add_run("Depolama Faktörü = ")
            run = faktor_par.add_run(f"{depolama_faktoru:.2f}")
            run.bold = True

            toplam_litre_fmt = str(int(round(sicak_su_gunluk_toplam_litre)))
            hesaplanan_fmt = str(int(round(hesaplanan_boyler_hacmi)))
            formula_par = doc.add_paragraph()
            formula_par.add_run(
                f"V = {depolama_faktoru:.2f} × {kullanma_es_faktoru:.2f} × {toplam_litre_fmt} = "
            )
            run = formula_par.add_run(f"{hesaplanan_fmt} L")
            run.bold = True

            emniyet_par = doc.add_paragraph()
            emniyet_par.add_run("V = ")
            run = emniyet_par.add_run(f"{secilen_boyler_hacmi} L")
            run.bold = True
            emniyet_par.add_run(" (Emniyetle)")
            if es_zaman_faktoru is not None:
                doc.add_paragraph(f"Excel eş zaman faktörü ({konut_sayisi} konut): {es_zaman_faktoru:.2f}")

            # Boyler / eşanjör ısıtma kapasite hesabı başlığı
            # Seçilen sisteme göre raporda da otomatik değişir.
            _rapor_hesap_tip = str(
                st.session_state.get("boyler_secili_tip_v57", "TEK SERPANTİNLİ BOYLER")
            )
            _rapor_hesap_basligi = (
                "PLAKALI EŞANJÖR ISITMA KAPASİTE HESABI:"
                if _rapor_hesap_tip == "PLAKALI EŞANJÖR"
                else "BOYLER ISITMA KAPASİTE HESABI:"
            )
            boyler_baslik = doc.add_paragraph()
            boyler_baslik_run = boyler_baslik.add_run(_rapor_hesap_basligi)
            boyler_baslik_run.bold = True
            boyler_baslik_run.font.color.rgb = RGBColor(0, 0, 0)

            doc.add_paragraph("Q = ms × c × Δts = mp × c × Δtp")
            doc.add_paragraph("Q : İletilmesi gereken toplam ısı miktarı. (kCal/h)")
            doc.add_paragraph("ms : Sekonder su debisi (lt/h)")
            doc.add_paragraph("mp : Primer su debisi (lt/h)")
            doc.add_paragraph("c : Suyun özgül ısısı. (°C) [1 kCal/kg.°C]")
            doc.add_paragraph("Δts : Sekonder devre giriş-çıkış suyu sıcaklık farkı. (°C)")
            doc.add_paragraph("Δtp : Primer devre giriş-çıkış suyu sıcaklık farkı. (°C)")
            doc.add_paragraph(
                f"Q = {boyler_ms:.0f} lt/h × 1 kCal/kg.°C × "
                f"({boyler_ts_cikis:.0f}-{boyler_ts_giris:.0f}) °C"
            )
            q_par = doc.add_paragraph()
            q_par.add_run(
                f"Q = {boyler_q_kcal_h:.0f} kcal/h ≈ {int(boyler_q_kw)} kW"
            ).bold = True

            # Boyler / eşanjör seçimi
            # Kullanıcının istediği rapor formatı:
            # Q, primer/sekonder rejim, adet, tip, hacim, debi ve Cihaz Poz No.
            secili_tip = st.session_state.get(
                "boyler_secili_tip_v57",
                "TEK SERPANTİNLİ BOYLER",
            )

            # Rapor oluşturulurken de poz seçimini yeniden doğrula. Böylece
            # Streamlit oturumunda seçim sonucu henüz oluşmamış olsa bile
            # Cihaz Poz No rapora mutlaka aktarılır.
            boyler_secim_sonucu = st.session_state.get(
                "boyler_secim_sonucu_v59",
                None,
            )

            if secili_tip == "TEK SERPANTİNLİ BOYLER":
                rapor_adet = max(1, int(st.session_state.get("boyler_adet_tek_serpantin_v59", 3)))
                if isinstance(boyler_secim_sonucu, dict) and int(boyler_secim_sonucu.get("adet", 0)) == rapor_adet and boyler_secim_sonucu.get("poz"):
                    _rapor_poz_kaydi = boyler_secim_sonucu
                else:
                    _rapor_poz_kaydi, _, _, _ = _boyler_poz_sec(secilen_boyler_hacmi, boyler_ms, rapor_adet)
                rapor_poz = _rapor_poz_kaydi["poz"]
                rapor_hacim = int(_rapor_poz_kaydi["hacim"])
                rapor_debi = int(_rapor_poz_kaydi.get("debi", _rapor_poz_kaydi.get("debi_80_60", 0)))
                rapor_tip = (
                    "Tek Bakır Boru Serpantinli, Dik Tip, Gövdesi İzolasyonlu, "
                    "Elektrostatik Toz Boyalı ( TS EN 13445-3, TS EN 12897, TS 736 )"
                )
                rapor_gunes_satiri = False
                rapor_alt_debi = rapor_ust_debi = 0
                rapor_gunes_rejimi = ""
            elif secili_tip == "ÇİFT SERPANTİNLİ BOYLER":
                rapor_adet = max(1, int(st.session_state.get("boyler_adet_cift_serpantin_v59", 3)))
                _cift_kayit = st.session_state.get("boyler_cift_secim_sonucu_v66")
                if not (isinstance(_cift_kayit, dict) and int(_cift_kayit.get("adet", 0)) == rapor_adet and _cift_kayit.get("poz")):
                    _cift_poz, _, _, _ = _cift_boyler_poz_sec(secilen_boyler_hacmi, boyler_ms, rapor_adet)
                    _cift_kayit = {"poz":_cift_poz["poz"], "hacim":_cift_poz["hacim"], "alt_debi":_cift_poz["alt_debi"], "ust_debi":_cift_poz["ust_debi"], "adet":rapor_adet, "gunes_rejimi":st.session_state.get("boyler_gunes_rejimi_v66", "60/40")}
                rapor_poz = _cift_kayit["poz"]
                rapor_hacim = int(_cift_kayit["hacim"])
                rapor_debi = int(_cift_kayit["alt_debi"] + _cift_kayit["ust_debi"])
                rapor_alt_debi = int(_cift_kayit["alt_debi"])
                rapor_ust_debi = int(_cift_kayit["ust_debi"])
                rapor_gunes_rejimi = str(
                    st.session_state.get(
                        "rej_gunes_ist_v68",
                        _cift_kayit.get("gunes_rejimi", "60/40"),
                    )
                )
                rapor_gunes_satiri = True
                rapor_tip = (
                    "Çift Bakır Boru Serpantinli, Dik Tip, Gövdesi İzolasyonlu, "
                    "Elektrostatik Toz Boyalı ( TS EN 13445-3, TS EN 12897, TS 736 )"
                )
            else:
                rapor_adet = max(1, int(st.session_state.get("boyler_adet_plakali_esanjör_v59", 2)))
                _plaka_kayit = st.session_state.get("plakali_esanjör_secim_sonucu_v75")
                if not (isinstance(_plaka_kayit, dict) and _plaka_kayit.get("poz")):
                    _plaka_q_birim_rapor = float(boyler_q_kcal_h) / float(max(1, rapor_adet))
                    _plaka_poz_rapor = _plakali_esanjör_poz_sec(_plaka_q_birim_rapor)
                    _plaka_kayit = {
                        "poz": _plaka_poz_rapor["poz"],
                        "q_kcal_h": _plaka_poz_rapor["q_kcal_h"],
                        "q_kw": _plaka_poz_rapor["q_kcal_h"] * 0.001163,
                        "primer_dp_mss": _plaka_poz_rapor["primer_dp_mss"],
                        "sekonder_dp_mss": float(_plaka_poz_rapor["primer_dp_mss"]),
                        "calisma_adet": rapor_adet,
                        "yedek_adet": 1,
                        "adet": rapor_adet + 1,
                    }
                rapor_poz = _plaka_kayit["poz"]
                rapor_hacim = int(round(secilen_boyler_hacmi / rapor_adet))
                rapor_debi = int(round(boyler_ms / rapor_adet))
                rapor_tip = "Plakalı eşanjör"
                rapor_gunes_satiri = False
                rapor_alt_debi = rapor_ust_debi = 0
                rapor_gunes_rejimi = ""

            # İstenen rapor formatı: etiket / iki nokta / değer şeklinde
            # üç sütunlu, sabit genişlikli ve hizalı tablo. Böylece etiket
            # uzunlukları değişse bile iki nokta üst üste ve değerler kaymaz.
            doc.add_paragraph("")
            boyler_rapor_baslik = doc.add_paragraph()
            boyler_rapor_baslik.paragraph_format.space_after = Pt(4)
            boyler_rapor_baslik_run = boyler_rapor_baslik.add_run(
                "Kullanma Sıcak Suyu Boyleri:" if secili_tip != "PLAKALI EŞANJÖR" else ""
            )
            boyler_rapor_baslik_run.bold = True
            boyler_rapor_baslik_run.font.color.rgb = RGBColor(0, 0, 0)
            boyler_rapor_baslik_run.font.name = "Times New Roman"
            boyler_rapor_baslik_run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
            boyler_rapor_baslik_run.font.size = Pt(11)
            boyler_rapor_baslik_run.font.bold = True
            boyler_rapor_baslik_run.font.italic = False
            boyler_rapor_baslik_run.font.color.rgb = RGBColor(0, 0, 0)

            # Q BOYLER burada kullanıcı tarafından elle değiştirilmiş nihai tam sayıdır.
            # Rapor her zaman bu değeri kullanır; küsüratlı hesap ayrıca yukarıda gösterilir.
            if secili_tip == "PLAKALI EŞANJÖR":
                _plaka_rapor = st.session_state.get("plakali_esanjör_secim_sonucu_v75", {})
                _plaka_poz_rapor = _plaka_rapor.get("poz", rapor_poz)
                # Rapor oluşturulurken çalışma adedi doğrudan kayıtlı seçimden
                # alınır; böylece önceki oturumdan kalan/olmayan bir değişken
                # nedeniyle NameError oluşmaz. Yedek eşanjör kapasite hesabına
                # hiçbir zaman dahil edilmez.
                _plaka_calisma_adet_rapor = max(
                    1,
                    int(_plaka_rapor.get("calisma_adet", rapor_adet - 1 if rapor_adet > 1 else rapor_adet))
                )
                _plaka_q_birim_rapor = float(boyler_q_kcal_h) / float(_plaka_calisma_adet_rapor)
                _plaka_q_rapor = int(_plaka_rapor.get("q_kcal_h", round(_plaka_q_birim_rapor)))
                _plaka_kw_rapor = _plaka_rapor.get("q_kw", _plaka_q_rapor * 0.001163)
                _plaka_primer_dp = float(_plaka_rapor.get("primer_dp_mss", 4.0))
                _plaka_sekonder_dp = _plaka_primer_dp

                _akum_rapor = st.session_state.get("plakali_akumulasyon_secim_sonucu_v75", {})
                _akum_katsayi_rapor = float(st.session_state.get("plakali_akumulasyon_katsayisi_v75", 1.0))
                _akum_gerekli_rapor = float(_akum_rapor.get("gerekli_hacim", toplam_tuketim * _akum_katsayi_rapor))
                _akum_hacim_rapor = int(_akum_rapor.get("hacim", 100))
                _akum_adet_rapor = int(_akum_rapor.get("adet", 1))
                _akum_toplam_rapor = int(_akum_rapor.get("toplam_hacim", _akum_adet_rapor * _akum_hacim_rapor))
                _akum_tip_rapor = _akum_rapor.get("tip", AKUMULASYON_TANKI_TIP)

                _plaka_calisma_adet_rapor = int(_plaka_rapor.get("calisma_adet", max(1, rapor_adet - 1)))
                _plaka_yedek_adet_rapor = int(_plaka_rapor.get("yedek_adet", 1))
                _plaka_toplam_adet_rapor = int(_plaka_rapor.get("adet", _plaka_calisma_adet_rapor + _plaka_yedek_adet_rapor))

                # 1. ekipman: Akümülasyon tankı
                # Önceki boş paragraf kaldırıldı; başlık Q hesabına daha yakın konumlanır.
                akum_baslik = doc.add_paragraph()
                akum_baslik_run = akum_baslik.add_run("SICAK SU AKÜMÜLASYON TANKI SEÇİMİ:")
                akum_baslik_run.bold = True
                akum_baslik_run.font.color.rgb = RGBColor(0, 0, 0)
                akum_baslik_run.font.name = "Arial"
                akum_baslik_run.font.size = Pt(11)
                boyler_rapor_satirlari = [
                    ("Tank hacmi", f"{_akum_hacim_rapor} L (PN 10)"),
                    ("Adet", str(_akum_adet_rapor)),
                    ("Tip", _akum_tip_rapor),
                ]
                if bool(st.session_state.get("boyler_poz_rapora_eklensin_v72", True)):
                    boyler_rapor_satirlari.append(("Cihaz Poz No", _akum_rapor.get("poz", "")))

                # 2. ekipman: Plakalı eşanjör
                # Raporda ayrı, tam genişlikte bir seçim başlığı olarak gösterilir.
                boyler_rapor_satirlari.extend([
                    ("__PLAKALI_ESANJOR_BASLIK__", ""),
                    ("Q", f"{_plaka_q_rapor:.0f} kcal/h ≈ {int(_plaka_kw_rapor)} kW (eşanjör başına)"),
                    ("Primer Devre", "80 / 60 °C sıcak su (Kazan)"),
                    ("Seconder Devre", "10 / 60 °C sıcak su"),
                    ("Primer Devre Basınç Kaybı", f"{_plaka_primer_dp:g} mSS"),
                    ("Seconder Devre Basınç Kaybı", f"{_plaka_sekonder_dp:g} mSS"),
                    ("Tip", "Plakalı eşanjör."),
                    ("Adet", f"{_plaka_toplam_adet_rapor} ({_plaka_yedek_adet_rapor} adet Yedek)"),
                ])
                if bool(st.session_state.get("boyler_poz_rapora_eklensin_v72", True)):
                    boyler_rapor_satirlari.append(("Cihaz Poz No", _plaka_poz_rapor))
            else:
                boyler_rapor_satirlari = [
                    ("Q BOYLER", f"{int(boyler_q_kw)} kW"),
                    (
                        "Isıtıcı Akışkan (Kazan)",
                        f"{boyler_tp_giris:.0f}/{boyler_tp_cikis:.0f} ºC sıcak su (4,0 mSS, basınç kaybı) (Kabul)",
                    ),
                ]
                if rapor_gunes_satiri:
                    boyler_rapor_satirlari.append((
                        "Isıtıcı Akışkan (Güneş)",
                        f"{rapor_gunes_rejimi.replace('/', '/')} ºC sıcak su (4,0 mSS, basınç kaybı) (Kabul)",
                    ))
                boyler_rapor_satirlari.extend([
                    ("Isıtılan Akışkan", f"{boyler_ts_cikis:.0f}/{boyler_ts_giris:.0f} ºC sıcak su (4,0 mSS, basınç kaybı) (Kabul)"),
                    ("Adet", str(rapor_adet)),
                    ("Tip", rapor_tip),
                    ("Boyler hacmi", f"{rapor_hacim} lt"),
                ])
                if rapor_gunes_satiri:
                    # Çift serpantinli boylerde iki serpantin debisi ayrı ayrı rapora aktarılır.
                    boyler_rapor_satirlari.extend([
                        ("Alt serpantin debisi", f"{rapor_alt_debi} lt/h"),
                        ("Üst serpantin debisi", f"{rapor_ust_debi} lt/h"),
                    ])
                else:
                    # Tek serpantinli boylerde tek debi gösterilir.
                    boyler_rapor_satirlari.append(("Boyler Debisi", f"{rapor_debi} lt/h"))
                if bool(st.session_state.get("boyler_poz_rapora_eklensin_v72", True)):
                    boyler_rapor_satirlari.append(("Cihaz Poz No", rapor_poz))

            # Word tablosu kullanıyoruz; kenarlıkları kaldırarak düz metin
            # görünümü korunur, ancak üç kolon sayesinde tüm satırlar simetrik
            # ve profesyonel biçimde hizalanır.
            boyler_tablo = doc.add_table(
                rows=len(boyler_rapor_satirlari), cols=3
            )
            boyler_tablo.autofit = False
            boyler_tablo.allow_autofit = False

            # Sayfa genişliğine göre: etiket 2.15", iki nokta 0.18", değer kalan alan.
            kolon_genislikleri = [Inches(2.15), Inches(0.18), Inches(4.65)]
            kalin_etiketler = {
                "Q BOYLER", "Adet", "Boyler hacmi",
                "Boyler Debisi", "Alt serpantin debisi", "Üst serpantin debisi", "Cihaz Poz No", "Q", "Primer Devre", "Seconder Devre", "Primer Devre Basınç Kaybı", "Seconder Devre Basınç Kaybı", "SICAK SU AKÜMÜLASYON TANKI", "Gerekli akümülasyon hacmi", "Tank hacmi"
            }

            for satir_no, (etiket, deger) in enumerate(boyler_rapor_satirlari):
                hucreler = boyler_tablo.rows[satir_no].cells

                # Plakalı eşanjör seçim başlığı: üç hücre birleştirilir,
                # böylece başlık sayfa genişliğine yayılır.
                if etiket == "__PLAKALI_ESANJOR_BASLIK__":
                    hucre = hucreler[0].merge(hucreler[1]).merge(hucreler[2])
                    hucre.width = sum(kolon_genislikleri)
                    hucre.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                    tcPr = hucre._tc.get_or_add_tcPr()
                    tcBorders = tcPr.first_child_found_in("w:tcBorders")
                    if tcBorders is None:
                        tcBorders = OxmlElement("w:tcBorders")
                        tcPr.append(tcBorders)
                    for kenar in ("top", "left", "bottom", "right", "insideH", "insideV"):
                        el = tcBorders.find(qn(f"w:{kenar}"))
                        if el is None:
                            el = OxmlElement(f"w:{kenar}")
                            tcBorders.append(el)
                        el.set(qn("w:val"), "nil")

                    p_baslik = hucre.paragraphs[0]
                    p_baslik.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p_baslik.paragraph_format.space_before = Pt(12)
                    p_baslik.paragraph_format.space_after = Pt(8)
                    run_baslik = p_baslik.add_run(
                        "KULLANMA SICAK SU SİSTEMİ PLAKALI EŞANJÖRÜ SEÇİMİ:"
                    )
                    run_baslik.font.name = "Times New Roman"
                    run_baslik._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
                    run_baslik.font.size = Pt(12)
                    run_baslik.bold = True
                    run_baslik.font.italic = False
                    run_baslik.font.color.rgb = RGBColor(0, 0, 0)
                    continue

                for hucre, genislik in zip(hucreler, kolon_genislikleri):
                    hucre.width = genislik
                    hucre.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    # Tablo kenarlıklarını tamamen kaldır.
                    tcPr = hucre._tc.get_or_add_tcPr()
                    tcBorders = tcPr.first_child_found_in("w:tcBorders")
                    if tcBorders is None:
                        tcBorders = OxmlElement("w:tcBorders")
                        tcPr.append(tcBorders)
                    for kenar in ("top", "left", "bottom", "right", "insideH", "insideV"):
                        el = tcBorders.find(qn(f"w:{kenar}"))
                        if el is None:
                            el = OxmlElement(f"w:{kenar}")
                            tcBorders.append(el)
                        el.set(qn("w:val"), "nil")

                p_etiket = hucreler[0].paragraphs[0]
                p_etiket.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_etiket.paragraph_format.space_after = Pt(3)
                p_etiket.paragraph_format.space_before = Pt(3)
                run_etiket = p_etiket.add_run(etiket)
                run_etiket.font.name = "Times New Roman"
                run_etiket._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
                run_etiket.font.size = Pt(10.5)
                run_etiket.font.bold = False
                run_etiket.font.italic = False
                run_etiket.font.color.rgb = RGBColor(0, 0, 0)

                p_iki_nokta = hucreler[1].paragraphs[0]
                p_iki_nokta.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_iki_nokta.paragraph_format.space_after = Pt(3)
                p_iki_nokta.paragraph_format.space_before = Pt(3)
                run_iki_nokta = p_iki_nokta.add_run(":")
                run_iki_nokta.font.name = "Times New Roman"
                run_iki_nokta._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
                run_iki_nokta.font.size = Pt(10.5)
                run_iki_nokta.font.bold = False
                run_iki_nokta.font.italic = False
                run_iki_nokta.font.color.rgb = RGBColor(0, 0, 0)

                p_deger = hucreler[2].paragraphs[0]
                p_deger.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_deger.paragraph_format.space_after = Pt(3)
                p_deger.paragraph_format.space_before = Pt(3)
                run_deger = p_deger.add_run(deger)
                run_deger.font.name = "Times New Roman"
                run_deger._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
                run_deger.font.size = Pt(10.5)
                run_deger.font.color.rgb = RGBColor(0, 0, 0)
                run_deger.bold = False
                run_deger.italic = False

            doc.add_paragraph("").paragraph_format.space_after = Pt(0)

        else:
            doc.add_paragraph("Herhangi bir kullanma sıcak suyu kullanım yeri seçilmemiştir.")

    # --- 6.3.4 KULLANMA SICAK SU TESİSATI RE-SİRKULASYON POMPASI SEÇİMİ ---
    if bolum_634_aktif:
        _rs_rapor = st.session_state.get("re_sirkulasyon_pompa_sonucu_v99", {})
        if _rs_rapor:
            doc.add_heading(_63_dinamik_baslik("rapor_bolum_634"), level=2)

            _rs_qb = float(_rs_rapor.get("q_boyler_kcal_h", 0.0))
            _rs_qb_kw = float(_rs_rapor.get("q_boyler_kw", _rs_qb * 0.001163))
            _rs_qh = float(_rs_rapor.get("q_hesap_m3h", 0.0))
            _rs_q_txt = f"{_rs_qb:,.0f}"
            _rs_kw_txt = f"{_rs_qb_kw:.0f}".replace(".", ",")
            _rs_factor = 1.0 + float(_rs_rapor.get("emniyet_orani", 15.0)) / 100.0
            _rs_factor_txt = f"{_rs_factor:.2f}".replace(".", ",")
            _rs_qh_txt = f"{_rs_qh:.2f}"

            # TÜM 6.3.4 ÇIKTISINI TEK BİR ÇERÇEVESİZ HÜCREDE TOPLA.
            # Böylece Word'deki ayrı paragraf/table akışlarından kaynaklanan
            # büyük boşluklar oluşmaz; formül ve pompa bilgileri tek blok halinde kalır.
            _rs_tbl = doc.add_table(rows=1, cols=1)
            _rs_tbl.autofit = False
            _rs_tbl.allow_autofit = False
            _rs_cell = _rs_tbl.cell(0, 0)
            _rs_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            _rs_tbl.columns[0].width = Inches(6.9)

            # Tablo kenarlıklarını tamamen kaldır.
            _tbl_pr = _rs_tbl._tbl.tblPr
            _tbl_borders = _tbl_pr.first_child_found_in("w:tblBorders")
            if _tbl_borders is None:
                _tbl_borders = OxmlElement("w:tblBorders")
                _tbl_pr.append(_tbl_borders)
            for _edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                _el = _tbl_borders.find(qn(f"w:{_edge}"))
                if _el is None:
                    _el = OxmlElement(f"w:{_edge}")
                    _tbl_borders.append(_el)
                _el.set(qn("w:val"), "nil")

            def _rs_cell_paragraph(alignment=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0):
                _p = _rs_cell.add_paragraph()
                _p.alignment = alignment
                _p.paragraph_format.space_before = Pt(before)
                _p.paragraph_format.space_after = Pt(after)
                _p.paragraph_format.line_spacing = 1.0
                return _p

            # İlk boş paragrafı temizle.
            _rs_cell.paragraphs[0].text = ""
            _rs_cell.paragraphs[0].paragraph_format.space_before = Pt(0)
            _rs_cell.paragraphs[0].paragraph_format.space_after = Pt(0)

            # Q BOYLER — sayfanın en soluna yaslı.
            _qpar = _rs_cell_paragraph(WD_ALIGN_PARAGRAPH.LEFT, 0, 5)
            _qpar.paragraph_format.left_indent = Inches(0)
            _qpar.paragraph_format.first_line_indent = Inches(0)
            _qr = _qpar.add_run(f"Q BOYLER = {_rs_q_txt} kcal/h ≈ {_rs_kw_txt} kW")
            _qr.font.name = "Arial"
            _qr.font.size = Pt(11.5)
            _qr.bold = True

            # Q BOYLER ile formül arasında yaklaşık bir Enter mesafesi.
            _formula_spacer = _rs_cell_paragraph(WD_ALIGN_PARAGRAPH.LEFT, 0, 0)
            _formula_spacer.paragraph_format.line_spacing = 1.0
            _formula_spacer.add_run(" ").font.size = Pt(10.5)

            # Formül: Word'de dağılmaması için üç sabit kolon kullanılır.
            # V= sola yaslı, pay ve payda mevcut örneğe göre bir karakter sola alınır.
            _ft = _rs_cell.add_table(rows=1, cols=3)
            _ft.autofit = False
            _ft.alignment = WD_TABLE_ALIGNMENT.LEFT
            _ft.columns[0].width = Inches(0.55)
            _ft.columns[1].width = Inches(3.00)
            _ft.columns[2].width = Inches(2.70)
            _fc = _ft.rows[0].cells
            for _c in _fc:
                _c.width = Inches(0.55) if _c is _fc[0] else (Inches(3.00) if _c is _fc[1] else Inches(2.70))
                _c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                _tcPr = _c._tc.get_or_add_tcPr()
                _b = _tcPr.first_child_found_in("w:tcBorders")
                if _b is None:
                    _b = OxmlElement("w:tcBorders"); _tcPr.append(_b)
                for _edge in ("top","left","bottom","right","insideH","insideV"):
                    _el = _b.find(qn(f"w:{_edge}"))
                    if _el is None:
                        _el = OxmlElement(f"w:{_edge}"); _b.append(_el)
                    _el.set(qn("w:val"), "nil")
            _p0 = _fc[0].paragraphs[0]; _p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _p0.paragraph_format.space_before = Pt(0); _p0.paragraph_format.space_after = Pt(0)
            _rv = _p0.add_run("V =")
            _rv.font.name="Courier New"; _rv.font.size=Pt(10.5); _rv.bold=False

            # Orta hücrede pay / çizgi / payda. Pay ve payda bir karakter sola çekildi.
            _pf = _fc[1].paragraphs[0]
            _pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _pf.paragraph_format.space_before = Pt(0); _pf.paragraph_format.space_after = Pt(0)
            _rf = _pf.add_run(
                f"{_rs_q_txt} × 0,05 × {_rs_factor_txt}\n"
                f"-------------------------------\n"
                f"          5.000"
            )
            _rf.font.name="Courier New"; _rf.font.size=Pt(10.5)

            _pr = _fc[2].paragraphs[0]
            _pr.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _pr.paragraph_format.space_before = Pt(0); _pr.paragraph_format.space_after = Pt(0)
            _rq = _pr.add_run(f"= {_rs_qh_txt} m³/h")
            _rq.font.name="Courier New"; _rq.font.size=Pt(10.5)

            _sp = _rs_cell_paragraph(WD_ALIGN_PARAGRAPH.LEFT, 0, 4)
            _sp.paragraph_format.line_spacing = 1.0

            _baslik = _rs_cell_paragraph(WD_ALIGN_PARAGRAPH.LEFT, 0, 3)
            _br = _baslik.add_run("Seçilen Pompa :")
            _br.font.name = "Arial"
            _br.font.size = Pt(11.5)
            _br.bold = True

            # Etiket, ayraç ve değer tek bir sabit tablo üzerinde hizalanır.
            _pump_tbl = _rs_cell.add_table(rows=0, cols=3)
            _pump_tbl.autofit = False
            _pump_tbl.allow_autofit = False
            _pump_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
            _pump_tbl.columns[0].width = Inches(1.20)
            _pump_tbl.columns[1].width = Inches(0.20)
            _pump_tbl.columns[2].width = Inches(5.50)

            def _pump_row(label, value, bold_value=True, equals=False):
                _cells = _pump_tbl.add_row().cells
                for _c, _w in zip(_cells, (1.20, 0.20, 5.50)):
                    _c.width = Inches(_w)
                    _c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                    _tcPr = _c._tc.get_or_add_tcPr()
                    _b = _tcPr.first_child_found_in("w:tcBorders")
                    if _b is None:
                        _b = OxmlElement("w:tcBorders"); _tcPr.append(_b)
                    for _edge in ("top","left","bottom","right","insideH","insideV"):
                        _el = _b.find(qn(f"w:{_edge}"))
                        if _el is None:
                            _el = OxmlElement(f"w:{_edge}"); _b.append(_el)
                        _el.set(qn("w:val"), "nil")
                p1, p2, p3 = (_cells[0].paragraphs[0], _cells[1].paragraphs[0], _cells[2].paragraphs[0])
                for _p in (p1,p2,p3):
                    _p.paragraph_format.space_before = Pt(0)
                    _p.paragraph_format.space_after = Pt(0)
                    _p.paragraph_format.line_spacing = 1.0
                p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r1 = p1.add_run(label); r1.font.name="Arial"; r1.font.size=Pt(11); r1.bold=False; r1.font.color.rgb = RGBColor(0,0,0)
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r2 = p2.add_run("=" if equals else ":"); r2.font.name="Arial"; r2.font.size=Pt(11); r2.bold=False; r2.font.color.rgb = RGBColor(0,0,0)
                p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r3 = p3.add_run(str(value)); r3.font.name="Arial"; r3.font.size=Pt(11); r3.bold=False; r3.font.color.rgb = RGBColor(0,0,0)

            _pump_row("V", f"{_rs_rapor.get('q_m3h', 0.0):.2f} m³/h")
            _pump_row("H", f"{_rs_rapor.get('h_mss', 0.0):g} mSS")
            _pump_row("Güç", f"{_rs_rapor.get('guc_kw', 0.20):.2f} kW")
            _pump_row("Adet", _rs_rapor.get("adet_str", ""))
            _pump_row("Tip", _rs_rapor.get("tip", ""), bold_value=False)
            if _rs_rapor.get("poz_rapora_aktar", True):
                _pump_row("Cihaz Poz No", _rs_rapor.get("poz", ""), bold_value=True, equals=True)

            # Pompa performans eğrisi rapora eklenir; marka/model adı raporda gösterilmez.
            _rs_curve = _rs_rapor.get("pompa_curve", [])
            if _rs_curve:
                _egri_sp = _rs_cell_paragraph(WD_ALIGN_PARAGRAPH.LEFT, 6, 2)
                _er = _egri_sp.add_run("Pompa Performans Eğrisi")
                _er.font.name = "Arial"; _er.font.size = Pt(11.5); _er.bold = True
                _egri_png = pompa_grafigi_png(
                    [x[0] for x in _rs_curve], [x[1] for x in _rs_curve],
                    float(_rs_rapor.get("q_m3h", 0.0)), float(_rs_rapor.get("h_mss", 0.0)),
                    "Pompa Performans Eğrisi", anonim=True,
                )
                _pic = _rs_cell_paragraph(WD_ALIGN_PARAGRAPH.CENTER, 0, 0)
                _pic.add_run().add_picture(_egri_png, width=Inches(5.8))

            _after = _rs_cell_paragraph(WD_ALIGN_PARAGRAPH.LEFT, 0, 0)
            _after.paragraph_format.line_spacing = 1.0


    # Raporun Word dosyasına dönüştürülmesi ve indirme düğmesinin oluşturulması.
    rapor_word_stillerini_uygula(doc)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    dosya_adi = (
        f"{aktif_is.replace(' ', '_')}_Rapor"
        if is_adi
        else "Mekanik_Uygulama_Raporu"
    )

    st.session_state["_rapor_hazir_docx_v134"] = buffer.getvalue()
    st.session_state["_rapor_hazir_adi_v134"] = dosya_adi


# SAYFA SONU ANKORU
st.markdown('<div id="sayfa_sonu"></div>', unsafe_allow_html=True)
