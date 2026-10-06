from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from datetime import datetime
import io
import base64
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
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from modules.fragment_runner import run_fragment
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
        f"{_base1}?k=&m={_il_url}",
        f"{_base2}?k=&m={_il_url}",
        f"{_base1}?m={_il_url}",
        f"{_base1}?k=undefined&m={_il_url}",
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
                with urllib.request.urlopen(_req, timeout=5) as _response:
                    _html = _response.read().decode("utf-8", errors="ignore")
                _deger, _tarih, _parsed_url = _parse_html(_html, _url)
                if _deger is not None:
                    return _deger, _tarih, _parsed_url
            except Exception:
                pass
            if _deneme < 2:
                _time.sleep(0.7 * (_deneme + 1))

    return None, None, _son_url


@st.cache_data(ttl=86400, show_spinner=False)
def mgm_aylik_ortalama_yagis_mm(il_adi):
    """MGM resmi iklim istatistiklerinden 12 aylık ortalama yağışları
    doğrudan HTML tablo satırından okur.

    Dönen değer: (aylik_dict, olcum_periyodu, url).
    """
    import re as _re
    import unicodedata as _unicodedata
    from html.parser import HTMLParser

    il_adi = str(il_adi or "").strip()
    if not il_adi:
        return {}, "", None

    _il_url = "".join(
        ch for ch in _unicodedata.normalize("NFKD", il_adi)
        if not _unicodedata.combining(ch)
    ).upper()
    _il_url = (
        _il_url.replace("Ç", "C").replace("Ğ", "G").replace("İ", "I")
        .replace("Ö", "O").replace("Ş", "S").replace("Ü", "U")
    )
    _il_url = MGM_IL_URL_KODU.get(il_adi, _il_url)

    _base = "https://www.mgm.gov.tr/veridegerlendirme/il-ve-ilceler-istatistik.aspx"
    # MGM'de kullanılan farklı çalışan URL biçimleri. k=H özellikle
    # Resmi İklim İstatistikleri tablosunu doğrudan döndürmektedir.
    _urls = [
        f"{_base}?k=&m={_il_url}",
        f"{_base}?k=H&m={_il_url}",
        f"{_base}?m={_il_url}",
        f"{_base}?k=undefined&m={_il_url}",
    ]
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
    _aylar = [
        "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
        "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
    ]
    _baslik = "Aylık Toplam Yağış Miktarı Ortalaması (mm)"

    class _MGMTableParser(HTMLParser):
        def __init__(self):
            super().__init__(convert_charrefs=True)
            self.rows = []
            self._row = None
            self._cell = None
            self._buf = []

        def handle_starttag(self, tag, attrs):
            tag = tag.lower()
            if tag == "tr":
                self._row = []
            elif tag in ("td", "th") and self._row is not None:
                self._cell = tag
                self._buf = []

        def handle_data(self, data):
            if self._cell is not None:
                self._buf.append(data)

        def handle_endtag(self, tag):
            tag = tag.lower()
            if tag in ("td", "th") and self._cell is not None:
                self._row.append(" ".join("".join(self._buf).split()))
                self._cell = None
                self._buf = []
            elif tag == "tr" and self._row is not None:
                if self._row:
                    self.rows.append(self._row)
                self._row = None

    def _sayiya_cevir(_text):
        _t = str(_text or "").strip().replace(" ", "")
        # MGM Türkçe ondalık gösterimini destekle.
        _t = _t.replace("mm", "").strip()
        if not _re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", _t):
            return None
        try:
            return float(_t.replace(",", "."))
        except ValueError:
            return None

    for _url in _urls:
        try:
            _req = urllib.request.Request(_url, headers=_headers)
            with urllib.request.urlopen(_req, timeout=5) as _response:
                _html = _response.read().decode("utf-8", errors="ignore")

            # 1) Önce gerçek HTML tablosunu parse et.
            _parser = _MGMTableParser()
            _parser.feed(_html)
            for _row in _parser.rows:
                _etiket = " ".join(str(_x) for _x in _row[:2])
                if _baslik.lower() in _etiket.lower():
                    _sayilar = []
                    for _hucre in _row[1:]:
                        _n = _sayiya_cevir(_hucre)
                        if _n is not None:
                            _sayilar.append(_n)
                    if len(_sayilar) >= 12:
                        _aylik = dict(zip(_aylar, _sayilar[:12]))
                        _metin = " ".join(_row)
                        _donem = ""
                        _m = _re.search(r"Ölçüm Periyodu\s*\(\s*([^\)]+)", _metin, flags=_re.I)
                        if _m:
                            _donem = _m.group(1).strip()
                        # Dönem çoğunlukla ayrı bir satırda olduğundan HTML
                        # genelinde de ara.
                        if not _donem:
                            _m = _re.search(r"Ölçüm Periyodu\s*\(\s*([^\)]+)", _html, flags=_re.I)
                            if _m:
                                _donem = _m.group(1).strip()
                        return _aylik, _donem, _url

            # 2) HTML yapısı değişirse, yalnızca ilgili satırı düz metinden
            # yakalayan güvenli bir geri dönüş yöntemi kullan.
            _metin = _re.sub(r"<script\b[^>]*>.*?</script>", " ", _html, flags=_re.I | _re.S)
            _metin = _re.sub(r"<style\b[^>]*>.*?</style>", " ", _metin, flags=_re.I | _re.S)
            _metin = _re.sub(r"<[^>]+>", " | ", _metin)
            _metin = _re.sub(r"\s+", " ", _metin).strip()
            _pos = _metin.lower().find(_baslik.lower())
            if _pos >= 0:
                _parca = _metin[_pos:_pos + 1200]
                _vals = _re.findall(r"(?<![\d.,])\d+(?:[.,]\d+)?", _parca)
                _nums = [float(v.replace(",", ".")) for v in _vals]
                if len(_nums) >= 12:
                    return dict(zip(_aylar, _nums[:12])), "", _url
        except Exception:
            continue

    return {}, "", _urls[0]


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
    with ThreadPoolExecutor(max_workers=20) as executor:
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
        <a href="#sayfa_orta">↕ Ortaya Git</a>
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
    # 7.1 yangın bölümü toplu seçim butonları Streamlit widget state'idir;
    # eski proje kayıtlarından yüklenmeleri ValueAssignmentNotAllowedError üretir.
    # 7.1 toplu seçim butonlarının eski ve yeni anahtarlarını proje verisi
    # olarak ASLA geri yükleme. Eski projelerde "_tum_sec" / "_tum_kaldir"
    # biçimindeki anahtarlar bulunduğunda StreamlitValueAssignmentNotAllowedError
    # oluşabilir.
    if "_tum_sec" in anahtar or "_tum_kaldir" in anahtar:
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
        # Rapor aynı Streamlit çalıştırmasında aşağıdaki rapor bloğunda
        # oluşturulur. Burada rerun yapılmaz; böylece rapor isteği kaybolmaz.
        st.session_state["_rapor_olustur_istegi_v134"] = True

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
  # Streamlit Cloud / yerel çalışmada çalışma dizini değişse bile
  # iklim_verileri.json dosyasını app.py'nin bulunduğu klasörden bul.
  dosya_adi = Path(__file__).resolve().parent / "iklim_verileri.json"
  if dosya_adi.exists():
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
    ("rapor_bolum_635", "SU YUMUŞATMA CİHAZI SEÇİMİ"),
]

# Sol menüdeki "Tümünü Seç / Tümünü Kaldır" işlemlerinin kullandığı anahtarlar.
BOLUM_SECIM_ANAHTARLARI = [
    "rapor_bolum_1", "rapor_bolum_2", "rapor_bolum_3", "rapor_bolum_4",
    "rapor_bolum_5", "rapor_bolum_51", "rapor_bolum_6", "rapor_bolum_61",
    "rapor_bolum_611", "rapor_bolum_62", "rapor_bolum_621", "rapor_bolum_622", "rapor_bolum_623",
    "rapor_bolum_63", "rapor_bolum_631", "rapor_bolum_631_1", "rapor_bolum_631_2",
    "rapor_bolum_631_2_1", "rapor_bolum_631_2_2", "rapor_bolum_631_2_3",
    "rapor_bolum_631_2_4", "rapor_bolum_631_2_5", "rapor_bolum_631_2_6",
    "rapor_bolum_631_2_7", "rapor_bolum_632",
    "rapor_bolum_633", "rapor_bolum_634", "rapor_bolum_635",
    "rapor_bolum_7", "rapor_bolum_71", "rapor_bolum_711", "rapor_bolum_712", "rapor_bolum_713",
    "rapor_bolum_72", "rapor_bolum_721", "rapor_bolum_722", "rapor_bolum_723", "rapor_bolum_724", "rapor_bolum_725",
    "rapor_bolum_73", "rapor_bolum_731", "rapor_bolum_732", "rapor_bolum_733",
    "rapor_bolum_74", "rapor_bolum_741", "rapor_bolum_742", "rapor_bolum_743", "rapor_bolum_744",
    "rapor_bolum_75", "rapor_bolum_751", "rapor_bolum_752", "rapor_bolum_753",
    "rapor_bolum_76", "rapor_bolum_761", "rapor_bolum_762", "rapor_bolum_763", "rapor_bolum_764", "rapor_bolum_765",
    "rapor_bolum_77", "rapor_bolum_78", "rapor_bolum_79",
]
def _63_dinamik_no(anahtar):
    """6.3 alt bölüm numarasını sabit tutar; seçim durumuna göre yeniden numaralandırmaz."""
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
# SOL MENÜ / BÖLÜM NAVİGASYONU
# ---------------------------------------------------------------------------
# Bölüm seçimleri artık sol menüde tutulur. Bölüm adları aynı zamanda
# sayfadaki ilgili başlığa bağlantıdır; böylece tıklandığında doğrudan
# seçilen bölüme gidilir. Checkbox'lar rapora dahil/hariç mantığını korur.
_BOLUM_NAV = [
    ("1. Kapak Bilgileri", "bolum_1", "rapor_bolum_1"),
    ("2. Uygulanacak Standart ve Yönetmelikler", "bolum_2", "rapor_bolum_2"),
    ("3. Mekanik Tesisat Proje Kapsamı", "bolum_3", "rapor_bolum_3"),
    ("4. Tesiste Kullanılacak Isı İletim Akışkanları", "bolum_4", "rapor_bolum_4"),
    ("5. İklim, Konfor Şartları ve Tasarım Kriterleri", "bolum_5", "rapor_bolum_5"),
    ("5.1 Dış Hava Tasarım Kriterleri", "bolum_51", "rapor_bolum_51"),
    ("6. Sıhhi Tesisat", "bolum_6", "rapor_bolum_6"),
    ("6.1 Sıhhi Tesisat Ön Bilgiler", "bolum_61", "rapor_bolum_61"),
    ("6.1.1 Temiz Su Hesabı", "bolum_611", "rapor_bolum_611"),
    ("6.2 Pis Su Tesisatı Esasları", "bolum_62", "rapor_bolum_62"),
    ("6.2.1 Pis Su Hesabı", "bolum_621", "rapor_bolum_621"),
    ("6.2.2 Pis Su Terfi Pompaları", "bolum_622", "rapor_bolum_622"),
    ("6.2.3 Yağ Ayırıcı Seçimleri", "bolum_623", "rapor_bolum_623"),
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
    ("6.3.5 SU YUMUŞATMA CİHAZI SEÇİMİ", "bolum_635", "rapor_bolum_635"),
    ("7. YANGIN TESİSATI", "bolum_7", "rapor_bolum_7"),
    ("7.1 YANGIN TESİSATI GENEL ESASLARI", "bolum_71", "rapor_bolum_71"),
    ("7.1.1 Yangın Tesisatının Amacı ve Kapsamı", "bolum_711", "rapor_bolum_711"),
    ("7.1.2 Yangın Tesisatı Tasarım Esasları", "bolum_712", "rapor_bolum_712"),
    ("7.1.3 Yangın Tesisatı Standartları", "bolum_713", "rapor_bolum_713"),
    ("7.2 YANGIN TEHLİKE SINIFI VE TASARIM KRİTERLERİ", "bolum_72", "rapor_bolum_72"),
    ("7.2.1 Bina Kullanım Amacı", "bolum_721", "rapor_bolum_721"),
    ("7.2.2 Yangın Tehlike Sınıfı", "bolum_722", "rapor_bolum_722"),
    ("7.2.3 Yangın Bölmeleri", "bolum_723", "rapor_bolum_723"),
    ("7.2.4 Tasarım Kriterleri", "bolum_724", "rapor_bolum_724"),
    ("7.2.5 Tasarım Debisi", "bolum_725", "rapor_bolum_725"),
    ("7.3 YANGIN DOLABI SİSTEMİ TASARIMI VE HESAPLAMALARI", "bolum_73", "rapor_bolum_73"),
    ("7.4 HİDRANT SİSTEMİ TASARIMI VE HESAPLAMALARI", "bolum_74", "rapor_bolum_74"),
    ("7.5 SPRİNKLER (YAĞMURLAMA) SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_75", "rapor_bolum_75"),
    ("7.6 GAZLI SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_76", "rapor_bolum_76"),
    ("7.7 KÖPÜKLÜ SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_77", "rapor_bolum_77"),
    ("7.8 DAVLUMBAZ SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_78", "rapor_bolum_78"),
    ("7.9 MERDİVEN BASINÇLANDIRMA SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_79", "rapor_bolum_79"),
    ("7.10 ASANSÖR BASINÇLANDIRMA SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_710", "rapor_bolum_710"),
    ("7.11 DUMAN TAHLİYE SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_7110", "rapor_bolum_7110"),
    ("7.12 YANGIN SUYU DEPOLAMA SİSTEMİ TASARIM VE HESAPLAMALARI", "bolum_7120", "rapor_bolum_7120"),
    ("7.13 YANGIN POMPA GRUBU TASARIM VE HESAPLAMALARI", "bolum_7130", "rapor_bolum_7130"),
    ("7.14 YANGIN TESİSATI HİDROLİK HESAPLARI", "bolum_7140", "rapor_bolum_7140"),
    ("7.15 YANGIN TESİSATI SONUÇ TABLOSU", "bolum_7150", "rapor_bolum_7150"),
    ("8. ISITMA TESİSATI", "bolum_8", "rapor_bolum_8"),
    ("9. SOĞUTMA TESİSATI", "bolum_9", "rapor_bolum_9"),
    ("10. HAVALANDIRMA TESİSATI", "bolum_10", "rapor_bolum_10"),
]

with st.sidebar:
    st.markdown("## 📑 PROJE BÖLÜMLERİ")
    st.caption("Bölüm adına tıklayarak doğrudan o bölüme gidebilirsiniz.")

    c1, c2 = st.columns(2)
    with c1:
        st.button(
            "☑ TÜMÜNÜ SEÇ",
            key="rapor_tumunu_sec_v92",
            on_click=_tum_bolumleri_sec,
            use_container_width=True,
        )
    with c2:
        st.button(
            "☐ TÜMÜNÜ KALDIR",
            key="rapor_tumunu_kaldir_v92",
            on_click=_tum_bolumleri_kaldir,
            use_container_width=True,
        )

    st.markdown("### Bölümler")
    # Tüm yan sekmelerin seçim kutuları varsayılan olarak açık gelir.
    for _baslik, _anchor, _key in _BOLUM_NAV:
        if _key not in st.session_state:
            st.session_state[_key] = True

        if _anchor.startswith("bolum_63"):
            # 6.3 alt başlıklarında numara, seçili kardeş bölümlere göre anlık değişir.
            # Böylece örneğin 6.3.1 kapatılırsa 6.3.2 -> 6.3.1,
            # 6.3.3 -> 6.3.2 ve 6.3.4 -> 6.3.3 olur.
            _baslik = _63_sidebar_baslik(_key, _baslik)

        _c_nav, _c_chk = st.columns([8.5, 1.5], vertical_alignment="center")
        with _c_nav:
            st.markdown(
                f'<a class="proje-nav-tab" href="#{_anchor}">▸ {_baslik}</a>',
                unsafe_allow_html=True,
            )
        with _c_chk:
            st.checkbox(
                "",
                key=_key,
                label_visibility="collapsed",
            )

    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] { min-width: 330px; max-width: 380px; }
        [data-testid="stSidebar"] .proje-nav-tab {
            display:block;
            margin:3px 0 4px 0;
            padding:7px 8px;
            border-radius:6px;
            text-decoration:none !important;
            font-size:0.92rem;
            color:inherit;
            background:rgba(128,128,128,0.08);
        }
        [data-testid="stSidebar"] .proje-nav-tab:hover {
            background:rgba(128,128,128,0.18);
        }
        [data-testid="stSidebar"] .proje-nav-ana {
            margin-left:0;
            font-weight:600;
            background:rgba(70,120,200,0.10);
        }
        [data-testid="stSidebar"] [data-testid="stCheckbox"] {
            margin-bottom:0;
            transform:scale(1.05);
            transform-origin:center;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

# Ana ekranda kullanılacak eski değişken adları korunuyor.
bolum_2_aktif_ui = st.session_state.get("rapor_bolum_2", False)
bolum_3_aktif_ui = st.session_state.get("rapor_bolum_3", False)
bolum_4_aktif_ui = st.session_state.get("rapor_bolum_4", False)
bolum_5_aktif_ui = st.session_state.get("rapor_bolum_5", False)
bolum_51_aktif_ui = st.session_state.get("rapor_bolum_51", False)
bolum_6_aktif_ui = st.session_state.get("rapor_bolum_6", False)
bolum_61_aktif_ui = st.session_state.get("rapor_bolum_61", False)
bolum_611_aktif_ui = st.session_state.get("rapor_bolum_611", False)
bolum_62_aktif_ui = st.session_state.get("rapor_bolum_62", False)
bolum_621_aktif_ui = st.session_state.get("rapor_bolum_621", False)
bolum_622_aktif_ui = st.session_state.get("rapor_bolum_622", False)
bolum_623_aktif_ui = st.session_state.get("rapor_bolum_623", False)
bolum_63_aktif_ui = st.session_state.get("rapor_bolum_63", False)
bolum_631_aktif_ui = st.session_state.get("rapor_bolum_631", False)
bolum_632_aktif_ui = st.session_state.get("rapor_bolum_632", False)
bolum_633_aktif_ui = st.session_state.get("rapor_bolum_633", False)
bolum_634_aktif_ui = st.session_state.get("rapor_bolum_634", False)
bolum_635_aktif_ui = st.session_state.get("rapor_bolum_635", False)

# ---------------------------------------------------------------------------
# ALT GRUPLAR İÇİN TOPLU SEÇİM BUTONLARI
# ---------------------------------------------------------------------------
def _toplu_secim_butonlari(anahtarlar, grup_adi=None):
    """Alt checkbox grubunu güvenli biçimde topluca seç/kaldır."""
    if not anahtarlar:
        return

    # Widget anahtarları sabit olmalı; hash() kullanmıyoruz.
    grup_imza = grup_adi or "_".join(anahtarlar)
    sec_key = f"alt_grup_sec_v55_{grup_imza}"
    kaldir_key = f"alt_grup_kaldir_v55_{grup_imza}"

    def _grubu_sec():
        for anahtar in anahtarlar:
            st.session_state[anahtar] = True

    def _grubu_kaldir():
        for anahtar in anahtarlar:
            st.session_state[anahtar] = False

    col_sec, col_kaldir = st.columns(2)
    with col_sec:
        st.button(
            "☑ TÜMÜNÜ SEÇ",
            key=sec_key,
            on_click=_grubu_sec,
            use_container_width=True,
        )
    with col_kaldir:
        st.button(
            "☐ TÜMÜNÜ KALDIR",
            key=kaldir_key,
            on_click=_grubu_kaldir,
            use_container_width=True,
        )

# ---------------------------------------------------------------------------
# HİYERARŞİK AKTİFLİK
# ---------------------------------------------------------------------------
# Burada checkbox Session State'lerini değiştirmiyoruz.
# Örneğin sadece 6.3.2 seçilmişse:
#   rapor_bolum_632 = True
#   bolum_63_aktif = True
#   bolum_6_aktif = True
# olur; fakat 6.3 checkbox'ı kullanıcı arayüzünde zorla işaretlenmez.
# Böylece TÜMÜNÜ KALDIR sonrası herhangi bir alt bölüm seçilebilir.
bolum_51_aktif = bool(st.session_state.get("rapor_bolum_51", False))
bolum_5_aktif = bool(st.session_state.get("rapor_bolum_5", False)) or bolum_51_aktif

bolum_611_aktif = bool(st.session_state.get("rapor_bolum_611", False))
bolum_61_aktif = bool(st.session_state.get("rapor_bolum_61", False)) or bolum_611_aktif

bolum_621_aktif = bool(st.session_state.get("rapor_bolum_621", False))
bolum_622_aktif = bool(st.session_state.get("rapor_bolum_622", False))
bolum_623_aktif = bool(st.session_state.get("rapor_bolum_623", False))
bolum_62_aktif = (
    bool(st.session_state.get("rapor_bolum_62", False))
    or bolum_621_aktif
    or bolum_622_aktif
    or bolum_623_aktif
)

# 6.3.1 ana checkbox'ı bu grubun ana aç/kapat kontrolüdür.
# Alt checkbox'lar kendi alt başlıklarını kontrol eder; ana bölüm kapatıldığında
# alt seçimler açık kalsa bile 6.3.1 rapora dahil edilmez ve numaralandırmada
# kardeş bölüm olarak sayılmaz.
bolum_631_aktif = bool(st.session_state.get("rapor_bolum_631", False))
bolum_631_1_aktif = bolum_631_aktif and bool(st.session_state.get("rapor_bolum_631_1", False))
bolum_631_2_alt_anahtarlar = [
    "rapor_bolum_631_2", "rapor_bolum_631_2_1", "rapor_bolum_631_2_2",
    "rapor_bolum_631_2_3", "rapor_bolum_631_2_4", "rapor_bolum_631_2_5",
    "rapor_bolum_631_2_6", "rapor_bolum_631_2_7",
]
bolum_631_2_aktif = bolum_631_aktif and any(
    bool(st.session_state.get(k, False)) for k in bolum_631_2_alt_anahtarlar
)
bolum_632_aktif = bool(st.session_state.get("rapor_bolum_632", False))
bolum_633_aktif = bool(st.session_state.get("rapor_bolum_633", False))
bolum_634_aktif = bool(st.session_state.get("rapor_bolum_634", False))
bolum_635_aktif = bool(st.session_state.get("rapor_bolum_635", False))
bolum_63_aktif = (
    bool(st.session_state.get("rapor_bolum_63", False))
    or bolum_631_aktif
    or bolum_632_aktif
    or bolum_633_aktif
    or bolum_634_aktif
    or bolum_635_aktif
)

bolum_6_aktif = (
    bool(st.session_state.get("rapor_bolum_6", False))
    or bolum_61_aktif or bolum_611_aktif
    or bolum_62_aktif or bolum_621_aktif or bolum_622_aktif or bolum_623_aktif
    or bolum_63_aktif or bolum_631_aktif or bolum_632_aktif or bolum_633_aktif or bolum_634_aktif or bolum_635_aktif
)

bolum_2_aktif = bool(st.session_state.get("rapor_bolum_2", False))
bolum_3_aktif = bool(st.session_state.get("rapor_bolum_3", False))
bolum_4_aktif = bool(st.session_state.get("rapor_bolum_4", False))

# Bölümler bağımsız seçilebildiği için rapor değişkenlerinin güvenli
# başlangıç değerleri burada tanımlanır. İlgili modül aktifse gerçek değerlerle
# aşağıda üzerine yazılır.
secilen_depo_tipi_metni = ""
sih_depo_tipleri = []
# 6.2.3 Yağ Ayırıcı seçimleri için güvenli başlangıç değerleri.
yag_ayirici_maddeleri = [
    "Mutfak/yemekhane atık suyu yağ ayırıcıdan geçirilecektir.",
    "Yağ ayırıcı kapasitesi, sisteme gelen atık su debisine göre belirlenecektir.",
    "Yağ ayırıcı hesabında mutfak ekipmanlarının eş zamanlı kullanım durumu dikkate alınacaktır.",
    "Yağ ayırıcı, kolay temizlenebilir ve bakım yapılabilir özellikte olacaktır.",
    "Yağ ayırıcı üzerinde yeterli büyüklükte bakım ve temizleme kapağı bulunacaktır.",
    "Yağ ayırıcı, yağ ve katı maddelerin kanalizasyon sistemine taşınmasını önleyecek şekilde seçilecektir.",
    "Yağ ayırıcı çıkışında gerekli koku kontrolü ve havalandırma düzeni sağlanacaktır.",
    "Yağ ayırıcının montajı, bakım ve temizlik sırasında kolay erişilebilecek şekilde yapılacaktır.",
    "Yağ ayırıcının giriş ve çıkış bağlantı çapları tesisat boru çaplarına uygun olacaktır.",
    "Yağ ayırıcı kapasitesi ve bağlantı çapı raporda gösterilecektir.",
]
yag_ayirici_secimler = [bool(st.session_state.get(f"yag_ayirici_sec_{i}", True)) for i in range(1, 11)]
ek_yag_ayirici_notu = str(st.session_state.get("ek_yag_ayirici_notu", ""))


# ---------------------------------------------------------------------------
# 6.2.3 YAĞ AYIRICI HESAP MOTORU — YAĞ AYIRICI HESABI.xlsx ile birebir
# ---------------------------------------------------------------------------
# Excel'deki 4 sayfa YA-01 ... YA-04 aynı hesap şablonunu kullanır.
# Programda ekipman adetleri kullanıcı tarafından değiştirilebilir; aşağıdaki
# varsayılan adetler Excel dosyasındaki ilgili sayfalardaki örnek değerlerdir.
YAG_AYIRICI_EKIPMANLARI = [
    ("Pişirme Kazan Çıkışı Ø 25 mm", 1.0, "normal"),
    ("Pişirme Kazan Çıkışı Ø 50 mm", 2.0, "normal"),
    ("Devirme Tipi Kazan Çıkışı Ø 70 mm", 1.0, "normal"),
    ("Devirme Tipi Kazan Çıkışı Ø 100 mm", 3.0, "normal"),
    ("Sifonlu Evye Çıkışı Ø 40", 0.8, "normal"),
    ("Sifonlu Evye Çıkışı Ø 50", 1.5, "normal"),
    ("Sifonsuz Evye Çıkışı Ø 40", 2.5, "normal"),
    ("Sifonsuz Evye Çıkışı Ø 50", 4.0, "normal"),
    ("Bulaşık Makinası", 2.0, "bulasik"),
    ("Devirme Tipi Kızartma Tavası", 1.0, "normal"),
    ("Normal Kızartma Tavası", 0.1, "normal"),
    ("Yüksek Basınçlı/Buharlı Temizleme Makinası", 2.0, "normal"),
    ("Kabuk Soyma Makinası", 1.5, "normal"),
    ("Sebze Yıkama Makinası", 2.0, "normal"),
    ("Pişirme Fırını", 0.5, "normal"),
    ("Yer süzgeci", 0.5, "normal"),
    ("Musluk DN 15 (R 1/2\")", 0.5, "normal"),
    ("DN 20 (R 3/4\")", 1.0, "normal"),
    ("DN 25 (R 1\")", 1.7, "normal"),
]

# Program başlangıcında bütün yağ ayırıcı ekipman adetleri 0'dır.
# Excel'deki YA-01 ... YA-04 örnek adetleri referans alınmış olsa da kullanıcıya
# başlangıç değeri olarak aktarılmaz.
YAG_AYIRICI_VARSAYILAN_ADETLERI = [0] * len(YAG_AYIRICI_EKIPMANLARI)

# 2026 mekanik tesisat poz listesine göre standart yağ ayırıcılar.
# Kapasiteler: 1, 2, 3, 4, 7 ve 10 L/s.
YAG_AYIRICI_POZLARI = [
    {"poz": "25.620.1201", "kapasite": 1.0, "tanim": "Kapasite: 1 lt/sn, et kalınlığı: min.1,5 mm, yağ hacmi: 47 litre, 880x510x490 mm yağ ayırıcı AISI 304 kalite 18/8 Cr-Ni"},
    {"poz": "25.620.1202", "kapasite": 2.0, "tanim": "Kapasite: 2 lt/sn, et kalınlığı: min.1,5 mm, yağ hacmi: 80 litre, 1190x660x710 mm yağ ayırıcı AISI 304 kalite 18/8 Cr-Ni"},
    {"poz": "25.620.1203", "kapasite": 3.0, "tanim": "Kapasite: 3 lt/sn, et kalınlığı: min.1,5 mm, yağ hacmi: 135 litre, 1250x850x970 mm yağ ayırıcı AISI 304 kalite 18/8 Cr-Ni"},
    {"poz": "25.620.1204", "kapasite": 4.0, "tanim": "Kapasite: 4 lt/sn, et kalınlığı: min.2 mm, yağ hacmi: 160 litre, 1580x910x1030 mm yağ ayırıcı AISI 304 kalite 18/8 Cr-Ni"},
    {"poz": "25.620.1205", "kapasite": 7.0, "tanim": "Kapasite: 7 lt/sn, et kalınlığı: min.3 mm, yağ hacmi: 350 litre, 2000x1000x1300 mm yağ ayırıcı AISI 304 kalite 18/8 Cr-Ni"},
    {"poz": "25.620.1206", "kapasite": 10.0, "tanim": "Kapasite: 10 lt/sn, et kalınlığı: min.3 mm, yağ hacmi: 500 litre, 2500x1430x1300 mm yağ ayırıcı AISI 304 kalite 18/8 Cr-Ni"},
]

def _yag_ayirici_poz_otomatik_sec(ns):
    """NS değerini karşılayan en küçük standart yağ ayırıcı pozunu seçer."""
    try:
        gerekli = float(ns)
    except Exception:
        gerekli = 0.0
    if gerekli <= 0:
        return None
    for _poz in YAG_AYIRICI_POZLARI:
        if _poz["kapasite"] >= gerekli:
            return dict(_poz)
    return None

def _yag_ayirici_zi(adet, ekipman_tipi="normal"):
    """Excel'deki Zi(n) kademelerini aynen uygular."""
    n = int(adet)
    if ekipman_tipi == "bulasik":
        # Excel Bulaşık Makinası satırı: 0/1=0.60, 2=0.50, 3=0.40,
        # 4=0.34, 5 ve üzeri=0.30.
        if n <= 1:
            return 0.60
        if n == 2:
            return 0.50
        if n == 3:
            return 0.40
        if n == 4:
            return 0.34
        return 0.30
    # Diğer bütün Excel satırları: 0/1=0.45, 2=0.31, 3=0.25,
    # 4=0.21, 5 ve üzeri=0.20.
    if n <= 1:
        return 0.45
    if n == 2:
        return 0.31
    if n == 3:
        return 0.25
    if n == 4:
        return 0.21
    return 0.20


def _yag_ayirici_hesapla(ya_adi, adetler, fd=1.0, ft=1.0, fr=1.0, secilen_kapasite=0.0, secilen_poz=None, poz_rapora_aktar=False, poz_secim_modu="Otomatik"):
    satirlar = []
    qs = 0.0
    for i, ((ekipman, qi, tip), adet) in enumerate(zip(YAG_AYIRICI_EKIPMANLARI, adetler), start=1):
        n = int(adet)
        n_x_qi = n * float(qi)
        zi = _yag_ayirici_zi(n, tip)
        pis_su_debisi = n_x_qi * zi
        qs += pis_su_debisi
        # Adedi 0 olan ekipmanlar hesap sonucuna ve rapora dahil edilmez.
        # Böylece kullanılmayan ekipmanlar raporda gereksiz yer kaplamaz.
        if n > 0:
            satirlar.append({
                "sira": i,
                "ekipman": ekipman,
                "adet": n,
                "qi": float(qi),
                "n_x_qi": n_x_qi,
                "zi": zi,
                "pis_su_debisi": pis_su_debisi,
            })
    ns = qs * float(fd) * float(ft) * float(fr)
    return {
        "ya_adi": ya_adi,
        "satirlar": satirlar,
        "qs": qs,
        "fd": float(fd),
        "ft": float(ft),
        "fr": float(fr),
        "ns": ns,
        "secilen_kapasite": float(secilen_kapasite),
        "secilen_poz": dict(secilen_poz) if isinstance(secilen_poz, dict) else None,
        "poz_rapora_aktar": bool(poz_rapora_aktar),
        "poz_secim_modu": str(poz_secim_modu),
    }

# 6.2.3 hesap sonuçları rapor üretiminden önce güvenli biçimde hazır tutulur.
yag_ayirici_hesaplari = {}
yag_ayirici_secilenler = list(st.session_state.get("yag_ayirici_secilenler", ["YA-01"]))
sih_sec_depo_tipi = []
depo_gerekli_hacim_m3 = 0.0
depo_gerekli_hacim_litre = 0.0
otomatik_poz_kayitlari = []
poz_numarasi = ""
poz_tipi = ""
sicak_su_hesap_detaylari = []
sicak_su_gunluk_toplam_litre = 0.0
sicak_su_yapi_tipi = ""
kullanma_es_faktoru = 0.0
depolama_faktoru = 0.0
hidrofor_hesaplari = []

# ---------------------------------------------------------------------------
# ANA TESİSAT KATEGORİLERİ - GÖRSEL DÜZEN
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* Ana kategori sekmelerini daha okunaklı ve belirgin yap */
    div[data-testid="stTabs"] {
        margin-top: 6px;
    }

    div[data-testid="stTabs"] [role="tablist"] {
        gap: 4px !important;
        border-bottom: 2px solid #D7DEE8 !important;
        overflow-x: auto !important;
        padding-bottom: 0 !important;
    }

    div[data-testid="stTabs"] button[role="tab"] {
        font-size: 16px !important;
        font-weight: 700 !important;
        padding: 11px 16px !important;
        min-height: 44px !important;
        border-radius: 8px 8px 0 0 !important;
        color: #334155 !important;
        background: #F3F6FA !important;
        border: 1px solid #D7DEE8 !important;
        border-bottom: 3px solid transparent !important;
        white-space: nowrap !important;
    }

    div[data-testid="stTabs"] button[role="tab"]:hover {
        color: #0B3D91 !important;
        background: #EAF2F8 !important;
    }

    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        color: #0B3D91 !important;
        background: #EAF2F8 !important;
        border-color: #B8C7D9 !important;
        border-bottom: 4px solid #0B3D91 !important;
    }

    div[data-testid="stTabs"] button[role="tab"] p {
        font-size: 16px !important;
        font-weight: 700 !important;
        margin: 0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# ANA TESİSAT KATEGORİLERİ
# ---------------------------------------------------------------------------
_t_genel, _t_sihhi, _t_yangin, _t_isitma, _t_sogutma, _t_havalandirma = st.tabs(
    [
        "1–5 GENEL BİLGİLER",
        "6. SIHHİ TESİSAT",
        "7. YANGIN TESİSATI",
        "8. ISITMA TESİSATI",
        "9. SOĞUTMA TESİSATI",
        "10. HAVALANDIRMA",
    ]
)

with _t_genel:
    # --- 1. SEKME / BÖLÜM: KAPAK BİLGİLERİ ---
    st.markdown('<div id="bolum_1"></div>', unsafe_allow_html=True)
    st.header("1. Kapak Bilgileri")
    sirket_adi = st.text_input(
        "Şirket / Kuruluş İsmi",
        "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ",
        key="sirket_adi",
    )
    is_adi = st.text_input("İşin Adı / Proje Başlığı", "", key="is_adi", on_change=_is_adi_degisti)
    rapor_turu = st.text_input(
        "Rapor Türü", "MEKANİK TESİSAT UYGULAMA PROJESİ HESAP RAPORU"
    )
    hazirlayan = st.text_input("Hazırlayan Mühendis", "Mehmet Küçük")
    mmo_no = st.text_input("MMO Oda No", "109913")
    tarih = st.text_input("Rapor Tarihi", bugun_ay_yil)

    if bolum_2_aktif:
      # --- 2. SEKME / BÖLÜM: UYGULANACAK STANDART VE YÖNETMELİKLER ---
      st.markdown('<div id="bolum_2"></div>', unsafe_allow_html=True)
      st.header("2. UYGULANACAK STANDART VE YÖNETMELİKLER")
      st.caption("Rapor seçimi: " + ("Dahil" if bolum_2_aktif else "Hariç"))
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

    if bolum_3_aktif:
      # --- 3. SEKME / BÖLÜM: MEKANİK TESİSAT PROJE KAPSAMI ---
      st.markdown('<div id="bolum_3"></div>', unsafe_allow_html=True)
      st.header("3. MEKANİK TESİSAT PROJE KAPSAMI")
      st.caption("Rapor seçimi: " + ("Dahil" if bolum_3_aktif else "Hariç"))
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

    # -----------------------------------------------------------------------------
    # BÖLÜM 4 SEÇİMİ KAPALI OLSA BİLE KULLANILAN GÜVENLİ VARSAYILANLAR
    # -----------------------------------------------------------------------------
    # 6.3.3 Boyler hesabı, 4. bölümün rapora dahil edilip edilmediğinden bağımsız
    # çalışabilir. Bu nedenle rej_boyler değişkeni mutlaka önceden tanımlı olmalıdır.
    # Böylece "TÜMÜNÜ KALDIR" sonrasında yalnızca 6.3.3 tekrar açıldığında
    # NameError oluşmaz. 4. bölüm açılırsa aşağıdaki selectbox değeri bunu günceller.
    st.session_state.setdefault("rej_boyler_v53", "80/60")
    st.session_state.setdefault("rej_gunes_ist_v68", "60/40")
    rej_boyler = st.session_state.get("rej_boyler_v53", "80/60")

    if bolum_4_aktif:
      # --- 4. SEKME / BÖLÜM: TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI ---
      st.markdown('<div id="bolum_4"></div>', unsafe_allow_html=True)
      st.header("4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI")
      st.caption("Rapor seçimi: " + ("Dahil" if bolum_4_aktif else "Hariç"))
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
          "chk_gunes_ist",
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

        chk_gunes_ist = st.checkbox(
            "11. Güneş Enerjisi Isıtma Tesisatı",
            key="chk_gunes_ist",
            value=True,
        )
        rej_gunes_ist = st.selectbox(
            "Güneş Enerjisi Isıtma Rejimi:",
            ["60/40"],
            index=0,
            key="rej_gunes_ist_v68",
        )

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
        # Boyler ısıtma rejimi varsayılan olarak 80/60 olsun.
        # Seçim bölüm 4 kapalıyken de korunur; bölüm tekrar açıldığında aynı değer gelir.
        rej_boyler = st.selectbox(
            "Boyler Isıtma Rejimi:",
            sicaklik_secenekleri,
            index=sicaklik_secenekleri.index(st.session_state["rej_boyler_v53"]),
            key="rej_boyler_v53",
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

    if bolum_5_aktif:
      # --- 5. BÖLÜM: İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ ---
      st.markdown('<div id="bolum_5"></div>', unsafe_allow_html=True)
      st.header("5. İKLİM, KONFOR ŞARTLARI VE TASARIM KRİTERLERİ")
      st.caption("Rapor seçimi: " + ("Dahil" if bolum_5_aktif else "Hariç"))
      if bolum_51_aktif:
        st.markdown('<div id="bolum_51"></div>', unsafe_allow_html=True)
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




# 6. SIHHİ TESİSAT — modüler dosyadan çalıştırılır
run_fragment("sihhi.py", globals())
# 7. YANGIN TESİSATI — modüler dosyadan çalıştırılır
run_fragment("yangin.py", globals())

with _t_isitma:
    st.header("8. ISITMA TESİSATI")
    st.info("Isıtma tesisatı modülü bu sekme altında yer alacaktır.")

with _t_sogutma:
    st.header("9. SOĞUTMA TESİSATI")
    st.info("Soğutma tesisatı modülü bu sekme altında yer alacaktır.")

with _t_havalandirma:
    st.header("10. HAVALANDIRMA TESİSATI")
    st.info("Havalandırma tesisatı modülü bu sekme altında yer alacaktır.")



# RAPOR OLUŞTURMA — modüler rapor motoru
run_fragment("rapor.py", globals())
