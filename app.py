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
        f"{_base}?k=H&m={_il_url}",
        f"{_base}?k=undefined&m={_il_url}",
        f"{_base}?m={_il_url}",
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
            with urllib.request.urlopen(_req, timeout=20) as _response:
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
    "rapor_bolum_611", "rapor_bolum_62", "rapor_bolum_621", "rapor_bolum_622", "rapor_bolum_623",
    "rapor_bolum_63", "rapor_bolum_631", "rapor_bolum_631_1", "rapor_bolum_631_2",
    "rapor_bolum_631_2_1", "rapor_bolum_631_2_2", "rapor_bolum_631_2_3",
    "rapor_bolum_631_2_4", "rapor_bolum_631_2_5", "rapor_bolum_631_2_6",
    "rapor_bolum_631_2_7", "rapor_bolum_632",
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
    ("6.2.3 YAĞ AYIRICI SEÇİMLERİ", "bolum_623", "rapor_bolum_623"),
    ("6.3 Sıhhi Tesisat Cihaz Seçimleri", "bolum_63", "rapor_bolum_63"),
    ("6.3.1 SU DEPOSU KAPASİTE HESAPLAMALARI", "bolum_631", "rapor_bolum_631"),
    ("6.3.1.1 Kullanma Suyu Deposu Seçimi", "bolum_631_1", "rapor_bolum_631_1"),
    ("6.3.1.2 Yağmur Suyu Deposu Seçimi", "bolum_631_2", "rapor_bolum_631_2"),
    ("6.3.1.2.1 Yağmur Suyu Toplama Hesabı", "bolum_631_2_1", "rapor_bolum_631_2_1"),
    ("6.3.1.2.2 Yağmur Suyu Filtresi Seçimi", "bolum_631_2_2", "rapor_bolum_631_2_2"),
    ("6.3.1.2.3 Yağmur Suyu Deposu Hacim Hesabı", "bolum_631_2_3", "rapor_bolum_631_2_3"),
    ("6.3.1.2.4 Taşma Hattı Hesabı", "bolum_631_2_4", "rapor_bolum_631_2_4"),
    ("6.3.1.2.5 Taşma Sifonu / Koku Kapanı", "bolum_631_2_5", "rapor_bolum_631_2_5"),
    ("6.3.1.2.6 Depo Girişi / Sakin Giriş", "bolum_631_2_6", "rapor_bolum_631_2_6"),
    ("6.3.1.2.7 Havalandırma ve Haşere Koruması", "bolum_631_2_7", "rapor_bolum_631_2_7"),
    ("6.3.2 KULLANMA SOĞUK SUYU HİDROFORU SEÇİMİ", "bolum_632", "rapor_bolum_632"),
    ("6.3.3 KULLANMA SICAK SUYU İHTİYACI HESAPLARI", "bolum_633", "rapor_bolum_633"),
    ("6.3.4 KULLANMA SICAK SU TESİSATI RE-SİRKULASYON POMPASI SEÇİMİ", "bolum_634", "rapor_bolum_634"),
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
bolum_63_aktif = (
    bool(st.session_state.get("rapor_bolum_63", False))
    or bolum_631_aktif
    or bolum_632_aktif
    or bolum_633_aktif
    or bolum_634_aktif
)

bolum_6_aktif = (
    bool(st.session_state.get("rapor_bolum_6", False))
    or bolum_61_aktif or bolum_611_aktif
    or bolum_62_aktif or bolum_621_aktif or bolum_622_aktif or bolum_623_aktif
    or bolum_63_aktif or bolum_631_aktif or bolum_632_aktif or bolum_633_aktif or bolum_634_aktif
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


with _t_sihhi:
    if bolum_6_aktif:
      # 6.3 alt bölümleri, 6.1 kapalı olsa bile bağımsız çalışabilmelidir.
      # Bu nedenle 6.1 içindeki widgetlar henüz oluşturulmamış olsa dahi
      # 6.3 tarafında kullanılan değişkenlere güvenli varsayılanlar verilir.
      # 6.1 açıksa aşağıdaki değerler gerçek kullanıcı seçimleriyle değiştirilir.
      sih_sec_depo_tipi = True
      sih_depo_tipleri = ["Paslanmaz Çelik Modüler su deposu"]
      sih_depo_konumlari = ["Bodrum kat"]
      sih_sec_yagmur_depo_tipi = True
      sih_yagmur_depo_tipleri = ["Betonarme Su Deposu"]

      # --- 6. BÖLÜM: SIHHİ TESİSAT ---
      st.markdown('<div id="bolum_6"></div>', unsafe_allow_html=True)
      st.header("6. SIHHİ TESİSAT")
      st.caption("Rapor seçimi: " + ("Dahil" if bolum_6_aktif else "Hariç"))
      if bolum_61_aktif:
        st.markdown('<div id="bolum_61"></div>', unsafe_allow_html=True)
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
            "sih_sec_yagmur_depo_tipi",
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

        sih_sec_yagmur_depo_tipi = st.checkbox(
            "Binada yağmur suyu depolaması için belirtilen tipte su deposu kullanılmıştır.",
            key="sih_sec_yagmur_depo_tipi",
            value=True,
        )
        _yagmur_depo_tipleri_61 = [
            "Paslanmaz Modüler Çelik Su Deposu",
            "Galvaniz Modüler Çelik Su Deposu",
            "Betonarme Su Deposu",
        ]
        sih_yagmur_depo_tipi = st.selectbox(
            "Yağmur Suyu Deposu Tipi:",
            _yagmur_depo_tipleri_61,
            index=2,
            key="sih_yagmur_depo_tipi",
        )
        # Rapor tarafında mevcut yapı korunur; seçim tek bir depo tipidir.
        sih_yagmur_depo_tipleri = [sih_yagmur_depo_tipi]

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

        # Üstteki sıcak su üretim yöntemi aşağıdaki BOYLER / EŞANJÖR SEÇİMİ
        # bölümünün ana seçimidir. Tek/çift serpantin seçimi alt sekmeye otomatik
        # aktarılır; böylece aynı seçim iki kez yapılmaz.
        _ust_boyler_tip_map = {
            "Dik tip hijyenik tek serpantinli boyler": "TEK SERPANTİNLİ BOYLER",
            "Dik tip hijyenik çift serpantinli boyler": "ÇİFT SERPANTİNLİ BOYLER",
            "Plakalı eşanjör akümülasyon tankı": "PLAKALI EŞANJÖR",
        }
        _ust_secili_boylerler = [
            _ust_boyler_tip_map[x]
            for x in sih_sicak_su_yontemleri
            if x in _ust_boyler_tip_map
        ]
        if len(_ust_secili_boylerler) == 1:
            st.session_state["boyler_secili_tip_v57"] = _ust_secili_boylerler[0]
        elif "Dik tip hijyenik çift serpantinli boyler" in sih_sicak_su_yontemleri:
            # İki boyler tipi aynı anda seçilmişse daha özel olan çift serpantinli
            # boyler alt hesapta esas alınır.
            st.session_state["boyler_secili_tip_v57"] = "ÇİFT SERPANTİNLİ BOYLER"
        elif len(_ust_secili_boylerler) > 0:
            st.session_state["boyler_secili_tip_v57"] = _ust_secili_boylerler[-1]

        # Üst bölümde seçilen boyler tipini aşağıdaki sekme görünümüne de zorla aktar.
        # st.tabs() aktif sekmeyi programatik olarak değiştiremez; bu nedenle aşağıda
        # sekme görünümü yatay radio ile oluşturuluyor. Böylece üst seçim yapıldığı
        # anda alt bölümde aynı sekme aktif hale gelir.
        _boyler_alt_sekme_index_map = {
            "TEK SERPANTİNLİ BOYLER": 0,
            "ÇİFT SERPANTİNLİ BOYLER": 1,
            "PLAKALI EŞANJÖR": 2,
        }
        if st.session_state.get("boyler_secili_tip_v57") in _boyler_alt_sekme_index_map:
            st.session_state["boyler_alt_sekme_v71"] = _boyler_alt_sekme_index_map[
                st.session_state["boyler_secili_tip_v57"]
            ]
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
      st.markdown('<div id="bolum_611"></div>', unsafe_allow_html=True)
      st.subheader("6.1.1 Temiz Su Sarfiyat Yükleme Birimleri ve Çap Tayini Girdileri")


      if bolum_62_aktif:
        # --- 6.2 PİS SU TESİSATI ---
        st.markdown('<div id="bolum_62"></div>', unsafe_allow_html=True)
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

        st.markdown('<div id="bolum_621"></div>', unsafe_allow_html=True)
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
        st.markdown('<div id="bolum_622"></div>', unsafe_allow_html=True)
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


        def hidrofor_standart_egrisi(q_m3h, h_mss):
          """Hidrofor için üretici verisi bulunmadığında kullanılan standart Q-H eğrisi.
          Çalışma noktası eğri üzerinde tutulur; bu eğri üretici katalog eğrisi değildir.
          """
          q = max(float(q_m3h), 0.1)
          h = max(float(h_mss), 0.1)
          oranlar = [0.0, 0.50, 0.75, 1.00, 1.25, 1.50, 1.75, 2.00]
          head_carpan = [1.40, 1.22, 1.10, 1.00, 0.82, 0.63, 0.43, 0.22]
          q_curve = [q * r for r in oranlar]
          h_curve = [h * k for k in head_carpan]
          return q_curve, h_curve


        def hidrofor_pompa_secim_egrisi(q_m3h, h_mss, marka_secimi="Standart Pompa"):
          # Hidrofor tarafı pis su pompası poz/eğri tablosundan bağımsızdır.
          # Wilo/Grundfos/Lowara için doğrulanmış hidrofor Q-H veri seti eklenene kadar
          # çalışma noktasını garanti eden standart hidrofor eğrisi kullanılır.
          q_egrisi, h_egrisi = hidrofor_standart_egrisi(q_m3h, h_mss)
          if marka_secimi == "Standart Pompa":
            baslik = "Standart Hidrofor Pompası - Q-H Karakteristik Eğrisi"
          elif marka_secimi == "Wilo":
            baslik = "Wilo - Hidrofor Q-H Eğrisi (standart veri seti)"
          elif marka_secimi == "Grundfos":
            baslik = "Grundfos - Hidrofor Q-H Eğrisi (standart veri seti)"
          elif marka_secimi == "Lowara":
            baslik = "Lowara - Hidrofor Q-H Eğrisi (standart veri seti)"
          else:
            baslik = "Otomatik - Standart Hidrofor Q-H Eğrisi"
          model = {
              "marka": marka_secimi if marka_secimi != "Otomatik (Wilo + Grundfos + Lowara)" else "Standart",
              "seri": "Hidrofor",
              "model": "Standart Q-H",
              "curve": list(zip(q_egrisi, h_egrisi)),
              "h_calisma": float(h_mss),
              "p2_kw": 0.0,
              "kaynak": "Standart hidrofor Q-H eğrisi; üretici katalog eğrisi değildir.",
          }
          return q_egrisi, h_egrisi, baslik, model


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
            "Pis Su Pompası üreticisi / seçim modu",
            ["Otomatik (Wilo + Grundfos)", "Wilo", "Grundfos"],
            index=0,
            key="pompa_marka_secimi",
        )

        hidrofor_marka_secimi = st.selectbox(
            "Hidrofor pompası üreticisi / seçim modu",
            ["Otomatik (Wilo + Grundfos + Lowara)", "Wilo", "Grundfos", "Lowara", "Standart Pompa"],
            index=4,
            key="hidrofor_marka_secimi",
            help="Üreticiye ait doğrulanmış Q-H eğrisi veri seti bulunmadığında sistem standart hidrofor Q-H eğrisi ile devam eder.",
        )

        if secilen_psp_listesi:
          st.write(
              "Her bir terfi pompası çukuru için bina tipi, armatürler, emniyet"
              " faktörleri ve asıl/yedek adetlerini ayrı ayrı girin:"
          )

          # Pis su pompalarını hidroforlarda olduğu gibi yan yana sekmelerde göster.
          psp_tab_basliklari = [
              f"6.2.2.{i} {psp}" for i, psp in enumerate(secilen_psp_listesi, start=1)
          ]
          psp_tabs = st.tabs(psp_tab_basliklari)

          for psp, psp_tab in zip(secilen_psp_listesi, psp_tabs):
            with psp_tab:
              st.markdown(f"### ⚙️ {psp} ÖZEL DEBİ VE GÜÇ HESAP MODÜLÜ")
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
                  "yedek_adedi": yedek_adedi,
                  "toplam_adet": toplam_adet,
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

        # ---------------------------------------------------------------------------
        # 6.2.3 YAĞ AYIRICI SEÇİMLERİ
        # ---------------------------------------------------------------------------
        if bolum_623_aktif:
            st.markdown('<div id="bolum_623"></div>', unsafe_allow_html=True)
            st.subheader("6.2.3 YAĞ AYIRICI SEÇİMLERİ")
            st.caption(
                "Bu bölüm şimdilik yalnızca mutfak / yemekhane kaynaklı yağ ayırıcıları kapsamaktadır. "
                "Petrol / hidrokarbon ayırıcıları ileride ayrı bir başlık altında kurgulanacaktır."
            )

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

            yag_ayirici_keys = [f"yag_ayirici_sec_{i}" for i in range(1, len(yag_ayirici_maddeleri) + 1)]
            _toplu_secim_butonlari(yag_ayirici_keys, grup_adi="yag_ayirici_623")

            yag_ayirici_secimler = []
            yag_ayirici_tab1, yag_ayirici_tab2 = st.tabs([
                "🧈 YAĞ AYIRICI GENEL ESASLARI",
                "📝 İLAVE YAĞ AYIRICI MADDELERİ",
            ])

            with yag_ayirici_tab1:
                st.markdown("#### Seçilecek maddeler")
                for i, madde in enumerate(yag_ayirici_maddeleri, start=1):
                    secili = st.checkbox(
                        madde,
                        key=f"yag_ayirici_sec_{i}",
                        value=True,
                    )
                    yag_ayirici_secimler.append(secili)

            with yag_ayirici_tab2:
                ek_yag_ayirici_notu = st.text_area(
                    "İlave Yağ Ayırıcı Seçim Maddesi (Her satıra bir tane)",
                    "",
                    height=120,
                    key="ek_yag_ayirici_notu",
                )

      if bolum_63_aktif:
        # --- 6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ ---
        if bolum_631_aktif:
            st.markdown('<div id="bolum_63"></div>', unsafe_allow_html=True)
            st.header("6.3 SIHHİ TESİSAT CİHAZ SEÇİMLERİ")
            st.markdown('<div id="bolum_631"></div>', unsafe_allow_html=True)
            st.subheader(_63_dinamik_baslik("rapor_bolum_631"))
            st.markdown('<div id="bolum_631_1"></div>', unsafe_allow_html=True)
            st.markdown("#### 6.3.1.1 KULLANMA SUYU DEPOSU SEÇİMİ:")

            poz_gosterilsin_mi = st.checkbox(
                "Poz numarasını göster",
                value=True,
                key="poz_gosterilsin_mi",
            )

            genel_bilgiler_tab = st.container()
            with genel_bilgiler_tab:
                st.markdown("#### Genel Bilgiler")
                st.markdown(f"#### {_63_dinamik_no('rapor_bolum_631')}.1 KULLANMA SUYU İHTİYACININ BELİRLENMESİ")

                # Hesap türü sekmelerinin daha büyük ve okunabilir görünmesi
                st.markdown(
                    """
                    <style>
                    div[data-testid="stRadio"] > label {
                        font-size: 1.15rem !important;
                        font-weight: 600 !important;
                    }
                    div[data-testid="stRadio"] div[role="radiogroup"] label {
                        font-size: 1.10rem !important;
                        font-weight: 500 !important;
                        min-height: 2rem !important;
                    }
                    div[data-testid="stRadio"] div[role="radiogroup"] label p {
                        font-size: 1.10rem !important;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True,
                )

                hesap_modu = st.radio(
                    "Su ihtiyacı hesabı türü",
                    ["Genel Su Tüketimi", "Konutlar", "Hastaneler"],
                    horizontal=True,
                    key="su_hesap_modu",
                )

                su_hesap_detaylari = []
                secilen_su_kategorileri = []
                su_gunluk_ihtiyac_litre = 0.0
                # Rapor oluşturma bölümünde de kullanılacağı için değişken tüm hesap
                # modlarında tanımlı olmalıdır. Hastane modunda aşağıda özel olarak doldurulur.
                su_tuketim_secenekleri = {}

                if hesap_modu == "Hastaneler":
                    st.markdown("##### Hastane Su İhtiyacı Hesabı")
                    yatak_sayisi = st.number_input(
                        "Yatak sayısı", min_value=0, value=100, step=1, key="hastane_yatak_sayisi"
                    )
                    hastane_satirlari = [
                        ("Hasta", 2.0, 135.0),
                        ("Personel", 3.0, 45.0),
                        ("Geçici hasta", 4.0, 15.0),
                    ]
                    # Hastane seçiliyken rapordaki tüketim değerleri tablosunun
                    # oluşturulabilmesi için hastane tüketim seçeneklerini tanımla.
                    su_tuketim_secenekleri = {
                        cihaz: ("Kişi", tuketim)
                        for cihaz, _katsayi, tuketim in hastane_satirlari
                    }
                    hastane_tablo = []
                    for cihaz, katsayi, tuketim in hastane_satirlari:
                        toplam_sayi = yatak_sayisi * katsayi
                        toplam_litre = toplam_sayi * tuketim
                        hastane_tablo.append({
                            "CİHAZ": cihaz,
                            "Yatak Sayısı": yatak_sayisi,
                            "katsayı": katsayi,
                            "TOPLAM sayısı": toplam_sayi,
                            "Tüketim [lt/gün]": tuketim,
                            "TOPLAM [lt/gün]": toplam_litre,
                        })
                        su_gunluk_ihtiyac_litre += toplam_litre
                        su_hesap_detaylari.append({
                            "kategori": cihaz, "birim": "Kişi",
                            "birim_degeri": tuketim, "miktar": toplam_sayi,
                            "ihtiyac_litre": toplam_litre,
                            "yatak_sayisi": yatak_sayisi,
                            "katsayi": katsayi,
                        })
                    hastane_tablo.append({
                        "CİHAZ": "GENEL TOPLAM",
                        "Yatak Sayısı": "", "katsayı": "",
                        "TOPLAM sayısı": sum(r["TOPLAM sayısı"] for r in hastane_tablo),
                        "Tüketim [lt/gün]": "",
                        "TOPLAM [lt/gün]": su_gunluk_ihtiyac_litre,
                    })
                    st.table(hastane_tablo)
                    su_gunluk_ihtiyac_m3 = su_gunluk_ihtiyac_litre / 1000.0
                    st.metric("Hastane günlük toplam su ihtiyacı", f"{su_gunluk_ihtiyac_litre:,.0f} L/gün".replace(",", "."))
                    st.caption(f"{su_gunluk_ihtiyac_m3:g} m³/gün")

                else:
                    if hesap_modu == "Konutlar":
                        su_tuketim_secenekleri = {
                            "Konutlar - Lavabolu": ("Kişi", 70.0),
                            "Konutlar - Duşlu": ("Kişi", 90.0),
                            "Konutlar - Küvetli": ("Kişi", 160.0),
                        }
                    else:
                        su_tuketim_secenekleri = {
                            "Fabrikalar": ("Kişi", 45.0),
                            "Bürolar": ("Kişi", 45.0),
                            "Okullar - Gündüzlü": ("Kişi", 45.0),
                            "Okullar - Yatılı": ("Kişi", 135.0),
                            "Bahçe sulama": ("m²", 1.5),
                            "Konutlar - Lavabolu": ("Kişi", 70.0),
                            "Konutlar - Duşlu": ("Kişi", 90.0),
                            "Konutlar - Küvetli": ("Kişi", 160.0),
                            "Oteller - Duşlu": ("Kişi", 100.0),
                            "Oteller - Küvetli": ("Kişi", 175.0),
                            "Çocuk Yuvaları": ("Kişi", 125.0),
                            "Kreşler": ("Kişi", 125.0),
                            "Kışlalar": ("Kişi", 60.0),
                            "Lokantalar": ("Kişi", 75.0),
                            "Oto Yıkama - Temizlik": ("Gün", 100.0),
                            "Askeri binalar - Yatılı": ("Kişi", 135.0),
                            "Askeri binalar - Yatılı olmayan": ("Kişi", 45.0),
                        }
                    st.markdown("##### Su Tüketim Değerleri Tablos")
                    su_tuketim_tablosu = [
                        {"Kullanım amacı": kategori, "Birim": birim,
                         "Birim tüketim değeri": f"{deger:g} L/{birim}/gün"}
                        for kategori, (birim, deger) in su_tuketim_secenekleri.items()
                    ]
                    st.table(su_tuketim_tablosu)

                    if hesap_modu == "Konutlar":
                        hane_kisi_sayisi = st.number_input(
                            "Hane başına kişi sayısı", min_value=1, value=4, step=1,
                            key="konut_hane_kisi_sayisi",
                        )
                        toplam_hane_sayisi = st.number_input(
                            "Toplam hane sayısı", min_value=1, value=1, step=1,
                            key="konut_toplam_hane_sayisi",
                        )
                        toplam_kisi_sayisi = hane_kisi_sayisi * toplam_hane_sayisi
                        st.metric("Toplam kişi sayısı", f"{toplam_kisi_sayisi:,.0f}".replace(",", "."))

                        secilen_konut_kategorisi = st.selectbox(
                            "Konut tipi / su tüketim kategorisi",
                            options=list(su_tuketim_secenekleri.keys()),
                            key="secilen_konut_kategorisi",
                        )
                        secilen_su_kategorileri = [secilen_konut_kategorisi]
                        su_birim, su_birim_degeri = su_tuketim_secenekleri[secilen_konut_kategorisi]
                        kategori_ihtiyaci_litre = toplam_kisi_sayisi * su_birim_degeri
                        su_gunluk_ihtiyac_litre += kategori_ihtiyaci_litre
                        su_hesap_detaylari.append({
                            "kategori": secilen_konut_kategorisi, "birim": su_birim,
                            "birim_degeri": su_birim_degeri, "miktar": toplam_kisi_sayisi,
                            "ihtiyac_litre": kategori_ihtiyaci_litre,
                        })
                        st.caption(
                            f"Hesap: {toplam_hane_sayisi:g} hane × {hane_kisi_sayisi:g} kişi = "
                            f"{toplam_kisi_sayisi:g} kişi × {su_birim_degeri:g} L/kişi-gün"
                        )
                        st.markdown("##### Konut Hesap Tablosu")
                        st.table([{
                            "Hane başına kişi sayısı": hane_kisi_sayisi,
                            "Toplam hane sayısı": toplam_hane_sayisi,
                            "Toplam kişi sayısı": toplam_kisi_sayisi,
                            "Birim tüketim": f"{su_birim_degeri:g} L/kişi-gün",
                            "Günlük ihtiyaç": f"{kategori_ihtiyaci_litre:,.0f} L/gün".replace(",", "."),
                        }])
                    else:
                        secilen_su_kategorileri = st.multiselect(
                            "Kullanım amacı / su tüketim kategorileri",
                            options=list(su_tuketim_secenekleri.keys()),
                            key="secilen_su_kategorileri",
                        )
                        if secilen_su_kategorileri:
                            for sira, kategori in enumerate(secilen_su_kategorileri):
                                su_birim, su_birim_degeri = su_tuketim_secenekleri[kategori]
                                su_miktari = st.number_input(
                                    f"{kategori} miktarı ({su_birim})", min_value=0.0, value=1.0, step=1.0,
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
                    st.metric("Günlük toplam su ihtiyacı", f"{su_gunluk_ihtiyac_litre:,.2f} L/gün".replace(",", "X").replace(".", ",").replace("X", "."))
                    if not secilen_su_kategorileri:
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

                            # Poz numarası elle değiştirildiğinde parantez içindeki kapasite de
                            # aynı pozun tanımlı kapasitesinden otomatik olarak alınır.
                            poz_kapasite_eslesmesi = next(
                                (kapasite for kapasite, poz in kayitlar if poz == kullanilacak_poz),
                                None,
                            )
                            gosterilecek_kapasite = (
                                float(poz_kapasite_eslesmesi)
                                if poz_kapasite_eslesmesi is not None
                                else float(manuel_kapasite)
                            )
                            if poz_kapasite_eslesmesi is None and kullanilacak_poz:
                                st.warning(
                                    f"{depo_tipi}: '{kullanilacak_poz}' poz numarası kapasite listesinde bulunamadı. "
                                    "Mevcut kapasite değeri kullanılacaktır."
                                )
                        else:
                            # Elle düzenleme kapalıyken poz, elle girilen kapasiteye en yakın
                            # standart kapasiteye göre otomatik olarak yeniden belirlenir.
                            kullanilacak_poz = kapasiteye_uygun_poz
                            gosterilecek_kapasite = float(manuel_kapasite)
                            st.caption(
                                f"Kapasiteye en yakın otomatik seçilen poz: {kullanilacak_poz} "
                                f"({gosterilecek_kapasite:g} m³)"
                            )

                        otomatik_poz_kayitlari.append((depo_tipi, gosterilecek_kapasite, kullanilacak_poz))
                    else:
                        st.warning(f"{depo_tipi} için kapasite listesi bulunamadı.")

                if poz_gosterilsin_mi and otomatik_poz_kayitlari:
                    for depo_tipi, secilen_kapasite, secilen_poz in otomatik_poz_kayitlari:
                        st.markdown(
                            f'<div style="color:#000000;"><strong>Cihaz Poz No:</strong> {secilen_poz} '
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

        st.markdown('<div id="bolum_631_2"></div>', unsafe_allow_html=True)
        st.markdown("#### 6.3.1.2 YAĞMUR SUYU DEPOSU SEÇİMİ:")

        # ------------------------------------------------------------------
        # 6.3.1.2 YAĞMUR SUYU DEPOSU SEÇİMİ
        # Pompa seçimi burada yapılmaz. Yağmur suyu hidrofor/pompa ihtiyacı
        # aşağıdaki 6.3.2 hidrofor modülüne aktarılabilir.
        # ------------------------------------------------------------------
        yagmur_aktif = st.checkbox(
            "Yağmur suyu sistemi hesabını aktif et",
            value=bool(st.session_state.get("yagmur_suyu_aktif", True)),
            key="yagmur_suyu_aktif",
        )

        yagmur_hesap = {}
        yagmur_secimler = {}
        yagmur_maddeleri = [
            "Yağmur suyu toplama hesabı yapılacaktır.",
            "Yağmur suyu filtresi, hesaplanan yağış debisine uygun kapasitede seçilecektir.",
            "Yağmur suyu deposu hacmi, toplanabilir yağmur suyu ve kullanım ihtiyacı dikkate alınarak belirlenecektir.",
            "Depo taşma hattı, sisteme gelebilecek maksimum yağış debisini güvenli şekilde uzaklaştıracak kapasitede olacaktır.",
            "Taşma hattının kanalizasyona bağlanması halinde koku kapanı ve geri tepme koruması sağlanacaktır.",
            "Depo girişinde sakin giriş düzeni ile depo içerisindeki tortunun yeniden süspanse olması önlenecektir.",
            "Depo havalandırması yapılacak ve havalandırma açıklıkları haşere girişine karşı korunacaktır.",
        ]

        if yagmur_aktif:
            st.markdown('<div id="bolum_631_2_1"></div>', unsafe_allow_html=True)
            st.markdown("##### • YAĞMUR SUYU TOPLAMA HESABI")
            c1, c2, c3 = st.columns(3)
            with c1:
                yagmur_cati_alani = st.number_input(
                    "Yağmur suyu toplama alanı A (m²)", min_value=0.0,
                    value=float(st.session_state.get("yagmur_cati_alani", 1000.0)),
                    step=10.0, key="yagmur_cati_alani"
                )
            with c2:
                # Seçilen il için MGM'den üç alternatif yağış verisi alınır.
                _mgm_yagis_mm, _mgm_yagis_tarih, _mgm_yagis_url = mgm_gunluk_en_yuksek_yagis_mm(secilen_il)
                _mgm_aylik, _mgm_aylik_periyot, _mgm_aylik_url = mgm_aylik_ortalama_yagis_mm(secilen_il)
                _aylik_degerler = list(_mgm_aylik.values())
                _ortalama_aylik_yagis = (sum(_aylik_degerler) / 12.0) if len(_aylik_degerler) == 12 else None
                _en_yuksek_ay = max(_mgm_aylik, key=_mgm_aylik.get) if _mgm_aylik else None
                _en_yuksek_ay_yagis = _mgm_aylik.get(_en_yuksek_ay) if _en_yuksek_ay else None

                _yagis_yontemleri = [
                    "Günlük Toplam En Yüksek Yağış Miktarı",
                    "Ortalama Aylık Yağış Miktarı",
                    "En Yüksek Aylık Ortalama Yağış Miktarı",
                ]
                _yagis_yontemi = st.radio(
                    "Tasarım yağış verisi seçimi",
                    _yagis_yontemleri,
                    index=_yagis_yontemleri.index(st.session_state.get("yagmur_yagis_yontemi", _yagis_yontemleri[0]))
                    if st.session_state.get("yagmur_yagis_yontemi", _yagis_yontemleri[0]) in _yagis_yontemleri else 0,
                    key="yagmur_yagis_yontemi",
                )

                if _yagis_yontemi == _yagis_yontemleri[0]:
                    _otomatik_p = _mgm_yagis_mm
                    _yagis_aciklama = (
                        f"MGM günlük toplam en yüksek yağış: **{_mgm_yagis_mm:.1f} mm** "
                        f"({_mgm_yagis_tarih})" if _mgm_yagis_mm is not None else "MGM verisi alınamadı."
                    )
                elif _yagis_yontemi == _yagis_yontemleri[1]:
                    _otomatik_p = _ortalama_aylik_yagis
                    _yagis_aciklama = (
                        f"12 aylık ortalama yağışların aritmetik ortalaması: **{_ortalama_aylik_yagis:.1f} mm**"
                        if _ortalama_aylik_yagis is not None else "MGM aylık ortalama yağış verisi alınamadı."
                    )
                else:
                    _otomatik_p = _en_yuksek_ay_yagis
                    _yagis_aciklama = (
                        f"En yüksek aylık ortalama yağış: **{_en_yuksek_ay} – {_en_yuksek_ay_yagis:.1f} mm**"
                        if _en_yuksek_ay_yagis is not None else "MGM aylık ortalama yağış verisi alınamadı."
                    )

                _onceki_yagis_yontemi = st.session_state.get("yagmur_yagis_yontemi_onceki", "")
                _onceki_mgm_il = st.session_state.get("yagmur_mgm_il", "")
                if (_onceki_mgm_il != secilen_il or _onceki_yagis_yontemi != _yagis_yontemi) and _otomatik_p is not None:
                    st.session_state["yagmur_yagis"] = float(_otomatik_p)
                st.session_state["yagmur_mgm_il"] = secilen_il
                st.session_state["yagmur_yagis_yontemi_onceki"] = _yagis_yontemi

                yagmur_yagis = st.number_input(
                    "Tasarım yağış yüksekliği P (mm)", min_value=0.0,
                    step=1.0, key="yagmur_yagis",
                    help="Seçilen yönteme göre MGM verisinden otomatik gelir; istenirse proje tasarım kriterine göre manuel değiştirilebilir."
                )
                st.caption(_yagis_aciklama)
                if _yagis_yontemi != _yagis_yontemleri[0] and _mgm_aylik:
                    st.caption(
                        "MGM aylık ortalama yağışları: " +
                        " | ".join(f"{_ay}: {_deger:.1f} mm" for _ay, _deger in _mgm_aylik.items())
                    )
                if _mgm_yagis_mm is not None:
                    st.caption(f"Kaynak: MGM Resmi İklim İstatistikleri — {_mgm_yagis_url}")
            with c3:
                yagmur_akis_katsayisi = st.number_input(
                    "Akış katsayısı C", min_value=0.0, max_value=1.0,
                    value=float(st.session_state.get("yagmur_akis_katsayisi", 0.90)),
                    step=0.05, format="%.2f", key="yagmur_akis_katsayisi"
                )

            # Sarnıca alınacak yağmur suyu oranı ve filtre etkinlik katsayısı.
            # Kullanıcıya yüzde olarak gösterilir; hesapta yüzde değerleri katsayıya çevrilir.
            r1, r2 = st.columns(2)
            with r1:
                yagmur_sarnic_orani = st.number_input(
                    "Sarnıca alınacak yağmur suyu oranı (%)",
                    min_value=0.0, max_value=100.0,
                    value=float(st.session_state.get("yagmur_sarnic_orani", 80.0)),
                    step=1.0, format="%.0f", key="yagmur_sarnic_orani"
                )
            with r2:
                yagmur_filtre_etkinlik = st.number_input(
                    "Filtre etkinlik katsayısı (%)",
                    min_value=0.0, max_value=100.0,
                    value=float(st.session_state.get("yagmur_filtre_etkinlik", 90.0)),
                    step=1.0, format="%.0f", key="yagmur_filtre_etkinlik"
                )

            # Ham yağış hacmi: çatı alanı, tasarım yağışı ve akış katsayısından.
            yagmur_ham_toplanabilir_m3 = yagmur_cati_alani * yagmur_yagis * yagmur_akis_katsayisi / 1000.0

            # Sarnıç hacmi: ham yağış hacmi × sarnıca alınacak oran × filtre etkinliği.
            yagmur_toplanabilir_m3 = (
                yagmur_ham_toplanabilir_m3
                * yagmur_sarnic_orani / 100.0
                * yagmur_filtre_etkinlik / 100.0
            )

            _yr_hesap_str = (
                f"Sarnıca alınacak yağmur suyu: **V = V_ham × %{yagmur_sarnic_orani:.0f} × "
                f"%{yagmur_filtre_etkinlik:.0f} = {yagmur_toplanabilir_m3:,.2f} m³**"
                .replace(",", "X").replace(".", ",").replace("X", ".")
            )
            st.info(_yr_hesap_str)
            st.caption(
                f"Ham yağış hacmi: {yagmur_ham_toplanabilir_m3:.2f} m³ | "
                f"Sarnıca alınacak oran: %{yagmur_sarnic_orani:.0f} | "
                f"Filtre etkinliği: %{yagmur_filtre_etkinlik:.0f}"
            )

            st.markdown('<div id="bolum_631_2_2"></div>', unsafe_allow_html=True)
            st.markdown("##### • YAĞMUR SUYU DEPOSU HACİM HESABI")

            # Depo hacmi, seçilen ilin MGM aylık ortalama yağışlarının yıllık
            # toplamı esas alınarak hesaplanır. Tasarım kriteri: yıllık toplam
            # yağış hacminin %6'sı depolanacaktır.
            _mgm_yillik_yagis_mm = (
                sum(float(v) for v in (_mgm_aylik or {}).values())
                if isinstance(_mgm_aylik, dict) and len(_mgm_aylik) == 12
                else None
            )
            if _mgm_yillik_yagis_mm is not None:
                yagmur_yillik_toplam_hacim_m3 = (
                    yagmur_cati_alani * _mgm_yillik_yagis_mm * yagmur_akis_katsayisi / 1000.0
                )
                yagmur_depolama_orani = 6.0
                yagmur_gerekli_depo = yagmur_yillik_toplam_hacim_m3 * yagmur_depolama_orani / 100.0
                # Nihai depo hacmi otomatik olarak 5 m³'ün katına yukarı yuvarlanır.
                # Kullanıcı, aşağıdaki sayı alanından bu değere manuel müdahale edebilir.
                yagmur_otomatik_depo_hacmi = (
                    math.ceil(yagmur_gerekli_depo / 5.0) * 5.0
                    if yagmur_gerekli_depo > 0 else 0.0
                )

                st.write(
                    f"Yıllık toplam yağış: **{_mgm_yillik_yagis_mm:.2f} mm** "
                    f"(MGM aylık ortalamalarının toplamı)"
                )
                st.write(
                    f"Yıllık toplanabilir yağış hacmi: **{yagmur_yillik_toplam_hacim_m3:.2f} m³/yıl**"
                )
                st.write(
                    f"Depolanacak oran: **%{yagmur_depolama_orani:.0f}**"
                )
                st.write(
                    f"Hesaplanan gerekli depo hacmi: **{yagmur_gerekli_depo:.2f} m³**"
                )
                st.write(
                    f"Otomatik seçilen depo hacmi (5 m³ katına yukarı yuvarlanmış): **{yagmur_otomatik_depo_hacmi:.2f} m³**"
                )
                st.caption(
                    f"V_yıllık = A × P_yıllık × C / 1000 = "
                    f"{yagmur_cati_alani:.2f} × {_mgm_yillik_yagis_mm:.2f} × "
                    f"{yagmur_akis_katsayisi:.2f} / 1000 = "
                    f"{yagmur_yillik_toplam_hacim_m3:.2f} m³/yıl"
                )
                st.caption(
                    f"V_depo = V_yıllık × %{yagmur_depolama_orani:.0f} = "
                    f"{yagmur_yillik_toplam_hacim_m3:.2f} × %{yagmur_depolama_orani:.0f} = "
                    f"{yagmur_gerekli_depo:.2f} m³"
                )
            else:
                yagmur_yillik_toplam_hacim_m3 = 0.0
                yagmur_depolama_orani = 6.0
                yagmur_gerekli_depo = 0.0
                yagmur_otomatik_depo_hacmi = 0.0
                st.warning(
                    f"{secilen_il} için MGM'nin 12 aylık yağış verisi alınamadı. "
                    "Yıllık toplam yağışa bağlı depo hacmi hesaplanamadı."
                )

            # Üstteki Genel Bilgiler bölümünde seçilen yağmur suyu depo tipi
            # burada otomatik olarak gösterilir; burada ikinci bir tip seçimi yoktur.
            st.write(f"Yağmur suyu deposu tipi: **{sih_yagmur_depo_tipi}**")

            # Otomatik 5 m³ katı doğrudan kullanıcı müdahale alanına başlangıç
            # değeri olarak aktarılır. Kullanıcı daha önce elle değiştirmişse
            # kendi değeri korunur; yeniden hesaplanan otomatik değer zorla
            # üzerine yazılmaz.
            _yagmur_mevcut_depo = st.session_state.get("yagmur_secilen_depo", None)
            if _yagmur_mevcut_depo is None:
                _yagmur_depo_baslangic = float(yagmur_otomatik_depo_hacmi)
                st.session_state["yagmur_secilen_depo"] = _yagmur_depo_baslangic
            else:
                _yagmur_depo_baslangic = float(_yagmur_mevcut_depo)

            yagmur_secilen_depo = st.number_input(
                "Yağmur suyu deposu hacmi (m³) — elle müdahale edilebilir",
                min_value=0.0,
                step=0.5,
                key="yagmur_secilen_depo",
                help="Program hesaplanan hacmi 5 m³'ün bir üst katına otomatik yuvarlar ve bu değeri başlangıç olarak buraya aktarır. İsterseniz elle değiştirebilirsiniz."
            )

            # Yağmur suyu deposu poz bağlantısı, Genel Bilgiler / kullanma suyu
            # deposu seçiminde kullanılan aynı 25.150.xx poz-kapasite tablosuna
            # bağlanır. Betonarme depolar için bu tabloda poz bulunmadığından
            # Cihaz Poz No üretilmez.
            _yagmur_depo_poz_kapasiteleri = {
                "Paslanmaz Modüler Çelik Su Deposu": [
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
                "Galvaniz Modüler Çelik Su Deposu": [
                    (1.25, "25.150.1301"), (2.50, "25.150.1302"), (3.75, "25.150.1303"),
                    (5.00, "25.150.1304"), (6.25, "25.150.1305"), (7.50, "25.150.1306"),
                    (10.0, "25.150.1307"), (12.5, "25.150.1308"), (15.0, "25.150.1309"),
                    (20.0, "25.150.1310"), (22.5, "25.150.1311"), (25.0, "25.150.1312"),
                    (30.0, "25.150.1313"), (37.5, "25.150.1314"), (40.0, "25.150.1315"),
                    (45.0, "25.150.1316"), (50.0, "25.150.1317"), (56.0, "25.150.1318"),
                    (59.6, "25.150.1319"), (62.0, "25.150.1320"), (75.0, "25.150.1321"),
                ],
            }

            _yagmur_depo_poz_kayitlari = _yagmur_depo_poz_kapasiteleri.get(
                str(sih_yagmur_depo_tipi or "").strip(), []
            )

            def _yagmur_depo_en_yakin_poz(hedef_m3):
                if not _yagmur_depo_poz_kayitlari or hedef_m3 <= 0:
                    return None
                return min(
                    _yagmur_depo_poz_kayitlari,
                    key=lambda kayit: abs(float(kayit[0]) - float(hedef_m3)),
                )

            # Manuel nihai hacme göre Cihaz Poz No otomatik belirlenir.
            # Betonarme depoda poz listesi olmadığı için poz boş bırakılır.
            yagmur_depo_poz_kaydi = _yagmur_depo_en_yakin_poz(yagmur_secilen_depo)
            yagmur_depo_poz_kapasitesi = (
                float(yagmur_depo_poz_kaydi[0]) if yagmur_depo_poz_kaydi else None
            )
            yagmur_depo_poz = yagmur_depo_poz_kaydi[1] if yagmur_depo_poz_kaydi else ""

            # Daha önceki depo/hidrofor modüllerindeki aynı rapora aktar / kaldır
            # mantığı burada da kullanılır. Poz programda görünmeye devam eder;
            # yalnızca rapora aktarılıp aktarılmayacağı kullanıcı tarafından seçilir.
            yagmur_depo_poz_rapora_eklensin_key = "yagmur_depo_poz_rapora_eklensin_v1"
            yagmur_depo_poz_rapora_eklensin = st.checkbox(
                "Cihaz Poz Numarasını Hesap Raporuna Aktar",
                value=bool(st.session_state.get(yagmur_depo_poz_rapora_eklensin_key, True)),
                key=yagmur_depo_poz_rapora_eklensin_key,
                help=(
                    "İşaretli ise modüler paslanmaz veya modüler galvaniz su deposunun "
                    "seçilen Cihaz Poz No bilgisi hesap raporuna eklenir. "
                    "İşaret kaldırılırsa poz programda görünür ancak rapora aktarılmaz. "
                    "Betonarme su deposunda Cihaz Poz No gösterilmez."
                ),
            )

            if yagmur_depo_poz:
                st.write(f"Cihaz Poz No: **{yagmur_depo_poz}**")
            elif str(sih_yagmur_depo_tipi or "").strip() == "Betonarme Su Deposu":
                st.caption("Betonarme su deposu için tanımlı 25.150.xx Cihaz Poz No bulunmadığından poz numarası gösterilmez.")

            st.markdown('<div id="bolum_631_2_3"></div>', unsafe_allow_html=True)
            st.markdown("##### • YAĞMUR SUYU FİLTRESİ SEÇİMİ")

            # Çevre, Şehircilik ve İklim Değişikliği Bakanlığı mekanik tesisat
            # birim fiyat tariflerindeki Vortex filtre pozları. Filtre seçimi
            # yağış süresinden veya m³/h hesabından değil, doğrudan yağmur suyu
            # toplama alanından (m²) yapılır.
            YAGMUR_VORTEX_FILTRE_POZLARI = {
                "Yerüstü Vortex Filtre": [
                    {"poz": "25.181.5101", "kapasite_m2": 200,  "debi_ls": 4},
                    {"poz": "25.181.5102", "kapasite_m2": 500,  "debi_ls": 12},
                    {"poz": "25.181.5103", "kapasite_m2": 1000, "debi_ls": 25},
                    {"poz": "25.181.5104", "kapasite_m2": 3000, "debi_ls": 80},
                ],
                "Yeraltı Vortex Filtre": [
                    {"poz": "25.181.5201", "kapasite_m2": 200,  "debi_ls": 4},
                    {"poz": "25.181.5202", "kapasite_m2": 500,  "debi_ls": 12},
                    {"poz": "25.181.5203", "kapasite_m2": 1000, "debi_ls": 25},
                    {"poz": "25.181.5204", "kapasite_m2": 3000, "debi_ls": 80},
                ],
            }

            f1, f2 = st.columns(2)
            with f1:
                yagmur_filtre_tipi = st.selectbox(
                    "Vortex filtre tipi",
                    list(YAGMUR_VORTEX_FILTRE_POZLARI.keys()),
                    index=0, key="yagmur_filtre_tipi"
                )

            yagmur_filtre_pozlari = YAGMUR_VORTEX_FILTRE_POZLARI[yagmur_filtre_tipi]

            # Filtre adedi, toplam toplama alanının mevcut en büyük poz kapasitesine
            # bölünmesiyle gereken minimum adetten başlatılır. Kullanıcı adedi artırabilir.
            _en_buyuk_filtre_kapasitesi_m2 = float(yagmur_filtre_pozlari[-1]["kapasite_m2"])
            yagmur_filtre_gerekli_adet = max(
                1, int(math.ceil(yagmur_cati_alani / _en_buyuk_filtre_kapasitesi_m2))
            )
            _mevcut_filtre_adedi = int(st.session_state.get("yagmur_filtre_adet", 0) or 0)
            if _mevcut_filtre_adedi < yagmur_filtre_gerekli_adet:
                st.session_state["yagmur_filtre_adet"] = yagmur_filtre_gerekli_adet

            yagmur_filtre_adet = st.number_input(
                "Filtre adedi",
                min_value=1,
                step=1,
                value=int(st.session_state.get(
                    "yagmur_filtre_adet", yagmur_filtre_gerekli_adet
                )),
                key="yagmur_filtre_adet",
                help="Toplam toplama alanı, filtre adedine bölünür. Her filtreye düşen alanı karşılayan en küçük Bakanlık pozu otomatik seçilir. Adedi artırdığınızda poz da otomatik olarak yeniden seçilir."
            )
            yagmur_filtre_adet = int(yagmur_filtre_adet)

            # KRİTİK SEÇİM MANTIĞI:
            # Önce toplam alanı filtre adedine bölüyoruz. Poz seçimi toplam alana
            # göre değil, tek filtreye düşen alana göre yapılıyor. Böylece örneğin
            # 5.000 m² alan için 5 adet filtre seçilirse filtre başına 1.000 m²
            # düşer ve 1.000 m²'lik poz otomatik olarak seçilir.
            yagmur_filtre_basina_alan_m2 = (
                yagmur_cati_alani / yagmur_filtre_adet if yagmur_filtre_adet > 0 else 0.0
            )
            yagmur_filtre_secimi = next(
                (x for x in yagmur_filtre_pozlari
                 if yagmur_filtre_basina_alan_m2 <= float(x["kapasite_m2"])),
                yagmur_filtre_pozlari[-1]
            )
            yagmur_filtre_kapasite_m2 = float(yagmur_filtre_secimi["kapasite_m2"])
            yagmur_filtre_debisi_ls = float(yagmur_filtre_secimi["debi_ls"])
            yagmur_filtre_poz = yagmur_filtre_secimi["poz"]

            yagmur_filtre_toplam_kapasite_m2 = yagmur_filtre_kapasite_m2 * yagmur_filtre_adet
            yagmur_filtre_toplam_debisi_ls = yagmur_filtre_debisi_ls * yagmur_filtre_adet
            yagmur_filtre_kapasite_yetersiz = yagmur_cati_alani > yagmur_filtre_toplam_kapasite_m2

            with f2:
                st.metric(
                    "Tek filtre kapasitesi",
                    f"{yagmur_filtre_kapasite_m2:,.0f} m²"
                )

            st.write(
                f"Toplama alanı: **{yagmur_cati_alani:,.2f} m²** → "
                f"Filtre adedi: **{yagmur_filtre_adet} adet** → "
                f"Filtre başına düşen alan: **{yagmur_filtre_basina_alan_m2:,.2f} m²**"
            )
            st.write(
                f"Seçilen kapasite: **{yagmur_filtre_kapasite_m2:,.0f} m²/adet** → "
                f"Toplam kapasite: **{yagmur_filtre_toplam_kapasite_m2:,.0f} m²**"
            )
            st.write(
                f"Tek filtre maksimum debisi: **{yagmur_filtre_debisi_ls:.0f} L/s** → "
                f"Toplam maksimum debi: **{yagmur_filtre_toplam_debisi_ls:.0f} L/s**"
            )
            st.write(f"Cihaz Poz No: **{yagmur_filtre_poz}**")

            # Poz numarasının rapora aktarılıp aktarılmayacağı kullanıcı tarafından
            # seçilebilir. Varsayılan olarak rapora dahil edilir.
            yagmur_filtre_poz_rapora_eklensin = st.checkbox(
                "Filtre Cihaz Poz No rapora eklensin",
                value=bool(st.session_state.get("yagmur_filtre_poz_rapora_eklensin", True)),
                key="yagmur_filtre_poz_rapora_eklensin"
            )

            # Bakanlık pozlarındaki tüm filtre kapasitelerini kullanıcıya göster.
            st.markdown("**Vortex filtre poz ve kapasite tablosu:**")
            _filtre_tablo_satirlari = []
            for _tip_adi, _poz_listesi in YAGMUR_VORTEX_FILTRE_POZLARI.items():
                for _poz_kaydi in _poz_listesi:
                    _filtre_tablo_satirlari.append({
                        "Filtre Tipi": _tip_adi,
                        "Cihaz Poz No": _poz_kaydi["poz"],
                        "Toplama Alanı Kapasitesi (m²)": _poz_kaydi["kapasite_m2"],
                        "Maksimum Debi (L/s)": _poz_kaydi["debi_ls"],
                    })
            st.dataframe(
                _filtre_tablo_satirlari,
                use_container_width=True,
                hide_index=True
            )

            if yagmur_filtre_kapasite_yetersiz:
                st.warning(
                    f"Girilen {yagmur_filtre_adet} adet filtre ile toplam kapasite "
                    f"{yagmur_filtre_toplam_kapasite_m2:,.0f} m² olup, "
                    f"{yagmur_cati_alani:,.2f} m² toplama alanını karşılamıyor. "
                    f"Minimum gerekli adet: {yagmur_filtre_gerekli_adet}."
                )

            # Taşma hattı hesabında kullanılmak üzere mevcut yağış debisi hesabı
            # korunur; bu değer artık filtre seçiminde kullanılmaz.
            yagmur_sure_dk = st.number_input(
                "Tasarım yağış süresi (dk)", min_value=1.0,
                value=float(st.session_state.get("yagmur_sure_dk", 15.0)),
                step=1.0, key="yagmur_sure_dk"
            )
            yagmur_debi_m3h = yagmur_ham_toplanabilir_m3 / (yagmur_sure_dk / 60.0) if yagmur_sure_dk > 0 else 0.0

            st.markdown('<div id="bolum_631_2_4"></div>', unsafe_allow_html=True)
            st.markdown("##### • TAŞMA HATTI HESABI")
            st.caption("Taşma hattı hesabı, çatıdan oluşan ham yağmur suyu hacminin seçilen tasarım süresine dağıtılması ve ardından emniyet katsayısı uygulanmasıyla yapılır.")

            # Hesap adımları program ekranında açıkça gösterilir.
            st.markdown("**1. Ham yağmur suyu hacmi**")
            st.latex(r"V_{ham} = A \times P \times C / 1000")
            st.markdown(
                f"**Vₕₐₘ = {yagmur_cati_alani:,.2f} m² × {yagmur_yagis:,.2f} mm × {yagmur_akis_katsayisi:.2f} / 1000 = {yagmur_ham_toplanabilir_m3:,.2f} m³**"
            )

            t1, t2 = st.columns(2)
            with t1:
                tasma_emniyet = st.number_input(
                    "Taşma hattı emniyet katsayısı (%)", min_value=0.0,
                    value=float(st.session_state.get("tasma_emniyet", 20.0)),
                    step=1.0, key="tasma_emniyet"
                )
            with t2:
                tasma_debisi = yagmur_debi_m3h * (1.0 + tasma_emniyet / 100.0)
                st.metric("Hesaplanan taşma debisi", f"{tasma_debisi:.2f} m³/h")

            st.markdown("**2. Tasarım yağış debisi**")
            st.latex(r"Q_{yağış} = V_{ham} / (t / 60)")
            st.markdown(
                f"**Qᵧₐğış = {yagmur_ham_toplanabilir_m3:,.2f} m³ / ({yagmur_sure_dk:,.2f} / 60) = {yagmur_debi_m3h:,.2f} m³/h**"
            )

            st.markdown("**3. Emniyet katsayısı uygulanmış taşma debisi**")
            st.latex(r"Q_{taşma} = Q_{yağış} \times (1 + E/100)")
            st.markdown(
                f"**Qₜₐşₘₐ = {yagmur_debi_m3h:,.2f} × (1 + {tasma_emniyet:.2f}/100) = {tasma_debisi:,.2f} m³/h**"
            )

            # ------------------------------------------------------------------
            # TAŞMA HATTI HİDROLİK KONTROLÜ - MANNING
            # Cazibeli taşma hattı için tam dolu boru hidrolik kapasitesi
            # taranır ve hesaplanan taşma debisini karşılayan en küçük DN seçilir.
            # Manning katsayısı ve eğim kullanıcı tarafından değiştirilebilir.
            # ------------------------------------------------------------------
            st.markdown("**4. Taşma hattı hidrolik kontrolü (Manning yöntemi)**")
            h1, h2 = st.columns(2)
            with h1:
                tasma_malzeme = st.selectbox(
                    "Taşma hattı boru malzemesi",
                    ["PVC", "PE", "Çelik", "Beton", "Diğer"],
                    index=["PVC", "PE", "Çelik", "Beton", "Diğer"].index(
                        st.session_state.get("tasma_malzeme", "PVC")
                    ),
                    key="tasma_malzeme",
                )
            _manning_n_varsayilan = {
                "PVC": 0.011,
                "PE": 0.011,
                "Çelik": 0.012,
                "Beton": 0.015,
                "Diğer": 0.013,
            }
            with h2:
                tasma_manning_n = st.number_input(
                    "Manning pürüzlülük katsayısı (n)",
                    min_value=0.001, max_value=0.100,
                    value=float(st.session_state.get("tasma_manning_n", _manning_n_varsayilan.get(tasma_malzeme, 0.013))),
                    step=0.001, format="%.3f", key="tasma_manning_n"
                )

            h3, h4 = st.columns(2)
            with h3:
                tasma_egim_yuzde = st.number_input(
                    "Taşma hattı eğimi (%)", min_value=0.01, max_value=20.0,
                    value=float(st.session_state.get("tasma_egim_yuzde", 1.0)),
                    step=0.1, format="%.2f", key="tasma_egim_yuzde"
                )
            with h4:
                # Taşma hattı için proje tasarım kriteri: maksimum akış hızı 3,00 m/s.
                tasma_max_hiz = st.number_input(
                    "Kabul edilen maksimum hız (m/s)", min_value=0.10, max_value=20.0,
                    value=float(st.session_state.get("tasma_max_hiz", 3.0)),
                    step=0.1, format="%.1f", key="tasma_max_hiz"
                )

            # Proje tasarım kriteri: taşma hattı tasarım debisi en fazla 200 L/s kabul edilir.
            # Hesaplanan gerçek değer ayrıca gösterilir; hidrolik ön boyutlandırmada
            # kullanılan tasarım debisi 200 L/s ile sınırlandırılır.
            TASMA_MAKS_TASARIM_DEBISI_LPS = 200.0
            tasma_hesaplanan_Q_lps = tasma_debisi / 3.6
            tasma_tasarim_Q_lps = min(tasma_hesaplanan_Q_lps, TASMA_MAKS_TASARIM_DEBISI_LPS)
            tasma_tasarim_debisi_m3h = tasma_tasarim_Q_lps * 3.6
            tasma_debi_sinirlandi = tasma_hesaplanan_Q_lps > TASMA_MAKS_TASARIM_DEBISI_LPS
            # Taşma hattı adedi: tek hat yeterli değilse aynı DN sınıfında paralel
            # iki veya üç hat seçilebilir. Her hatta düşen debi ayrı ayrı kontrol edilir.
            tasma_hat_adedi = st.selectbox(
                "Taşma hattı adedi",
                [1, 2, 3],
                index=max(0, min(2, int(st.session_state.get("tasma_hat_adedi", 1)) - 1)),
                format_func=lambda x: f"{x} hat",
                key="tasma_hat_adedi",
            )

            tasma_Q_lps = tasma_tasarim_Q_lps
            tasma_Q_m3s = tasma_Q_lps / 1000.0
            tasma_hat_Q_lps = tasma_Q_lps / max(1, int(tasma_hat_adedi))
            tasma_hat_Q_m3s = tasma_hat_Q_lps / 1000.0
            tasma_S = tasma_egim_yuzde / 100.0
            tasma_dn_listesi = [50, 65, 80, 100, 125, 150, 200, 250]

            # Her hat için gerekli DN, hat adedine bölünmüş debiye göre belirlenir.
            tasma_hidrolik_tablo = []
            for _dn in tasma_dn_listesi:
                _D = _dn / 1000.0
                _A = math.pi * _D**2 / 4.0
                _R = _D / 4.0
                _Qkap = (1.0 / tasma_manning_n) * _A * (_R ** (2.0 / 3.0)) * math.sqrt(tasma_S) if tasma_S > 0 and tasma_manning_n > 0 else 0.0
                _Qkap_manning_lps = _Qkap * 1000.0
                _V_manning = _Qkap / _A if _A > 0 else 0.0
                # DN200 ve DN250 için 3,00 m/s tasarım hızına göre ayrıca kesit kapasitesi hesaplanır.
                # Manning kapasitesi ve gerçek Manning hızı ayrı olarak korunur.
                _Qkap_hiz_lps = _A * tasma_max_hiz * 1000.0 if _dn in (200, 250) else _Qkap_manning_lps
                _Qkap_kontrol_lps = max(_Qkap_manning_lps, _Qkap_hiz_lps) if _dn in (200, 250) else _Qkap_manning_lps
                _uygun = (_Qkap_kontrol_lps / 1000.0 >= tasma_hat_Q_m3s)
                tasma_hidrolik_tablo.append({
                    "dn": _dn, "alan_m2": _A,
                    "q_kapasite_lps": _Qkap_kontrol_lps,
                    "q_manning_lps": _Qkap_manning_lps,
                    "q_hiz_lps": _Qkap_hiz_lps if _dn in (200, 250) else None,
                    "hiz_ms": _V_manning, "tasarim_hiz_ms": tasma_max_hiz if _dn in (200, 250) else _V_manning,
                    "uygun": _uygun,
                    "hat_gerekli_lps": tasma_hat_Q_lps,
                    "toplam_kapasite_lps": _Qkap_kontrol_lps * int(tasma_hat_adedi),
                })

            tasma_hidrolik_secilen = next((x for x in tasma_hidrolik_tablo if x["uygun"]), None)
            # Proje tasarım kriteri: her bir paralel taşma hattı maksimum DN250 ile sınırlıdır.
            # Tek hat DN250'yi kurtarmıyorsa 2 veya 3 paralel hat ile debi bölünür.
            tasma_cap = tasma_hidrolik_secilen["dn"] if tasma_hidrolik_secilen else 200
            tasma_capasite_lps = next((x["q_kapasite_lps"] for x in tasma_hidrolik_tablo if x["dn"] == tasma_cap), 0.0)
            tasma_hiz_ms = next((x["hiz_ms"] for x in tasma_hidrolik_tablo if x["dn"] == tasma_cap), 0.0)
            tasma_hidrolik_uygun = bool(tasma_hidrolik_secilen)
            tasma_toplam_kapasite_lps = tasma_capasite_lps * int(tasma_hat_adedi)

            # 1/2/3 hat seçenekleri içinde DN250 veya daha küçük çapla
            # tasarım debisini karşılayan ilk seçenek kullanıcıya önerilir.
            tasma_onerilen_hat_adedi = None
            tasma_oneri_cap = None
            for _adet in (1, 2, 3):
                _q_hat = tasma_Q_lps / _adet
                _uygun = next(
                    (x for x in tasma_hidrolik_tablo if x["q_kapasite_lps"] >= _q_hat and x["hiz_ms"] <= tasma_max_hiz),
                    None,
                )
                if _uygun is not None:
                    tasma_onerilen_hat_adedi = _adet
                    tasma_oneri_cap = _uygun["dn"]
                    break

            st.markdown(
                f"**Hesaplanan taşma debisi:** {tasma_hesaplanan_Q_lps:.2f} L/s = {tasma_debisi:.2f} m³/h"
            )
            if tasma_debi_sinirlandi:
                st.warning(
                    f"Proje tasarım kriteri gereği hidrolik kontrolde taşma debisi "
                    f"maksimum {TASMA_MAKS_TASARIM_DEBISI_LPS:.0f} L/s ile sınırlandırılmıştır. "
                    f"Hesaplanan değer {tasma_hesaplanan_Q_lps:.2f} L/s'tir."
                )
            st.markdown(
                f"**Tasarım taşma debisi:** **{tasma_tasarim_Q_lps:.2f} L/s** "
                f"= **{tasma_tasarim_debisi_m3h:.2f} m³/h** (üst sınır: {TASMA_MAKS_TASARIM_DEBISI_LPS:.0f} L/s)"
            )
            st.markdown(
                f"**Taşma hattı adedi:** {tasma_hat_adedi} hat → her hatta düşen tasarım debisi = "
                f"{tasma_hat_Q_lps:.2f} L/s"
            )
            if tasma_onerilen_hat_adedi is not None:
                st.info(
                    f"Önerilen minimum düzen: {tasma_onerilen_hat_adedi} hat × DN {tasma_oneri_cap}. "
                    f"Kullanıcı seçimi: {tasma_hat_adedi} hat."
                )
            st.markdown("**Manning formülünün sayısal uygulanması**")
            st.latex(r"Q = \frac{1}{n} \times A \times R^{2/3} \times S^{1/2}")
            if tasma_hidrolik_secilen:
                _secili_A = float(tasma_hidrolik_secilen.get("alan_m2", 0.0))
                _secili_D = float(tasma_hidrolik_secilen.get("dn", 0)) / 1000.0
                _secili_R = _secili_D / 4.0
                _secili_Qm3s = float(tasma_hidrolik_secilen.get("q_manning_lps", 0.0)) / 1000.0
                _secili_V = float(tasma_hidrolik_secilen.get("hiz_ms", 0.0))
                st.markdown(
                    f"**A = π × D² / 4 = π × {_secili_D:.3f}² / 4 = {_secili_A:.5f} m²**  "
                    f"  \n**R = D / 4 = {_secili_D:.3f} / 4 = {_secili_R:.5f} m**"
                )
                st.markdown(
                    f"**Q = (1 / {tasma_manning_n:.3f}) × {_secili_A:.5f} × "
                    f"({_secili_R:.5f})^(2/3) × ({tasma_S:.4f})^(1/2) "
                    f"= {_secili_Qm3s:.5f} m³/s = {(_secili_Qm3s*1000):.2f} L/s**"
                )
                st.markdown(
                    f"**V = Q / A = {_secili_Qm3s:.5f} / {_secili_A:.5f} = {_secili_V:.2f} m/s**"
                )
            else:
                st.markdown(
                    f"Manning: Q = (1/n) × A × R^(2/3) × S^(1/2) → "
                    f"n = {tasma_manning_n:.3f}, S = {tasma_S:.4f} ({tasma_egim_yuzde:.2f}%), "
                    f"her hat için Q gerekli = {tasma_hat_Q_lps:.2f} L/s"
                )
            st.markdown(
                f"**Her hat için seçilen minimum taşma hattı: DN {tasma_cap}**  "
                f"→ tek hat tasarım kapasitesi = {tasma_capasite_lps:.2f} L/s, "
                f"toplam tasarım kapasitesi = {tasma_toplam_kapasite_lps:.2f} L/s, "
                f"Manning hızı = {tasma_hiz_ms:.2f} m/s"
            )
            if tasma_cap in (200, 250):
                _secili_kayit = next((x for x in tasma_hidrolik_tablo if x["dn"] == tasma_cap), None)
                if _secili_kayit and _secili_kayit.get("q_hiz_lps") is not None:
                    st.markdown(
                        f"**DN {tasma_cap} için {tasma_max_hiz:.2f} m/s tasarım hızına göre kapasite:** "
                        f"Q = A × V = {_secili_kayit['alan_m2']:.5f} × {tasma_max_hiz:.2f} "
                        f"= **{_secili_kayit['q_hiz_lps']:.2f} L/s**"
                    )

            _tablo_satirlari = []
            for _x in tasma_hidrolik_tablo:
                _tablo_satirlari.append({
                    "DN": f"DN {_x['dn']}",
                    "Tasarım Kapasitesi (L/s)": f"{_x['q_kapasite_lps']:.2f}",
                    "Manning Kapasitesi (L/s)": f"{_x.get('q_manning_lps', _x['q_kapasite_lps']):.2f}",
                    "Manning Hızı (m/s)": f"{_x['hiz_ms']:.2f}",
                    "3 m/s Kapasitesi (L/s)": (f"{_x['q_hiz_lps']:.2f}" if _x.get('q_hiz_lps') is not None else "-"),
                    "Durum": "UYGUN" if _x["uygun"] else "YETERSİZ"
                })
            st.dataframe(_tablo_satirlari, use_container_width=True, hide_index=True)
            if tasma_hidrolik_uygun:
                st.success(
                    f"Hidrolik kontrol: {tasma_hat_adedi} hat × DN {tasma_cap}; "
                    f"her hat {tasma_hat_Q_lps:.2f} L/s, toplam kapasite {tasma_toplam_kapasite_lps:.2f} L/s. "
                    f"Hız {tasma_hiz_ms:.2f} m/s ile sınır içinde."
                )
            else:
                st.warning(
                    f"{tasma_hat_adedi} hat × DN250, hat başına {tasma_hat_Q_lps:.2f} L/s tasarım debisini karşılamıyor. "
                    "Hat adedini artırın (2 veya 3 hat) veya eğim/malzeme/çıkış koşullarını yeniden değerlendirin."
                )

            st.caption("Not: Bu kontrol, taşma hattını cazibeli ve tam dolu dairesel boru kabulüyle Manning kapasitesi üzerinden ön boyutlandırır. Son proje kontrolünde gerçek kotlar, çıkış koşulu ve akış rejimi ayrıca doğrulanmalıdır.")

            st.markdown('<div id="bolum_631_2_5"></div>', unsafe_allow_html=True)
            st.markdown("##### 6.3.1.2.5 TAŞKAN SİFONU / KOKU KAPANI SEÇİMİ")

            # 25.181.5400 grubu: ÇŞİDB mekanik tesisat pozları.
            # Seçim, yukarıdaki Manning hidrolik hesabından çıkan minimum DN
            # değerini karşılayan en küçük Taşkan Sifonu pozuna otomatik bağlanır.
            TASKAN_SIFONU_POZLARI = [
                {
                    "poz": "25.181.5401",
                    "dn": 110,
                    "tanim": "Ø 110 mm Taşkan Sifonu",
                    "ozellik": (
                        "Depo taşma hattında kanalizasyondan gelebilecek biyolojik zararlıların "
                        "girişini fiziksel bariyer yapısıyla engelleyen, su yüzeyindeki polenleri "
                        "süpüren özel sifon yapısına sahip Polietilen (HDPE) taşkan sifonu."
                    ),
                },
                {
                    "poz": "25.181.5402",
                    "dn": 160,
                    "tanim": "Ø 160 mm Taşkan Sifonu",
                    "ozellik": (
                        "Depo taşma hattında kanalizasyondan gelebilecek biyolojik zararlıların "
                        "girişini fiziksel bariyer yapısıyla engelleyen, su yüzeyindeki polenleri "
                        "süpüren özel sifon yapısına sahip Polietilen (HDPE) taşkan sifonu."
                    ),
                },
                {
                    "poz": "25.181.5403",
                    "dn": 200,
                    "tanim": "Ø 200 mm Taşkan Sifonu",
                    "ozellik": (
                        "Depo taşma hattında kanalizasyondan gelebilecek biyolojik zararlıların "
                        "girişini fiziksel bariyer yapısıyla engelleyen, su yüzeyindeki polenleri "
                        "süpüren özel sifon yapısına sahip Polietilen (HDPE) taşkan sifonu."
                    ),
                },
            ]

            # Her paralel taşma hattı için bir adet Taşkan Sifonu seçilir.
            # Taşma hattı DN250 olsa dahi mevcut 25.181.5400 poz grubundaki
            # en büyük sifon Ø200 olduğundan DN250 -> Ø200 sifon olarak seçilir.
            if int(tasma_cap) >= 200:
                tasma_sifonu_secim = next(
                    (x for x in TASKAN_SIFONU_POZLARI if int(x["dn"]) == 200),
                    None,
                )
            else:
                tasma_sifonu_secim = next(
                    (x for x in TASKAN_SIFONU_POZLARI if int(x["dn"]) >= int(tasma_cap)),
                    None,
                )
            tasma_sifonu_adedi = int(tasma_hat_adedi)

            sifon1, sifon2 = st.columns(2)
            with sifon1:
                yagmur_tasma_sifonu = st.checkbox(
                    "Taşkan sifonu / koku kapanı kullanılacaktır",
                    value=True,
                    key="yagmur_tasma_sifonu",
                )
            with sifon2:
                yagmur_tasma_sifonu_poz_rapora_eklensin = st.checkbox(
                    "Taşkan Sifonu Poz No rapora eklensin",
                    value=True,
                    key="yagmur_tasma_sifonu_poz_rapora_eklensin",
                )

            yagmur_geri_tepme = st.checkbox(
                "Geri tepme önleyici düzenek", value=True, key="yagmur_geri_tepme"
            )
            yagmur_kanal_baglanti = st.checkbox(
                "Taşma hattı kanalizasyona bağlanacak",
                value=False,
                key="yagmur_kanal_baglanti",
            )

            if yagmur_tasma_sifonu:
                if tasma_sifonu_secim:
                    st.success(
                        f"Otomatik Taşkan Sifonu seçimi: {tasma_sifonu_adedi} adet × "
                        f"{tasma_sifonu_secim['poz']} — {tasma_sifonu_secim['tanim']} "
                        f"(her hatta {tasma_hat_Q_lps:.2f} L/s)"
                    )
                    st.markdown(
                        f"**Taşkan Sifonu Özelliği:** {tasma_sifonu_secim['ozellik']}"
                    )
                else:
                    st.error(
                        f"Hidrolik hesap sonucu DN {tasma_cap} gerekiyor. "
                        "25.181.5400 Taşkan Sifonu grubunda en büyük mevcut poz Ø200 mm olduğundan uygun poz bulunamadı."
                    )

            st.markdown('<div id="bolum_631_2_6"></div>', unsafe_allow_html=True)
            st.markdown("##### AKIŞ DÜZENLEYİCİ (CAZİBE YAVAŞLATICI / SAKİNLEŞTİRİCİ GİRİŞ) SEÇİMİ")
            yagmur_sakin_giris = st.checkbox(
                "Akış düzenleyici (cazibe yavaşlatıcı / sakinleştirici giriş) kullanılacaktır",
                value=True,
                key="yagmur_sakin_giris",
            )
            yagmur_sakin_giris_poz_rapora_eklensin = st.checkbox(
                "Cihaz Poz No rapora eklensin",
                value=True,
                key="yagmur_sakin_giris_poz_rapora_eklensin",
            )
            # 25.181.5300 pozunun Bakanlık yapım şartındaki tanım.
            yagmur_sakin_giris_poz = "25.181.5300"
            yagmur_sakin_giris_malzeme_ozellik = (
                "Yağmur suyu deposu girişinde suyun hızını keserek dip tortusunun havalanmasını "
                "engelleyen, suyun oksijenlenmesini destekleyen, korozyona dayanıklı Polietilen (HDPE) "
                "veya Paslanmaz Çelik malzemeden mamul akış düzenleyicinin iş yerinde temini ve yerine "
                "montajı. (Yükseltici Parça ve Kapak Fiyata Dahildir.)"
            )
            yagmur_sakin_giris_fonksiyonu = (
                "Yağmur suyu indirme borularından gelen suyun depoya hızlı bir şekilde dökülmesini "
                "engeller. Depo tabanındaki tortuların yeniden karışıp suyu bulandırmasının önüne geçer "
                "ve suyun oksijenlenmesini destekler."
            )
            if yagmur_sakin_giris:
                st.markdown(f"**Cihaz Poz No:** {yagmur_sakin_giris_poz}")
                st.markdown(f"**Malzeme / Özellik:** {yagmur_sakin_giris_malzeme_ozellik}")
                st.markdown(f"**Fonksiyonu:** {yagmur_sakin_giris_fonksiyonu}")

            st.markdown('<div id="bolum_631_2_7"></div>', unsafe_allow_html=True)
            st.markdown("##### 6.3.1.2.7 HAVALANDIRMA VE HAŞERE KORUMASI")
            h1, h2 = st.columns(2)
            with h1:
                yagmur_havalandirma = st.checkbox("Depo havalandırması yapılacaktır", value=True, key="yagmur_havalandirma")
            with h2:
                yagmur_hasere = st.checkbox("Havalandırma açıklıkları haşere korumalı olacaktır", value=True, key="yagmur_hasere")

            st.caption("Not: Yağmur suyu pompası bu bölümde seçilmez; gerekli pompa/hidrofor seçimi ilgili hidrofor-pompa modülünde yapılır.")

            yagmur_secimler = {
                "toplama": True, "filtre": True, "depo": True,
                "tasma": True, "sifon": yagmur_tasma_sifonu, "sakin_giris": yagmur_sakin_giris,
                "havalandirma": yagmur_havalandirma, "hasere": yagmur_hasere,
            }
            yagmur_hesap = {
                "cati_alani": yagmur_cati_alani, "yagis": yagmur_yagis, "akis_katsayisi": yagmur_akis_katsayisi,
                "mgm_il": secilen_il, "mgm_yagis_mm": _mgm_yagis_mm,
                "mgm_yagis_tarih": _mgm_yagis_tarih, "mgm_url": _mgm_yagis_url,
                "yagis_yontemi": _yagis_yontemi,
                "mgm_aylik_yagis": _mgm_aylik,
                "mgm_aylik_periyot": _mgm_aylik_periyot,
                "mgm_ortalama_aylik_yagis": _ortalama_aylik_yagis,
                "mgm_en_yuksek_ay": _en_yuksek_ay,
                "mgm_en_yuksek_ay_yagis": _en_yuksek_ay_yagis,
                "ham_toplanabilir_m3": yagmur_ham_toplanabilir_m3,
                "sarnic_orani": yagmur_sarnic_orani,
                "filtre_etkinlik": yagmur_filtre_etkinlik,
                "toplanabilir_m3": yagmur_toplanabilir_m3, "sure_dk": yagmur_sure_dk,
                "debi_m3h": yagmur_debi_m3h,
                "filtre_tipi": yagmur_filtre_tipi,
                "filtre_poz": yagmur_filtre_poz,
                "filtre_kapasite_m2": yagmur_filtre_kapasite_m2,
                "filtre_basina_alan_m2": yagmur_filtre_basina_alan_m2,
                "filtre_adet": yagmur_filtre_adet,
                "filtre_gerekli_adet": yagmur_filtre_gerekli_adet,
                "filtre_toplam_kapasite_m2": yagmur_filtre_toplam_kapasite_m2,
                "filtre_debisi_ls": yagmur_filtre_debisi_ls,
                "filtre_toplam_debisi_ls": yagmur_filtre_toplam_debisi_ls,
                "filtre_kapasite_yetersiz": yagmur_filtre_kapasite_yetersiz,
                "filtre_poz_rapora_eklensin": yagmur_filtre_poz_rapora_eklensin,
                "mgm_yillik_yagis_mm": _mgm_yillik_yagis_mm,
                "yillik_toplam_hacim_m3": yagmur_yillik_toplam_hacim_m3,
                "depolama_orani": yagmur_depolama_orani,
                "yagmur_depo_tipi": sih_yagmur_depo_tipi,
                "yagmur_depo_poz": yagmur_depo_poz,
                "yagmur_depo_poz_kapasitesi": yagmur_depo_poz_kapasitesi,
                "yagmur_depo_poz_rapora_eklensin": yagmur_depo_poz_rapora_eklensin,
                "gerekli_depo": yagmur_gerekli_depo,
                "otomatik_depo_hacmi": yagmur_otomatik_depo_hacmi,
                "secilen_depo": yagmur_secilen_depo,
                "tasma_emniyet": tasma_emniyet, "tasma_debisi": tasma_debisi,
                "tasma_hesaplanan_Q_lps": tasma_hesaplanan_Q_lps,
                "tasma_tasarim_Q_lps": tasma_tasarim_Q_lps,
                "tasma_tasarim_debisi_m3h": tasma_tasarim_debisi_m3h,
                "tasma_debi_sinirlandi": tasma_debi_sinirlandi,
                "tasma_maks_tasarim_Q_lps": TASMA_MAKS_TASARIM_DEBISI_LPS,
                "tasma_cap": tasma_cap,
                "tasma_hat_adedi": tasma_hat_adedi,
                "tasma_hat_Q_lps": tasma_hat_Q_lps,
                "tasma_hat_Q_m3s": tasma_hat_Q_m3s,
                "tasma_toplam_kapasite_lps": tasma_toplam_kapasite_lps,
                "tasma_onerilen_hat_adedi": tasma_onerilen_hat_adedi,
                "tasma_oneri_cap": tasma_oneri_cap,
                "tasma_Q_lps": tasma_Q_lps, "tasma_Q_m3s": tasma_Q_m3s,
                "tasma_malzeme": tasma_malzeme, "tasma_manning_n": tasma_manning_n,
                "tasma_egim_yuzde": tasma_egim_yuzde, "tasma_max_hiz": tasma_max_hiz,
                "tasma_hidrolik_kapasite_lps": tasma_capasite_lps, "tasma_hidrolik_hiz_ms": tasma_hiz_ms,
                "tasma_hidrolik_uygun": tasma_hidrolik_uygun, "tasma_hidrolik_tablo": tasma_hidrolik_tablo,
                "sifon": yagmur_tasma_sifonu,
                "tasma_sifonu_poz": tasma_sifonu_secim["poz"] if tasma_sifonu_secim else "",
                "tasma_sifonu_dn": tasma_sifonu_secim["dn"] if tasma_sifonu_secim else 0,
                "tasma_sifonu_adedi": tasma_sifonu_adedi,
                "tasma_sifonu_tanim": tasma_sifonu_secim["tanim"] if tasma_sifonu_secim else "",
                "tasma_sifonu_ozellik": tasma_sifonu_secim["ozellik"] if tasma_sifonu_secim else "",
                "tasma_sifonu_poz_rapora_eklensin": yagmur_tasma_sifonu_poz_rapora_eklensin,
                "geri_tepme": yagmur_geri_tepme,
                "kanal_baglanti": yagmur_kanal_baglanti,
                "sakin_giris": yagmur_sakin_giris,
                "sakin_giris_poz": yagmur_sakin_giris_poz,
                "sakin_giris_poz_rapora_eklensin": yagmur_sakin_giris_poz_rapora_eklensin,
                "sakin_giris_malzeme_ozellik": yagmur_sakin_giris_malzeme_ozellik,
                "sakin_giris_fonksiyonu": yagmur_sakin_giris_fonksiyonu,
                "havalandirma": yagmur_havalandirma, "hasere": yagmur_hasere,
            }
            st.session_state["yagmur_hesap"] = yagmur_hesap
            st.session_state["yagmur_secimler"] = yagmur_secimler
        else:
            st.info("Yağmur suyu hesabı pasif. Bölüm rapora dahil edilmez.")

        # SAYFA ORTA ANKORU: sağdaki "Ortaya Git" butonu buraya gelir.
        st.markdown('<div id="sayfa_orta"></div>', unsafe_allow_html=True)

        if bolum_632_aktif:
            st.markdown('<div id="bolum_632"></div>', unsafe_allow_html=True)
            st.subheader(_63_dinamik_baslik("rapor_bolum_632"))
            st.markdown("#### Genel Bilgiler ve Hidrofor Seçim Esasları")

            hidrofor_genel_keys = [f"hidrofor_genel_{i}" for i in range(1, 16)]
            _toplu_secim_butonlari(hidrofor_genel_keys)

            hidrofor_genel_maddeleri = [
                "Hidrofor sistemi, binanın kullanma suyu ihtiyacını karşılayacak ve kullanım noktalarında gerekli basıncı sağlayacak şekilde seçilecektir.",
                "Hidrofor seçiminde binanın kullanım amacı, kullanıcı sayısı, günlük su tüketimi ve eş zamanlı kullanım şartları dikkate alınacaktır.",
                "Hidroforun gerekli debisi, binanın hesaplanan anlık kullanma suyu ihtiyacına göre belirlenecektir.",
                "Hidrofor seçiminde gerekli basma yüksekliği; bina kot farkı, boru hatlarındaki sürtünme kayıpları, armatür ve ekipman kayıpları ile kullanım noktalarında gerekli minimum basınç dikkate alınarak belirlenecektir.",
                "Hidrofor sistemi, kullanım noktalarında yeterli ve kararlı su basıncı sağlayacak şekilde seçilecektir.",
                "Hidrofor pompaları, sistemin ihtiyacına göre asıl ve yedek pompa çalışma düzenine uygun olarak seçilecektir.",
                "Birden fazla pompalı sistemlerde pompaların çalışma sırası otomatik olarak münavebeli olacak şekilde düzenlenecektir.",
                "Hidrofor sistemi, düşük debilerde gereksiz pompa çalışmasını önleyecek ve değişken su tüketimlerine uyum sağlayacak şekilde tasarlanacaktır.",
                "Hidrofor sisteminde kullanılacak pompaların seçiminde pompa debisi, basma yüksekliği ve motor gücü birlikte değerlendirilecektir.",
                "Seçilen pompaların çalışma noktası, pompa performans eğrisi üzerinde hesaplanan debi ve basma yüksekliğini karşılayacak bölgede olacaktır.",
                "Hidrofor sisteminin elektrik ve otomasyon panosu, pompaların otomatik devreye girip çıkmasını, münavebeli çalışmasını ve gerekli koruma fonksiyonlarını sağlayacak şekilde tasarlanacaktır.",
                "Sistemde kullanılacak basınç tankı, kontrol ekipmanları, çekvalf, vana, basınç sensörü ve benzeri yardımcı ekipmanlar hidrofor sisteminin çalışma şartlarına uygun olarak seçilecektir.",
                "Hidrofor seçiminde minimum ve maksimum çalışma basınçları dikkate alınacaktır.",
                "Hidroforun emiş ve basma bağlantıları, sistemde gereksiz basınç kayıpları ve hidrolik sorunlar oluşturmayacak şekilde düzenlenecektir.",
                "Hidrofor sistemi, kolay bakım, işletme ve servis yapılmasına olanak sağlayacak şekilde tesis edilecektir.",
            ]

            hidrofor_genel_secimler = []
            for i, madde in enumerate(hidrofor_genel_maddeleri, start=1):
                secili = st.checkbox(
                    madde,
                    value=True,
                    key=f"hidrofor_genel_{i}",
                )
                hidrofor_genel_secimler.append(secili)

            # Aynı hidrofor türü birden fazla kez eklenebilir.
            # Her tür için adet seçilir; her bir hidrofor bağımsız başlık ön eki alabilir.
            st.markdown("#### HİDROFOR SEÇİM BÖLÜMLERİ")
            hidrofor_turleri = [
                "KULLANMA SOĞUK SUYU HİDROFORU SEÇİMİ",
                "BAHÇE SULAMA SUYU HİDROFORU SEÇİMİ",
                "YAĞMUR SUYU HİDROFORU SEÇİMİ",
            ]

            st.write("Aynı hidrofor türünden birden fazla adet seçebilirsiniz:")
            adetler = {}
            c1, c2, c3 = st.columns(3)
            for col, tur in zip((c1, c2, c3), hidrofor_turleri):
                anahtar = "hidrofor_adet_" + tur.lower().replace(" ", "_")
                with col:
                    adetler[tur] = st.number_input(
                        tur,
                        min_value=0,
                        max_value=10,
                        value=1 if tur == "KULLANMA SOĞUK SUYU HİDROFORU SEÇİMİ" else 0,
                        step=1,
                        key=anahtar,
                    )

            hidrofor_tur_listesi = []
            for tur in hidrofor_turleri:
                hidrofor_tur_listesi.extend([tur] * int(adetler[tur]))

            hidrofor_tanimlari = []
            if hidrofor_tur_listesi:
                st.markdown("##### HİDROFOR BAŞLIKLARI")
                for i, secim in enumerate(hidrofor_tur_listesi, 1):
                    st.markdown(f"**Hidrofor {i}: {secim}**")
                    on_ek = st.text_input(
                        f"6.3.{_63_dinamik_no('rapor_bolum_632')}.{i} için başlık ön eki / özel tanım (isteğe bağlı)",
                        key=f"hidrofor_on_ek_{i}",
                        placeholder="Örn.: Blok A, Blok B, Otopark, 1. Etap",
                    ).strip()
                    baslik = f"{on_ek} {secim}".strip() if on_ek else secim
                    hidrofor_tanimlari.append(baslik)

            hidrofor_sekme_bilgileri = []
            hidrofor_hesap_kayitlari = []
            if hidrofor_tanimlari:
                # Program ekranındaki sekmeler kısa kodlarla gösterilir.
                # Rapor başlıkları ise aşağıdaki hidrofor_tanimlari üzerinden tam açılımıyla yazılır.
                hidrofor_kodlari = {
                    "KULLANMA SOĞUK SUYU HİDROFORU SEÇİMİ": "KUL-HİD",
                    "BAHÇE SULAMA SUYU HİDROFORU SEÇİMİ": "BAHÇE-HİD",
                    "YAĞMUR SUYU HİDROFORU SEÇİMİ": "YAĞM-HİD",
                }

                # Aynı tür için sıra numarası kendi içinde 1'den başlar:
                # KSS1, KSS2, BSS1, YSS1 ...
                tur_sayaclari = {tur: 0 for tur in hidrofor_turleri}
                hidrofor_sekme_etiketleri = []

                for baslik in hidrofor_tanimlari:
                    # Başlık ön eki varsa, türü sondan güvenli şekilde ayırıyoruz.
                    tur = next(
                        (t for t in hidrofor_turleri if baslik.endswith(t)),
                        baslik,
                    )
                    tur_sayaclari[tur] = tur_sayaclari.get(tur, 0) + 1
                    kod = hidrofor_kodlari.get(tur, "H")
                    sekme_kodu = f"{kod} {tur_sayaclari[tur]}"
                    hidrofor_sekme_etiketleri.append(sekme_kodu)

                hidrofor_tabs = st.tabs(hidrofor_sekme_etiketleri)

                for i, (tab, baslik, sekme_kodu) in enumerate(
                    zip(hidrofor_tabs, hidrofor_tanimlari, hidrofor_sekme_etiketleri), 1
                ):
                    baslik_buyuk = baslik.upper()
                    with tab:
                        # Sekmede yalnızca kısa kod gösterilir; tam başlık raporda kullanılır.
                        st.markdown(f"### {sekme_kodu}")
                        st.caption(
                            f"Program sekmesi: {sekme_kodu} | Rapor başlığı: 6.3.{_63_dinamik_no('rapor_bolum_632')}.{i} {baslik_buyuk}"
                        )
                        st.info(
                            "Bu hidrofor bağımsız bir seçim alanıdır. Aynı türden birden "
                            "fazla hidrofor eklenebilir ve her birinin başlık ön eki ayrı "
                            "olarak değiştirilebilir."
                        )

                        # Hesap bölümü aşağıdaki ortak hidrofor hesap modülünde
                        # yükleme birimi esaslı olarak oluşturulmaktadır.

            else:
                st.info("Hidrofor seçilmedi; rapora hidrofor seçim alt başlığı eklenmez.")


            # ================================================================
            # HİDROFOR HESABI - YÜKLEME BİRİMİ ESASLI HESAP VE POMPA SEÇİMİ
            # Kaynak hesap formülasyonu: hidrofor.doc
            # ================================================================
            hidrofor_hesaplari = []

            # Hidrofor Cihaz Poz tablosu. Poz seçimi toplam pompa adedinden
            # sonra, bir asıl pompanın debisi (m³/h) ve çalışma basıncı (mSS)
            # aralığına göre otomatik yapılır. Aralıklar 2026 Mekanik Tesisat
            # Birim Fiyatlarındaki 25.160.21xx-26xx gruplarından alınmıştır.
            HIDROFOR_POZ_TABLOSU = {
                1: [
                    {"poz": "25.160.2101", "qmin": 0,  "qmax": 5,  "hmin": 20, "hmax": 40},
                    {"poz": "25.160.2102", "qmin": 0,  "qmax": 5,  "hmin": 40, "hmax": 60},
                    {"poz": "25.160.2103", "qmin": 0,  "qmax": 5,  "hmin": 60, "hmax": 80},
                    {"poz": "25.160.2104", "qmin": 5,  "qmax": 15, "hmin": 20, "hmax": 40},
                    {"poz": "25.160.2105", "qmin": 5,  "qmax": 15, "hmin": 40, "hmax": 60},
                    {"poz": "25.160.2106", "qmin": 5,  "qmax": 15, "hmin": 60, "hmax": 80},
                    {"poz": "25.160.2107", "qmin": 15, "qmax": 30, "hmin": 20, "hmax": 40},
                    {"poz": "25.160.2108", "qmin": 15, "qmax": 30, "hmin": 40, "hmax": 60},
                    {"poz": "25.160.2109", "qmin": 15, "qmax": 30, "hmin": 60, "hmax": 80},
                ],
                2: [
                    {"poz": "25.160.2201", "qmin": 0,  "qmax": 10, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2202", "qmin": 0,  "qmax": 10, "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2203", "qmin": 10, "qmax": 30, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2204", "qmin": 10, "qmax": 30, "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2205", "qmin": 30, "qmax": 60, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2206", "qmin": 30, "qmax": 60, "hmin": 60, "hmax": 90},
                ],
                3: [
                    {"poz": "25.160.2301", "qmin": 0,  "qmax": 20,  "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2302", "qmin": 0,  "qmax": 20,  "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2303", "qmin": 20, "qmax": 50,  "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2304", "qmin": 20, "qmax": 50,  "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2305", "qmin": 50, "qmax": 80,  "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2306", "qmin": 50, "qmax": 80,  "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2307", "qmin": 80,  "qmax": 120, "hmin": 60, "hmax": 90},
                ],
                4: [
                    {"poz": "25.160.2401", "qmin": 0,  "qmax": 30, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2402", "qmin": 0,  "qmax": 30, "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2403", "qmin": 30, "qmax": 60, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2404", "qmin": 30, "qmax": 60, "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2405", "qmin": 60, "qmax": 90, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2406", "qmin": 60, "qmax": 90, "hmin": 60, "hmax": 90},
                ],
                5: [
                    {"poz": "25.160.2501", "qmin": 0,   "qmax": 40,  "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2502", "qmin": 0,   "qmax": 40,  "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2503", "qmin": 40,  "qmax": 80,  "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2504", "qmin": 40,  "qmax": 80,  "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2505", "qmin": 80,  "qmax": 120, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2506", "qmin": 80,  "qmax": 120, "hmin": 60, "hmax": 90},
                ],
                6: [
                    {"poz": "25.160.2601", "qmin": 0,   "qmax": 50,  "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2602", "qmin": 0,   "qmax": 50,  "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2603", "qmin": 50,  "qmax": 100, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2604", "qmin": 50,  "qmax": 100, "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2605", "qmin": 150, "qmax": 200, "hmin": 30, "hmax": 60},
                    {"poz": "25.160.2606", "qmin": 150, "qmax": 200, "hmin": 60, "hmax": 90},
                    {"poz": "25.160.2607", "qmin": 200, "qmax": 250, "hmin": 60, "hmax": 90},
                ],
            }

            HIDROFOR_TIP_METINLERI = {
                1: "Tek pompalı, düşey milli, frekans konvertörlü santrifüj pompalı hidrofor",
                2: "İki pompalı, düşey milli, frekans konvertörlü santrifüj pompalı hidrofor",
                3: "Üç pompalı, düşey milli, frekans konvertörlü santrifüj pompalı hidrofor",
                4: "Dört pompalı, düşey milli, frekans konvertörlü santrifüj pompalı hidrofor",
                5: "Beş pompalı, düşey milli, frekans konvertörlü santrifüj pompalı hidrofor",
                6: "Altı pompalı, düşey milli, frekans konvertörlü santrifüj pompalı hidrofor",
            }

            def hidrofor_poz_sec(toplam_pompa, q_pompa_m3h, h_mss):
                """Asıl pompa debisi ve çalışma basıncına göre hidrofor Cihaz Poz No seçer."""
                kayitlar = HIDROFOR_POZ_TABLOSU.get(int(toplam_pompa), [])
                if not kayitlar:
                    return None
                q = float(q_pompa_m3h)
                h = float(h_mss)
                for kayit in kayitlar:
                    q_uygun = (q >= kayit["qmin"] and (q < kayit["qmax"] or abs(q - kayit["qmax"]) < 1e-9))
                    h_uygun = (h >= kayit["hmin"] and (h < kayit["hmax"] or abs(h - kayit["hmax"]) < 1e-9))
                    if q_uygun and h_uygun:
                        return {**kayit, "aciklama": HIDROFOR_TIP_METINLERI.get(int(toplam_pompa), "Hidrofor")}
                return None

            # Genleşme tankı standart kapasite listesi.
            # Tank adedi seçildikten sonra toplam ihtiyaç tank adedine bölünür
            # ve tank başına ihtiyaç, listedeki ilk üst kapasiteye yükseltilir.
            # Poz numaraları ayrıca tanımlandığında aynı kayıtlar üzerinden eşleştirilir.
            HIDROFOR_GENLESME_TANK_POZ_TABLOSU = [
                {"kapasite_l": 25.0, "poz": ""},
                {"kapasite_l": 50.0, "poz": ""},
                {"kapasite_l": 80.0, "poz": ""},
                {"kapasite_l": 100.0, "poz": ""},
                {"kapasite_l": 150.0, "poz": ""},
                {"kapasite_l": 200.0, "poz": ""},
                {"kapasite_l": 250.0, "poz": ""},
                {"kapasite_l": 300.0, "poz": ""},
                {"kapasite_l": 500.0, "poz": ""},
                {"kapasite_l": 750.0, "poz": ""},
                {"kapasite_l": 1000.0, "poz": ""},
                {"kapasite_l": 1500.0, "poz": ""},
                {"kapasite_l": 2000.0, "poz": ""},
                {"kapasite_l": 2500.0, "poz": ""},
                {"kapasite_l": 3000.0, "poz": ""},
            ]

            def genlesme_tanki_sec(gerekli_litre, bolme_adedi, poz_tablosu):
                """Toplam gerekli hacmi bölme adedine göre tank başına indirir.
                Poz kapasitesi her zaman tank başına gerekli hacmi karşılayan ilk
                (yukarıdaki) standart kapasitedir. Seçilen kapasite tank başına
                gösterilir; toplam kapasite ayrıca hesaplanır.
                """
                bolme_adedi = max(1, int(bolme_adedi))
                gerekli_litre = max(0.0, float(gerekli_litre))
                birim_gerekli = gerekli_litre / bolme_adedi
                kayitlar = sorted(
                    [k for k in poz_tablosu if float(k.get("kapasite_l", 0)) > 0],
                    key=lambda x: float(x["kapasite_l"]),
                )
                if not kayitlar:
                    return {"birim_gerekli_l": birim_gerekli, "secilen_birim_l": birim_gerekli, "tank_adedi": bolme_adedi, "toplam_l": birim_gerekli * bolme_adedi, "poz": ""}
                # Daima yukarı yuvarla: tank başına ihtiyaçtan küçük kapasite seçilmez.
                secilen = next((k for k in kayitlar if float(k["kapasite_l"]) >= birim_gerekli), kayitlar[-1])
                birim = float(secilen["kapasite_l"])
                return {"birim_gerekli_l": birim_gerekli, "secilen_birim_l": birim, "tank_adedi": bolme_adedi, "toplam_l": birim * bolme_adedi, "poz": secilen.get("poz", "")}

            for i, (tab, baslik, sekme_kodu) in enumerate(
                zip(hidrofor_tabs, hidrofor_tanimlari, hidrofor_sekme_etiketleri), 1
            ):
                with tab:
                    st.markdown(f"### {sekme_kodu}")
                    st.caption(f"6.3.{_63_dinamik_no("rapor_bolum_632")}.{i} {baslik}")

                    # Alt bölümler ayrı sekmeler halinde yönetilir. Her sekmenin
                    # içindeki onay kutusu, o bölümün hesaba/rapora dahil edilip
                    # edilmeyeceğini belirler. Nihai karakteristikler bölümü,
                    # üst hesaplar kapalı olsa dahi elle doldurulabilir.
                    section_tabs = st.tabs([
                        "Hidrofor Hesaplamaları",
                        "Hidrofor Tankı Hesabı",
                        "Hidroforun Karakteristikleri",
                        "Pompa Performans Eğrisi",
                    ])
                    hesaplama_aktif_key = f"hidrofor_{i}_hesaplama_aktif"
                    tank_hesabi_aktif_key = f"hidrofor_{i}_tank_hesabi_aktif"
                    karakteristik_aktif_key = f"hidrofor_{i}_karakteristik_aktif"
                    egrisi_aktif_key = f"hidrofor_{i}_egrisi_aktif"

                    # Nihai alan anahtarları; hesap sekmeleri kapalıyken de kullanılabilir.
                    vp_final_key = f"hidrofor_{i}_vp_manuel"
                    palt_final_key = f"hidrofor_{i}_palt_manuel"
                    pust_final_key = f"hidrofor_{i}_pust_manuel"
                    tank_final_key = f"hidrofor_{i}_tank_toplam_ozet"
                    guc_final_key = f"hidrofor_{i}_guc_manuel"
                    poz_final_key = f"hidrofor_{i}_poz_manuel"
                    poz_final_manual_key_for_callback = f"{poz_final_key}__manual"
                    poz_final_last_auto_key_for_callback = f"{poz_final_key}__last_auto"

                    # Varsayılan/fallback değerler: yalnızca nihai sonuç girişi yapılacak
                    # projelerde üst hesapların çalışmasına gerek kalmaz.
                    yukleme_birimi = float(st.session_state.get(f"hidrofor_{i}_yb", 150.0))
                    emniyet_orani = float(st.session_state.get(f"hidrofor_{i}_emniyet", 15.0))
                    toplam_pompa = int(st.session_state.get(f"hidrofor_{i}_toplam_pompa", 2))
                    asil_pompa = 1 if toplam_pompa <= 2 else toplam_pompa - 1
                    yedek_pompa = 0 if toplam_pompa == 1 else 1
                    vm_lph = 0.0; vm_m3h = 0.0; vp_m3h_hesap = 0.0; vp_m3h = 0.0
                    vp_pompa_m3h = float(st.session_state.get(vp_final_key, 0.0))
                    kot_farki = float(st.session_state.get(f"hidrofor_{i}_hp", 10.0))
                    akma_basinci = float(st.session_state.get(f"hidrofor_{i}_ha", 5.0))
                    boru_kaybi = float(st.session_state.get(f"hidrofor_{i}_hb", 25.0))
                    sayac_kaybi = float(st.session_state.get(f"hidrofor_{i}_hc", 5.0))
                    p_alt_mss = float(st.session_state.get(palt_final_key, 0.0))
                    p_ust_mss = float(st.session_state.get(pust_final_key, 0.0))
                    p_alt_atu = p_alt_mss / 10.0; p_ust_atu = p_ust_mss / 10.0
                    schalt = float(st.session_state.get(f"hidrofor_{i}_schalt", 20.0))
                    vn_m3 = 0.0
                    tank_bolme = int(st.session_state.get(f"hidrofor_{i}_tank_bolme", 2))
                    tank_adedi = tank_bolme; tank_birim_gerekli_litre = 0.0
                    tank_birim_litre = float(st.session_state.get(tank_final_key, 0.0)) / max(tank_adedi, 1)
                    tank_toplam_litre = float(st.session_state.get(tank_final_key, 0.0))
                    tank_poz = ""; guc = float(st.session_state.get(guc_final_key, 0.0))
                    hidrofor_poz = str(st.session_state.get(poz_final_key, "") or "").strip()
                    hidrofor_poz_aciklama = HIDROFOR_TIP_METINLERI.get(toplam_pompa, "Hidrofor")
                    # Hesaplama sekmesi kapalıyken de bu değişkenler tanımlı olmalı.
                    # Pompa eğrisi yalnızca kendi sekmesi açıldığında hesaplanır.
                    h_calisma = p_alt_mss
                    hq_egrisi = []
                    hh_egrisi = []
                    hq_baslik = ""
                    hq_model = None
                    poz_rapora_eklensin_mi_hid = bool(st.session_state.get(f"hidrofor_poz_rapor_chk_{sekme_kodu}", True))

                    # Pompa adedi değiştiğinde otomatik Cihaz Poz No yeniden seçilsin.
                    def _pompa_adedi_degisti(poz_key, manual_key, last_auto_key):
                        st.session_state[manual_key] = False
                        st.session_state[last_auto_key] = None

                    with section_tabs[0]:
                        hesaplama_aktif = st.checkbox(
                            "Bu bölümü hesaba ve rapora dahil et", value=True, key=hesaplama_aktif_key
                        )
                        if hesaplama_aktif:
                            st.markdown("#### • Hidrofor Hesaplamaları")
                            st.caption("Yükleme birimi ve işletme basınçlarına göre otomatik hidrofor hesabı.")

                            # HİDROFOR GİRDİLERİ
                            hc1, hc2, hc3 = st.columns(3)
                            with hc1:
                                yukleme_birimi = st.number_input(
                                    f"{sekme_kodu} Toplam Yükleme Birimi (YB)",
                                    min_value=0.0, value=150.0, step=1.0,
                                    key=f"hidrofor_{i}_yb",
                                )
                            with hc2:
                                emniyet_orani = st.number_input(
                                    f"{sekme_kodu} Debi emniyet oranı (%)",
                                    min_value=0.0, max_value=100.0, value=15.0, step=1.0,
                                    key=f"hidrofor_{i}_emniyet",
                                )
                            with hc3:
                                toplam_pompa = st.selectbox(
                                    f"{sekme_kodu} Toplam pompa adedi",
                                    [1, 2, 3, 4, 5, 6],
                                    index=max(0, min(5, int(st.session_state.get(f"hidrofor_{i}_toplam_pompa", 2)) - 1)),
                                    format_func=lambda x: f"{x} pompa ({0 if x == 1 else 1} yedek)",
                                    key=f"hidrofor_{i}_toplam_pompa",
                                    on_change=_pompa_adedi_degisti,
                                    args=(poz_final_key, poz_final_manual_key_for_callback, poz_final_last_auto_key_for_callback),
                                )
                            asil_pompa = 1 if toplam_pompa <= 2 else toplam_pompa - 1
                            yedek_pompa = 0 if toplam_pompa == 1 else 1

                            st.markdown("**1. Gerekli debi hesabı**")
                            vm_lph = 3600.0 * (0.25 * math.sqrt(yukleme_birimi)) if yukleme_birimi > 0 else 0.0
                            vm_m3h = vm_lph / 1000.0
                            vp_m3h_hesap = vm_m3h * (1.0 + emniyet_orani / 100.0)
                            # Emniyetli toplam debi her zaman bir üst tam sayıya yuvarlanır.
                            vp_m3h = float(math.ceil(vp_m3h_hesap))
                            vp_pompa_m3h = vp_m3h / asil_pompa if asil_pompa else 0.0
        
                            st.latex(r"V_m = 3600 \times (0,25 \times \sqrt{Z})")
                            st.write(f"Z = **{yukleme_birimi:.0f} YB**")
                            st.write(f"Vm = 3600 × (0,25 × √{yukleme_birimi:.0f}) = **{vm_lph:.0f} L/h = {vm_m3h:.2f} m³/h**")
                            st.write(f"Vp = {vm_m3h:.2f} × (1 + {emniyet_orani:.0f}/100) = {vp_m3h_hesap:.2f} m³/h → **{vp_m3h:.0f} m³/h**")
                            st.write(f"Bir pompa debisi = {vp_m3h:.2f} / {asil_pompa} = **{vp_pompa_m3h:.2f} m³/h**")
        
                            st.markdown("**2. İşletme basınçları**")
                            pc1, pc2, pc3, pc4 = st.columns(4)
                            with pc1:
                                kot_farki = st.number_input("Kot farkı hp [mSS]", min_value=0.0, value=10.0, step=0.5, key=f"hidrofor_{i}_hp")
                            with pc2:
                                akma_basinci = st.number_input("Akma basıncı ha [mSS]", min_value=0.0, value=5.0, step=0.5, key=f"hidrofor_{i}_ha")
                            with pc3:
                                boru_kaybi = st.number_input("Boru kayıpları hb [mSS]", min_value=0.0, value=25.0, step=0.5, key=f"hidrofor_{i}_hb")
                            with pc4:
                                sayac_kaybi = st.number_input("Sayaç kayıpları hc [mSS]", min_value=0.0, value=5.0, step=0.5, key=f"hidrofor_{i}_hc")
        
                            p_alt_mss = kot_farki + akma_basinci + boru_kaybi + sayac_kaybi
                            p_alt_atu = p_alt_mss / 10.0
                            p_ust_atu = p_alt_atu + 1.5
                            p_ust_mss = p_ust_atu * 10.0
                            st.write(
                                f"Pa = hp + ha + hb + hc = {kot_farki:g} + {akma_basinci:g} + "
                                f"{boru_kaybi:g} + {sayac_kaybi:g} = **{p_alt_mss:.2f} mSS = {p_alt_atu:.2f} atü**"
                            )
                            st.write(f"Pu = Pa + 1,5 = **{p_ust_atu:.2f} atü = {p_ust_mss:.2f} mSS**")
                        else:
                            st.info("Hidrofor hesaplamaları kapalı. Nihai değerleri aşağıdaki karakteristikler sekmesinden elle girebilirsiniz.")

                    with section_tabs[1]:
                        tank_hesabi_aktif = st.checkbox(
                            "Bu bölümü hesaba ve rapora dahil et", value=True, key=tank_hesabi_aktif_key
                        )
                        if tank_hesabi_aktif:
                            st.markdown("#### • Hidrofor Tankı Hesabı")
                            st.markdown("**3. Hidrofor tankı nominal hacmi**")
                            schalt = st.number_input(
                                "Şalt sayısı S [defa/h]", min_value=1.0, value=20.0, step=1.0,
                                key=f"hidrofor_{i}_schalt",
                            )
                            if (p_ust_atu - p_alt_atu) > 0 and schalt > 0:
                                vn_m3 = 0.33 * vp_pompa_m3h * ((p_ust_atu + 1.0) / ((p_ust_atu - p_alt_atu) * schalt))
                            else:
                                vn_m3 = 0.0
                            st.latex(r"V_N = 0,33 \times Q_P \times \frac{P_{ÜST}+1}{(P_{ÜST}-P_{ALT})\times S}")
                            st.write(
                                f"VN = 0,33 × {vp_pompa_m3h:.2f} × ({p_ust_atu:.2f} + 1) / "
                                f"(({p_ust_atu:.2f} - {p_alt_atu:.2f}) × {schalt:.0f}) = **{vn_m3:.3f} m³ = {vn_m3*1000:.0f} L**"
                            )
        
                            st.success(f"**Hesaplanan hidrofor tankı nominal hacmi: {vn_m3:.3f} m³ = {vn_m3*1000:.0f} L**")
        
                            st.markdown("**Hidrofor genleşme tankı bölme ve kapasite seçimi**")
                            tank_bolme = st.selectbox(
                                "Tankı kaç parçaya böleceksiniz?",
                                [1, 2, 3, 4],
                                index=1,
                                format_func=lambda x: "Tek tank" if x == 1 else f"{x}'e böl",
                                key=f"hidrofor_{i}_tank_bolme",
                            )
                            tank_secim = genlesme_tanki_sec(vn_m3 * 1000.0, tank_bolme, HIDROFOR_GENLESME_TANK_POZ_TABLOSU)
                            tank_adedi = tank_secim["tank_adedi"]
                            tank_birim_gerekli_litre = tank_secim["birim_gerekli_l"]
                            otomatik_tank_birim_litre = tank_secim["secilen_birim_l"]
                            otomatik_tank_toplam_litre = tank_secim["toplam_l"]
                            tank_poz = tank_secim["poz"]
        
                            # Üst bölümde de tank kapasitesine doğrudan elle müdahale edilebilir.
                            # Üstte yapılan değişiklik, aşağıdaki 4. bölümdeki nihai değere aktarılır
                            # ve aşağıdaki daha önceki manuel seçim sıfırlanır.
                            tank_ust_key = f"hidrofor_{i}_tank_ust_toplam"
                            tank_ust_manual_key = f"{tank_ust_key}__manual"
                            tank_ust_last_auto_key = f"{tank_ust_key}__last_auto"
                            tank_alt_manual_key = f"hidrofor_{i}_tank_toplam_ozet__manual"
        
                            if tank_ust_manual_key not in st.session_state:
                                st.session_state[tank_ust_manual_key] = False
                            if tank_ust_last_auto_key not in st.session_state:
                                st.session_state[tank_ust_last_auto_key] = otomatik_tank_toplam_litre
        
                            # Otomatik hesap değiştiyse ve üst alan daha önce elle değiştirilmediyse
                            # üst alanı güncelle.
                            if (st.session_state[tank_ust_last_auto_key] != otomatik_tank_toplam_litre
                                    and not st.session_state[tank_ust_manual_key]):
                                st.session_state[tank_ust_key] = otomatik_tank_toplam_litre
                                st.session_state[tank_ust_last_auto_key] = otomatik_tank_toplam_litre
                            elif tank_ust_key not in st.session_state:
                                st.session_state[tank_ust_key] = otomatik_tank_toplam_litre
                                st.session_state[tank_ust_last_auto_key] = otomatik_tank_toplam_litre
        
                            def _ust_tank_degisti(widget_key, alt_manual_key):
                                # Üstteki manuel değişiklik aşağıdaki nihai alanı otomatik güncellesin.
                                st.session_state[f"{widget_key}__manual"] = True
                                st.session_state[f"{widget_key}__last_auto"] = st.session_state.get(widget_key, 0.0)
                                st.session_state[alt_manual_key] = False
        
                            tank_ust_toplam_litre = st.number_input(
                                "Genleşme tankı toplam kapasitesi [L]",
                                min_value=1.0, step=50.0,
                                key=tank_ust_key,
                                on_change=_ust_tank_degisti,
                                args=(tank_ust_key, tank_alt_manual_key),
                                help="Otomatik seçilen kapasite gelir. İsterseniz burada toplam tank kapasitesini elle değiştirebilirsiniz. Bu değer aşağıdaki 4. Pompa seçim kriterleri bölümüne otomatik aktarılır.",
                            )
        
                            tank_toplam_litre = float(tank_ust_toplam_litre)
                            tank_birim_litre = tank_toplam_litre / tank_adedi if tank_adedi else tank_toplam_litre
        
                            # Manuel üst kapasite standart poz kapasitesine karşılık geliyorsa pozunu bul.
                            tank_poz = ""
                            for _kayit in HIDROFOR_GENLESME_TANK_POZ_TABLOSU:
                                if abs(float(_kayit.get("kapasite_l", 0)) - tank_birim_litre) < 0.01:
                                    tank_poz = _kayit.get("poz", "") or ""
                                    break
        
                            st.write(f"Her tank için gerekli hacim: **{tank_birim_gerekli_litre:.0f} L**")
                            st.write(f"Nihai tank: **{tank_adedi} × {tank_birim_litre:.0f} L = {tank_toplam_litre:.0f} L**")
                            st.caption("Bu kapasiteyi burada da, aşağıdaki 4. Pompa seçim kriterleri bölümünde de elle değiştirebilirsiniz.")
        
                            if tank_poz:
                                st.success(f"Genleşme tankı Cihaz Poz No: **{tank_poz}**")
                            else:
                                st.info("Genleşme tankı poz numarası, seçilen otomatik kapasite için tanımlı değil.")
                            if tank_toplam_litre < vn_m3 * 1000:
                                st.warning("Mevcut poz kapasite listesi hesaplanan toplam hacmi karşılamıyor; birim poiyat listesine daha büyük kapasite kayıtları eklenmelidir.")
                            else:
                                st.success("Seçilen toplam tank hacmi hesaplanan nominal hacmi karşılıyor.")
                        else:
                            st.info("Hidrofor tankı hesabı kapalı. Tank kapasitesini nihai karakteristikler sekmesinden elle girebilirsiniz.")
                            tank_toplam_litre = float(st.session_state.get(tank_final_key, 0.0))
                            tank_adedi = max(1, int(st.session_state.get(f"hidrofor_{i}_tank_bolme", 1)))
                            tank_birim_litre = tank_toplam_litre / tank_adedi

                    # 6.3.2 bağımsız çalışabilsin: üst hesap bölümleri kapalı olduğunda
                    # bu değişkenler hiç oluşmamış olabilir. Her zaman güvenli nihai değerleri hazırla.
                    vp_pompa_m3h = float(locals().get("vp_pompa_m3h", st.session_state.get(f"hidrofor_{i}_vp_manuel", 0.0)))
                    p_alt_mss = float(locals().get("p_alt_mss", st.session_state.get(f"hidrofor_{i}_palt_manuel", 0.0)))
                    p_ust_mss = float(locals().get("p_ust_mss", st.session_state.get(f"hidrofor_{i}_pust_manuel", 0.0)))
                    tank_toplam_litre = float(locals().get("tank_toplam_litre", st.session_state.get(f"hidrofor_{i}_tank_toplam_ozet", 0.0)))

                    with section_tabs[2]:
                        karakteristik_aktif = st.checkbox(
                            "Bu bölümü hesaba ve rapora dahil et", value=True, key=karakteristik_aktif_key
                        )
                        if karakteristik_aktif:
                            st.markdown("#### • Hidroforun Karakteristikleri")
                            st.caption("Üst hesaplar kapalı olsa bile nihai hidrofor değerlerini doğrudan elle girebilirsiniz.")
                            # Nihai değer alanları: üst bölümdeki otomatik değerler değiştiğinde,
                            # kullanıcı o alanı daha önce elle değiştirmediyse burada da otomatik
                            # olarak güncellenir. Kullanıcı elle müdahale ettiyse manuel değer korunur.
                            def _sync_final_auto(widget_key, auto_value):
                                manual_key = f"{widget_key}__manual"
                                auto_key = f"{widget_key}__last_auto"
                                if manual_key not in st.session_state:
                                    st.session_state[manual_key] = False
                                if auto_key not in st.session_state or st.session_state[auto_key] != auto_value:
                                    if not st.session_state[manual_key]:
                                        st.session_state[widget_key] = auto_value
                                    st.session_state[auto_key] = auto_value
        
                            def _mark_final_manual(widget_key):
                                st.session_state[f"{widget_key}__manual"] = True
        
                            vp_final_key = f"hidrofor_{i}_vp_manuel"
                            palt_final_key = f"hidrofor_{i}_palt_manuel"
                            pust_final_key = f"hidrofor_{i}_pust_manuel"
                            tank_final_key = f"hidrofor_{i}_tank_toplam_ozet"
                            guc_final_key = f"hidrofor_{i}_guc_manuel"
                            poz_final_key = f"hidrofor_{i}_poz_manuel"
        
                            # Yukarıdaki hesapların otomatik sonuçları.
                            otomatik_vp_pompa = float(vp_pompa_m3h)
                            otomatik_p_alt_mss = float(p_alt_mss)
                            otomatik_p_ust_mss = float(p_ust_mss)
                            # Üst bölümdeki tank değeri, aşağıdaki nihai alanın otomatik kaynağıdır.
                            otomatik_tank_toplam_litre = float(tank_toplam_litre)
        
                            # Önce nihai debi ve basınçlar senkronize edilir; motor gücü bunlara göre hesaplanır.
                            _sync_final_auto(vp_final_key, otomatik_vp_pompa)
                            _sync_final_auto(palt_final_key, otomatik_p_alt_mss)
                            _sync_final_auto(pust_final_key, otomatik_p_ust_mss)
        
                            vp_pompa_m3h = st.number_input(
                                "Pompa debisi Vp [m³/h]",
                                min_value=0.0, step=1.0, format="%.0f",
                                key=vp_final_key,
                                on_change=_mark_final_manual,
                                args=(vp_final_key,),
                            )
                            p_alt_mss = st.number_input(
                                "İşletme alt basıncı Pa [mSS]",
                                min_value=0.0, step=0.5,
                                key=palt_final_key,
                                on_change=_mark_final_manual,
                                args=(palt_final_key,),
                            )
                            p_ust_mss = st.number_input(
                                "İşletme üst basıncı Pu [mSS]",
                                min_value=0.0, step=0.5,
                                key=pust_final_key,
                                on_change=_mark_final_manual,
                                args=(pust_final_key,),
                            )
                            p_alt_atu = p_alt_mss / 10.0
                            p_ust_atu = p_ust_mss / 10.0
        
                            # GENLEŞME TANKI NİHAİ DEĞERİ
                            # Otomatik değer yukarıdaki hesaplamadan gelir. Kullanıcı bu alanı
                            # elle değiştirdiğinde girilen değer TOPLAM tank hacmi kabul edilir.
                            # Tank adedi değişirse manuel toplam değer korunur ve tank başına
                            # değer otomatik olarak yeniden bölünür.
                            _sync_final_auto(tank_final_key, otomatik_tank_toplam_litre)
                            tank_toplam_litre = st.number_input(
                                "Genleşme tankı toplam kapasitesi [L]",
                                min_value=1.0, step=50.0,
                                key=tank_final_key,
                                on_change=_mark_final_manual,
                                args=(tank_final_key,),
                                help=(
                                    "Yukarıdaki tank kapasitesi otomatik olarak buraya aktarılır. "
                                    "Burada da elle değiştirebilirsiniz; bu değer rapora gidecek nihai TOPLAM tank kapasitesidir."
                                ),
                            )
                            tank_birim_litre = tank_toplam_litre / tank_adedi if tank_adedi else tank_toplam_litre
        
                            # Alt bölümdeki manuel değer, rapora gidecek nihai değerdir. Üst bölümde
                            # yeni bir değişiklik yapılırsa üst alanın callback'i bu manuel durumu sıfırlar
                            # ve yeni üst değeri tekrar aşağıya otomatik aktarır.
        
                            # Manuel girilen toplam kapasite standart poz kapasitesine tam eşit
                            # geliyorsa, tank adedi ile bölünerek pozdaki birim kapasite bulunur.
                            # Standart listede karşılığı yoksa otomatik poz eşleşmesi korunmaz.
                            manuel_birim_poz = ""
                            for _kayit in HIDROFOR_GENLESME_TANK_POZ_TABLOSU:
                                if abs(float(_kayit.get("kapasite_l", 0)) - tank_birim_litre) < 0.01:
                                    manuel_birim_poz = _kayit.get("poz", "") or ""
                                    break
                            if st.session_state.get(f"{tank_final_key}__manual", False):
                                tank_poz = manuel_birim_poz
        
                            st.write(
                                f"Toplam pompa adedi: **{toplam_pompa} adet "
                                f"({asil_pompa} Asıl + {yedek_pompa} Yedek)**"
                            )
        
                            # Pompa elektrik gücü, nihai debi ve alt basınca göre otomatik hesaplanır.
                            # Alt bölümler bağımsız seçilebildiği için değerleri tekrar güvenceye al.
                            vp_pompa_m3h = float(st.session_state.get(vp_final_key, st.session_state.get(f"hidrofor_{i}_vp_manuel", 0.0)))
                            p_alt_mss = float(st.session_state.get(palt_final_key, st.session_state.get(f"hidrofor_{i}_palt_manuel", 0.0)))
                            h_calisma = p_alt_mss
                            hidrofor_pompa_hesap = pompa_hidrolik_hesap(vp_pompa_m3h, h_calisma, 0.60, 0.90)
                            hesaplanan_guc_kw = float(hidrofor_pompa_hesap["motor_secim_kw"])
                            _sync_final_auto(guc_final_key, hesaplanan_guc_kw)
                            guc = st.number_input(
                                "Pompa elektrik gücü [kW/adet]",
                                min_value=0.0, step=0.1,
                                key=guc_final_key,
                                on_change=_mark_final_manual,
                                args=(guc_final_key,),
                            )
        
                            # Cihaz Poz No seçimi: poz aralığındaki basınç sınıfı,
                            # İŞLETME ÜST BASINCI (Pu) esas alınarak belirlenir.
                            # Pompa gücü hesabında ise çalışma basıncı olarak alt basınç (Pa)
                            # kullanılmaya devam eder.
                            # Kullanıcı isterse hidrofor Cihaz Poz No bilgisini hesap raporuna
                            # aktarabilir; istemezse poz programda görünür ancak raporda yazılmaz.
                            poz_rapora_eklensin_mi_hid = st.checkbox(
                                "Cihaz Poz Numarasını Hesap Raporuna Aktar",
                                value=True,
                                key=f"hidrofor_poz_rapor_chk_{sekme_kodu}",
                            )
                            h_poz_secim = p_ust_mss
                            poz_kaydi = hidrofor_poz_sec(toplam_pompa, vp_pompa_m3h, h_poz_secim)
                            otomatik_hidrofor_poz = poz_kaydi["poz"] if poz_kaydi else ""
                            _sync_final_auto(poz_final_key, otomatik_hidrofor_poz)
                            hidrofor_poz = st.text_input(
                                "Cihaz Poz No",
                                key=poz_final_key,
                                on_change=_mark_final_manual,
                                args=(poz_final_key,),
                            ).strip()
                            hidrofor_poz_aciklama = (
                                poz_kaydi.get("aciklama", HIDROFOR_TIP_METINLERI.get(toplam_pompa, "Hidrofor"))
                                if poz_kaydi else
                                "Seçilen debi ve basınç için tanımlı hidrofor poz aralığı bulunamadı."
                            )
                        else:
                            st.info("Hidroforun karakteristikleri rapora dahil değil. Manuel sonuç alanları kapatıldı.")

                    with section_tabs[3]:
                        egrisi_aktif = st.checkbox(
                            "Bu bölümü hesaba ve rapora dahil et", value=True, key=egrisi_aktif_key
                        )
                        if egrisi_aktif:
                            st.markdown("#### • Pompa Performans Eğrisi:")
                            # Bu bölümdeki nihai değerler pompa eğrisine de aktarılır.
                            vp_pompa_m3h = float(st.session_state.get(vp_final_key, st.session_state.get(f"hidrofor_{i}_vp_manuel", 0.0)))
                            h_calisma = float(st.session_state.get(palt_final_key, st.session_state.get(f"hidrofor_{i}_palt_manuel", 0.0)))
                            hq_egrisi, hh_egrisi, hq_baslik, hq_model = hidrofor_pompa_secim_egrisi(
                                vp_pompa_m3h, h_calisma, hidrofor_marka_secimi
                            )
        
                            # Pis Su Pompası Seçimi'ndeki mantıkla marka/model bilgisi
                            # PROGRAM EKRANINDA açıkça gösterilir; RAPORA aktarılmaz.
                            st.markdown("**Seçilen Pompa / Üretici Verisi**")
                            if hq_model:
                                marka = hq_model.get("marka", "")
                                seri = hq_model.get("seri", "")
                                model = hq_model.get("model", "")
                                kaynak = hq_model.get("kaynak", "")
                                st.success(f"Marka: **{marka}**  |  Model: **{model}**")
                                if seri:
                                    st.write(f"Seri: **{seri}**")
                                if kaynak:
                                    st.caption(f"Kaynak: {kaynak}")
        
                            hidrofor_grafik = pompa_grafigi_png(
                                hq_egrisi, hh_egrisi, vp_pompa_m3h, h_calisma,
                                hq_baslik, anonim=False,
                            )
                            st.image(hidrofor_grafik, caption=hq_baslik, use_container_width=True)
        
                            if hidrofor_poz:
                                st.success(f"Cihaz Poz No: **{hidrofor_poz}**")
                            else:
                                st.info("Cihaz Poz No elle girilebilir veya otomatik poz eşleşmesi boş bırakılabilir.")
                        else:
                            st.info("Pompa performans eğrisi rapora dahil değil.")

                    # Rapor için tüm bölüm durumları ve nihai değerler saklanır.
                    hesaplama_aktif = bool(st.session_state.get(hesaplama_aktif_key, True))
                    tank_hesabi_aktif = bool(st.session_state.get(tank_hesabi_aktif_key, True))
                    karakteristik_aktif = bool(st.session_state.get(karakteristik_aktif_key, True))
                    egrisi_aktif = bool(st.session_state.get(egrisi_aktif_key, True))
                    hidrofor_hesaplari.append({
                        "index": i,
                        "baslik": baslik.upper(),
                        "sekme_kodu": sekme_kodu,
                        "hesaplama_aktif": hesaplama_aktif,
                        "tank_hesabi_aktif": tank_hesabi_aktif,
                        "karakteristik_aktif": karakteristik_aktif,
                        "egrisi_aktif": egrisi_aktif,
                        "toplam_yb": yukleme_birimi,
                        "vm_lph": vm_lph,
                        "vm_m3h": vm_m3h,
                        "emniyet_orani": emniyet_orani,
                        "vp_m3h": vp_m3h,
                        "asil_pompa": asil_pompa,
                        "yedek_pompa": yedek_pompa,
                        "toplam_pompa": toplam_pompa,
                        "vp_pompa_m3h": vp_pompa_m3h,
                        "hp": kot_farki,
                        "ha": akma_basinci,
                        "hb": boru_kaybi,
                        "hc": sayac_kaybi,
                        "p_alt_mss": p_alt_mss,
                        "p_alt_atu": p_alt_atu,
                        "p_ust_atu": p_ust_atu,
                        "p_ust_mss": p_ust_mss,
                        "schalt": schalt,
                        "vn_m3": vn_m3,
                        "tank_adedi": tank_adedi,
                        "tank_bolme": tank_bolme,
                        "tank_birim_gerekli_litre": tank_birim_gerekli_litre,
                        "tank_birim_litre": tank_birim_litre,
                        "tank_toplam_litre": tank_toplam_litre,
                        "tank_poz": tank_poz,
                        "h": h_calisma,
                        "guc": guc,
                        "poz": hidrofor_poz,
                        "poz_rapora_eklensin_mi": poz_rapora_eklensin_mi_hid,
                        "poz_aciklama": hidrofor_poz_aciklama,
                        "pompa_curve": list(zip(hq_egrisi, hh_egrisi)),
                        "pompa_egrisi_basligi": hq_baslik,
                        "pompa_modeli": hq_model["model"] if hq_model else "",
                        "pompa_markasi": hq_model["marka"] if hq_model else "",
                    })

        def rapor_hidrofor_alt_basligi_ekle(doc, metin):
          """Hidrofor raporundaki alt başlıkları örnekteki gibi mavi/italik biçimde ekler."""
          p = doc.add_paragraph()
          p.paragraph_format.space_before = Pt(4)
          p.paragraph_format.space_after = Pt(4)
          run = p.add_run(f"• {metin}")
          run.bold = False
          run.italic = True
          run.font.name = "Arial"
          run.font.size = Pt(12)
          run.font.color.rgb = RGBColor(68, 114, 196)
          return p

        def rapor_word_stillerini_uygula(doc):
          """Tez/rapor sayfa akışı: başlıklar bölünmez, tablo satırları bölünmez."""
          for style_name in ["Normal","Body Text","List Paragraph","List Bullet","List Number"]:
            try:
              stl = doc.styles[style_name]
              stl.font.name = "Times New Roman"
              stl._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
              stl.font.size = Pt(12)
              stl.paragraph_format.widow_control = True
            except KeyError:
              pass

          h1 = doc.styles["Heading 1"]
          h1.font.name = "Times New Roman"
          h1._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
          h1.font.size = Pt(14); h1.font.bold = True
          # Ana başlıkların sayfa davranışı aşağıda başlığın numarasına göre
          # belirlenir: 2-5 normal akışta devam eder, 6 ve sonrası yeni sayfadan başlar.
          h1.paragraph_format.page_break_before = False
          h1.paragraph_format.keep_with_next = True
          h1.paragraph_format.keep_together = True

          for level in [2,3,4]:
            h = doc.styles[f"Heading {level}"]
            h.font.name = "Times New Roman"
            h._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
            h.font.size = Pt(12); h.font.bold = True
            h.paragraph_format.keep_with_next = True
            h.paragraph_format.keep_together = True
            h.paragraph_format.widow_control = True

          for paragraph in doc.paragraphs:
            for run in paragraph.runs:
              run.font.name = "Times New Roman"
              run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
              if paragraph.style and paragraph.style.name.startswith("Heading 1"):
                run.font.size = Pt(14); run.font.bold = True
              elif paragraph.style and paragraph.style.name.startswith("Heading"):
                run.font.size = Pt(12); run.font.bold = True
              elif run.font.size is None:
                run.font.size = Pt(12)
            paragraph.paragraph_format.widow_control = True
            if paragraph.style and paragraph.style.name.startswith("Heading"):
              paragraph.paragraph_format.keep_with_next = True
              paragraph.paragraph_format.keep_together = True

              # YALNIZCA ANA BAŞLIKLAR:
              # 1-5: rapor akışı kesilmez; başlık mevcut sayfada yeterli yer yoksa
              # Word keep_with_next sayesinde başlık + ilk içerik birlikte taşınır.
              # 6, 7, 8, 9...: her zaman yeni sayfanın başından başlar.
              if paragraph.style.name == "Heading 1":
                _m = re.match(r"^\s*(\d+)\.", paragraph.text or "")
                _ana_no = int(_m.group(1)) if _m else None
                paragraph.paragraph_format.page_break_before = (
                    _ana_no is not None and _ana_no >= 6
                )

          # TABLO SAYFA KURALI:
          # - Satırlar hiçbir zaman iki sayfaya bölünmez.
          # - Tablo sayfaya sığabilecek durumdaysa tablo baştan sona aynı sayfada
          #   tutulur; sığmıyorsa Word tabloyu sonraki sayfaya taşır.
          # - Tablo öncesindeki açıklama ve alt başlık tablo ile birlikte taşınır.
          # - Çok uzun tablolarda tablo sayfalar arasında devam edebilir; ancak
          #   satır bölünmez ve ilk satır her sayfada başlık olarak tekrarlanır.
          for table in doc.tables:
            _satir_sayisi = len(table.rows)

            for row_index, row in enumerate(table.rows):
              trPr = row._tr.get_or_add_trPr()

              # Satırın sayfalar arasında bölünmesini kesin olarak engelle.
              if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))

              # İlk satır sonraki satırla birlikte kalsın.
              # 25 satıra kadar olan tabloların tamamını mümkün olduğunca
              # tek sayfada tutuyoruz.
              _satir_keep = (
                  row_index < _satir_sayisi - 1
                  if _satir_sayisi <= 25
                  else row_index == 0
              )

              for cell in row.cells:
                for paragraph in cell.paragraphs:
                  paragraph.paragraph_format.widow_control = True
                  if _satir_keep:
                    paragraph.paragraph_format.keep_with_next = True
                    _pPr = paragraph._p.get_or_add_pPr()
                    if _pPr.find(qn("w:keepNext")) is None:
                      _pPr.append(OxmlElement("w:keepNext"))
                  paragraph.paragraph_format.keep_together = True

              # Uzun tablolar bir sonraki sayfada da kolon başlıklarını
              # tekrar etsin.
              if row_index == 0:
                _tbl_header = trPr.find(qn("w:tblHeader"))
                if _tbl_header is None:
                  _tbl_header = OxmlElement("w:tblHeader")
                  trPr.append(_tbl_header)

            # Tabloya hemen önceki paragrafı bul. Bu paragraf açıklama ise
            # tabloyla; başlık ise açıklama ve tabloyla birlikte taşınır.
            _body = table._tbl.getparent()
            _idx = _body.index(table._tbl)
            _onceki_paragraflar = []

            _j = _idx - 1
            while _j >= 0 and len(_onceki_paragraflar) < 2:
              _el = _body[_j]
              if _el.tag == qn("w:p"):
                for _p in doc.paragraphs:
                  if _p._p is _el:
                    _onceki_paragraflar.append(_p)
                    break
              elif _el.tag == qn("w:tbl"):
                break
              _j -= 1

            # Tablo öncesindeki son paragraf ve onun hemen üstündeki başlık/
            # açıklama tabloyla birlikte kalsın.
            for _p in _onceki_paragraflar:
              _p.paragraph_format.keep_with_next = True
              _p.paragraph_format.keep_together = True
              _p.paragraph_format.widow_control = True
              _pPr = _p._p.get_or_add_pPr()
              if _pPr.find(qn("w:keepNext")) is None:
                _pPr.append(OxmlElement("w:keepNext"))

          try:
            settings = doc.settings.element
            upd = settings.find(qn("w:updateFields"))
            if upd is None:
              upd = OxmlElement("w:updateFields")
              settings.append(upd)
            upd.set(qn("w:val"), "true")
          except Exception:
            pass


    def re_sirk_formul_gorseli_png(q_boyler_kcal_h, q_temsiz_m3h, q_emniyetli_m3h, q_secim_m3h, q_boyler_kw=None, emniyet_orani=15.0):
        """6.3.4 debi hesabını kompakt ve örnek görseldeki düzende üretir."""
        if q_boyler_kw is None:
            q_boyler_kw = float(q_boyler_kcal_h) * 0.001163

        import matplotlib.pyplot as plt

        factor = 1.0 + float(emniyet_orani) / 100.0
        factor_txt = f"{factor:.2f}".replace(".", ",")
        q_txt = f"{float(q_boyler_kcal_h):,.0f}".replace(",", ".")
        kw_txt = f"{float(q_boyler_kw):.0f}".replace(".", ",")
        qh_txt = f"{float(q_emniyetli_m3h):.2f}".replace(".", ",")

        # Sıkı yerleşim: önceki sürümde oluşan büyük boşlukların tamamı kaldırıldı.
        fig = plt.figure(figsize=(7.2, 1.85), facecolor="white")
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis("off")

        ax.text(0.04, 0.82,
                rf"$Q_{{BOYLER}} = {q_txt}\;kcal/h \approx {kw_txt}\;kW$",
                ha="left", va="center", fontsize=13, color="black")

        ax.text(0.075, 0.43, r"$V =$",
                ha="left", va="center", fontsize=14, color="black")
        ax.text(0.30, 0.57,
                rf"${q_txt} \times 0,05 \times {factor_txt}$",
                ha="center", va="center", fontsize=13, color="black")
        ax.plot([0.18, 0.42], [0.43, 0.43], color="black", linewidth=0.8)
        ax.text(0.30, 0.29, r"$5.000$",
                ha="center", va="center", fontsize=13, color="black")
        ax.text(0.48, 0.43,
                rf"$= {qh_txt}\;m^3/h$",
                ha="left", va="center", fontsize=13, color="black")

        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=240, bbox_inches="tight", pad_inches=0.02, facecolor="white")
        plt.close(fig)
        buf.seek(0)
        return buf

    # 6.3.4 Re-sirkülasyon pompası Cihaz Poz No tablosu
    # 2026 ÇŞİDB mekanik tesisat poz aralıkları esas alınmıştır.
    RE_SIRK_POMPA_POZ_TABLOSU = [
        {"poz":"25.350.3001", "qmin":0.5, "qmax":3.5, "hmin":1.0, "hmax":3.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (0,5-3,5 m³/h) (1-3 mSS)."},
        {"poz":"25.350.3002", "qmin":3.5, "qmax":7.0, "hmin":1.0, "hmax":3.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (3,5-7,0 m³/h) (1-3 mSS)."},
        {"poz":"25.350.3003", "qmin":7.0, "qmax":11.0, "hmin":1.0, "hmax":3.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (7-11 m³/h) (1-3 mSS)."},
        {"poz":"25.350.3004", "qmin":3.0, "qmax":6.0, "hmin":3.0, "hmax":5.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (3-6 m³/h) (3-5 mSS)."},
        {"poz":"25.350.3005", "qmin":6.0, "qmax":9.0, "hmin":3.0, "hmax":5.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (6-9 m³/h) (3-5 mSS)."},
        {"poz":"25.350.3006", "qmin":9.0, "qmax":12.0, "hmin":3.0, "hmax":5.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (9-12 m³/h) (3-5 mSS)."},
        {"poz":"25.350.3007", "qmin":12.0, "qmax":17.0, "hmin":3.0, "hmax":5.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (12-17 m³/h) (3-5 mSS)."},
        {"poz":"25.350.3008", "qmin":12.0, "qmax":20.0, "hmin":5.0, "hmax":10.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (12-20 m³/h) (5-10 mSS)."},
        {"poz":"25.350.3009", "qmin":20.0, "qmax":28.0, "hmin":5.0, "hmax":10.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (20-28 m³/h) (5-10 mSS)."},
        {"poz":"25.350.3010", "qmin":28.0, "qmax":36.0, "hmin":5.0, "hmax":10.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (28-36 m³/h) (5-10 mSS)."},
        {"poz":"25.350.3011", "qmin":36.0, "qmax":50.0, "hmin":5.0, "hmax":10.0, "tanim":"Değişken devirli (frekans konvertörlü) ıslak rotorlu sirkülasyon pompası (36-50 m³/h) (5-10 mSS)."},
    ]

    # Wilo / Grundfos program içi seçim verileri. Marka-model ekranda gösterilir; rapora aktarılmaz.
    RE_SIRK_URETICI_POMPALAR = [
        {"marka":"Wilo", "model":"Wilo-Stratos MAXO-Z 25/0,5-6", "p2_kw":0.20, "curve":[(0.0,6.0),(1.0,5.8),(2.0,5.3),(3.0,4.6),(4.0,3.6),(5.0,2.5)], "kaynak":"Wilo"},
        {"marka":"Wilo", "model":"Wilo-Stratos MAXO-Z 30/0,5-8", "p2_kw":0.25, "curve":[(0.0,8.0),(1.0,7.7),(2.0,7.2),(3.0,6.5),(4.0,5.6),(5.0,4.5),(6.0,3.2)], "kaynak":"Wilo"},
        {"marka":"Grundfos", "model":"Grundfos MAGNA3 25-60", "p2_kw":0.18, "curve":[(0.0,6.0),(1.0,5.8),(2.0,5.2),(3.0,4.5),(4.0,3.6),(5.0,2.5)], "kaynak":"Grundfos"},
        {"marka":"Grundfos", "model":"Grundfos MAGNA3 25-80", "p2_kw":0.25, "curve":[(0.0,8.0),(1.0,7.7),(2.0,7.1),(3.0,6.4),(4.0,5.5),(5.0,4.4),(6.0,3.1)], "kaynak":"Grundfos"},
    ]

    def re_sirk_pompa_pozu_sec(q_m3h, h_mss):
        # Önce H çalışma bandını sağlayan poz grubu bulunur; ardından gerekli Q'yu
        # karşılayan en küçük poz seçilir. Böylece örneğin Q=2 m³/h, H=5 mSS
        # için H=3-5 bandındaki 25.350.3004 seçilebilir.
        h_uygun = [k for k in RE_SIRK_POMPA_POZ_TABLOSU if k["hmin"] <= h_mss <= k["hmax"]]
        if not h_uygun:
            return "SINIR DIŞI", "Girilen H çalışma noktası 25.350.3001-3011 poz basınç aralıkları dışındadır!", "GECERSIZ"

        q_uygun = [k for k in h_uygun if q_m3h <= k["qmax"]]
        if q_uygun:
            secilen = min(q_uygun, key=lambda k: k["qmax"])
            return secilen["poz"], secilen["tanim"], "UYGUN"

        return "SINIR DIŞI", "Girilen Q çalışma noktası 25.350.3001-3011 poz kapasite sınırları dışındadır!", "GECERSIZ"

    def re_sirk_egri_degeri(curve, q):
        if not curve or q < curve[0][0] or q > curve[-1][0]:
            return None
        for i in range(len(curve)-1):
            q0,h0=curve[i]; q1,h1=curve[i+1]
            if q0 <= q <= q1:
                if q1 == q0: return h0
                oran=(q-q0)/(q1-q0)
                return h0 + oran*(h1-h0)
        return curve[-1][1]

    def re_sirk_uretici_sec(q_m3h, h_mss, marka_secimi="Otomatik (Wilo + Grundfos)"):
        adaylar=[]
        for pompa in RE_SIRK_URETICI_POMPALAR:
            if marka_secimi == "Wilo" and pompa["marka"] != "Wilo": continue
            if marka_secimi == "Grundfos" and pompa["marka"] != "Grundfos": continue
            h_egri = re_sirk_egri_degeri(pompa["curve"], q_m3h)
            if h_egri is None or h_egri < h_mss: continue
            adaylar.append((h_egri-h_mss, pompa["p2_kw"], pompa))
        if not adaylar:
            return None
        adaylar.sort(key=lambda x:(x[0],x[1]))
        secilen=dict(adaylar[0][2])
        secilen["h_calisma"] = re_sirk_egri_degeri(secilen["curve"], q_m3h)
        return secilen

    def re_sirk_pompa_secim(q_m3h, h_mss, marka_secimi="Otomatik (Wilo + Grundfos)"):
        poz, tanim, durum = re_sirk_pompa_pozu_sec(q_m3h, h_mss)
        model = re_sirk_uretici_sec(q_m3h, h_mss, marka_secimi) if durum == "UYGUN" else None
        if model:
            return list(model["curve"]), model, poz, tanim, durum
        return [], None, poz, tanim, durum

    # --- 6.3.3 KULLANMA SICAK SUYU İHTİYACI HESAPLARI ---
    # Kaynak: kullanıcı tarafından yüklenen BOYLER SEÇİMİ.xlsx / Sayfa1.
    # Excel'deki alt-üst tüketim aralığı programda bilgi amaçlı korunur.
    # Başlangıç birim tüketim değerleri, kullanıcının ayrı ayrı sağladığı
    # yapı tipi tablolarındaki ortalama tüketim değerlerinden alınmıştır.
    # Kullanıcı bu değerleri tabloda manuel olarak değiştirebilir.
    SICAK_SU_EXCEL_VERILERI = {
        "Bağımsız Ev": {
            "Özel Lavabo": "7,5 -9", "Genel Lavabo": "-", "Banyo": "90-250",
            "Bulaşık Makinası": "40-68", "Evye": "35-45", "Çamaşır Teknesi": "70-90",
            "Çamaşır Makinası": "70-90", "Duş": "136-250",
        },
        "Apartman": {
            "Özel Lavabo": "7,5 -9", "Genel Lavabo": "15-18", "Banyo": "76-250",
            "Bulaşık Makinası": "40-68", "Evye": "35-45", "Çamaşır Teknesi": "70-90",
            "Çamaşır Makinası": "70-90", "Duş": "114-250",
        },
        "Hastane": {
            "Özel Lavabo": "7,5 -9", "Genel Lavabo": "20-27", "Banyo": "76-250",
            "Bulaşık Makinası": "160-680", "Evye": "70-90", "Çamaşır Teknesi": "75-126",
            "Çamaşır Makinası": "75-126", "Duş": "250-340",
        },
        "Otel": {
            "Özel Lavabo": "7,5 -9", "Genel Lavabo": "30-36", "Banyo": "76-250",
            "Bulaşık Makinası": "160-760", "Evye": "70-136", "Çamaşır Teknesi": "75-126",
            "Çamaşır Makinası": "75-126", "Duş": "250-340",
        },
        "İşyeri": {
            "Özel Lavabo": "7,5 -9", "Genel Lavabo": "23-27", "Banyo": "-",
            "Bulaşık Makinası": "-", "Evye": "35-90", "Çamaşır Teknesi": "-",
            "Çamaşır Makinası": "-", "Duş": "114-136",
        },
        "Okul": {
            "Özel Lavabo": "7,5 -9", "Genel Lavabo": "50-68", "Banyo": "-",
            "Bulaşık Makinası": "75-450", "Evye": "35-90", "Çamaşır Teknesi": "-",
            "Çamaşır Makinası": "-", "Duş": "250-1000",
        },
        "Endüstriyel Tesis": {
            "Özel Lavabo": "7,5 -9", "Genel Lavabo": "40-54", "Banyo": "-",
            "Bulaşık Makinası": "75-450", "Evye": "70-90", "Çamaşır Teknesi": "-",
            "Çamaşır Makinası": "-", "Duş": "750-1000",
        },
    }

    # Kullanıcının ayrı ayrı sağladığı yapı tipi tablolarındaki ortalama birim
    # tüketim değerleri. Program bu değerleri doğrudan başlangıç değeri olarak
    # kullanır; kullanıcı yine tabloda manuel olarak değiştirebilir.
    SICAK_SU_ORNEK_ORTALAMA_TUKETIMLERI = {
        "Bağımsız Ev": {
            "Özel Lavabo": 9, "Genel Lavabo": 0, "Banyo": 150,
            "Bulaşık Makinası": 50, "Evye": 45, "Çamaşır Teknesi": 80,
            "Çamaşır Makinası": 80, "Duş": 150,
        },
        "Apartman": {
            "Özel Lavabo": 9, "Genel Lavabo": 15, "Banyo": 150,
            "Bulaşık Makinası": 50, "Evye": 40, "Çamaşır Teknesi": 80,
            "Çamaşır Makinası": 80, "Duş": 150,
        },
        "Hastane": {
            "Özel Lavabo": 9, "Genel Lavabo": 25, "Banyo": 250,
            "Bulaşık Makinası": 200, "Evye": 75, "Çamaşır Teknesi": 100,
            "Çamaşır Makinası": 100, "Duş": 250,
        },
        "Otel": {
            "Özel Lavabo": 9, "Genel Lavabo": 30, "Banyo": 150,
            "Bulaşık Makinası": 200, "Evye": 75, "Çamaşır Teknesi": 75,
            "Çamaşır Makinası": 75, "Duş": 250,
        },
        "İşyeri": {
            "Özel Lavabo": 9, "Genel Lavabo": 25, "Banyo": 0,
            "Bulaşık Makinası": 0, "Evye": 70, "Çamaşır Teknesi": 0,
            "Çamaşır Makinası": 0, "Duş": 120,
        },
        "Okul": {
            "Özel Lavabo": 9, "Genel Lavabo": 50, "Banyo": 0,
            "Bulaşık Makinası": 100, "Evye": 50, "Çamaşır Teknesi": 0,
            "Çamaşır Makinası": 0, "Duş": 250,
        },
        "Endüstriyel Tesis": {
            "Özel Lavabo": 9, "Genel Lavabo": 40, "Banyo": 0,
            "Bulaşık Makinası": 100, "Evye": 75, "Çamaşır Teknesi": 0,
            "Çamaşır Makinası": 0, "Duş": 750,
        },
    }


    KULLANMA_ES_FAKTORLERI = {
        "Bağımsız Ev": 0.30,
        "Apartman": None,
        "Hastane": 0.25,
        "Otel": 0.25,
        "İşyeri": 0.30,
        "Okul": 0.40,
        "Endüstriyel Tesis": 0.40,
    }
    DEPOLAMA_FAKTORLERI = {
        "Bağımsız Ev": 0.70,
        "Apartman": 1.25,
        "Hastane": 0.60,
        "Otel": 0.80,
        "İşyeri": 2.00,
        "Okul": 1.00,
        "Endüstriyel Tesis": 1.00,
    }
    ES_ZAMAN_FAKTORLERI = {
        1: 1.00, 2: 0.75, 3: 0.60, 4: 0.58, 5: 0.55, 6: 0.54, 7: 0.51,
        8: 0.49, 10: 0.45, 15: 0.42, 18: 0.40, 20: 0.38, 25: 0.36, 30: 0.34,
        40: 0.32, 50: 0.31, 60: 0.30, 70: 0.30, 80: 0.30, 90: 0.30, 100: 0.30,
    }

    def sicak_su_aralik_ortalama_10(aralik):
        """Excel alt-üst aralığının ortalamasını en yakın 10 L'ye yarım-yukarı yuvarlar."""
        if not aralik or str(aralik).strip() == "-":
            return None
        metin = str(aralik).strip().replace("–", "-").replace("—", "-").replace(" ", "")
        try:
            parcalar = metin.split("-")
            if len(parcalar) != 2:
                return None
            def sayiya_cevir(x):
                return float(str(x).replace(",", "."))
            alt = sayiya_cevir(parcalar[0])
            ust = sayiya_cevir(parcalar[1])
            ort = (alt + ust) / 2.0
            return int((ort + 5.0) // 10.0) * 10
        except (TypeError, ValueError):
            return None

    sicak_su_hesap_detaylari = []
    sicak_su_gunluk_toplam_litre = 0.0
    sicak_su_yapi_tipi = "Bağımsız Ev"

    if bolum_633_aktif:
        st.markdown('<div id="bolum_633"></div>', unsafe_allow_html=True)
        st.markdown(f"### • {_63_dinamik_baslik("rapor_bolum_633")}")
        st.caption("Kaynak tüketim değerleri: BOYLER SEÇİMİ.xlsx / Sayfa1")
        sicak_su_yapi_tipi = st.selectbox(
            "Yapı / kullanım tipi",
            list(SICAK_SU_EXCEL_VERILERI.keys()),
            key="sicak_su_yapi_tipi",
        )
        kaynak_satirlari = SICAK_SU_EXCEL_VERILERI[sicak_su_yapi_tipi]

        # Kullanım yerleri artık seçim kutusu yerine sabit ve düzenlenebilir bir tablodur.
        # Kullanıcı yalnızca Adet ve Birim Tüketim değerlerini değiştirir.
        # Adet = 0 olan satırlar hesaba dahil edilmez.
        tablo_key = f"sicak_su_kullanim_tablosu_v32_{sicak_su_yapi_tipi}"
        satirlar = []
        for kullanim, kaynak_aralik in kaynak_satirlari.items():
            # Başlangıç değeri doğrudan kullanıcının ilgili yapı tipi için
            # sağladığı ortalama tüketim tablosundan alınır.
            birim = float(SICAK_SU_ORNEK_ORTALAMA_TUKETIMLERI[sicak_su_yapi_tipi].get(kullanim, 0))
            satirlar.append({
                "Kullanım Yeri": kullanim,
                "Standart Aralığı [L]": str(kaynak_aralik),
                "Birim Tüketim [L]": birim,
                "Adet": 0,
            })

        # Excel'de bulunmayan fakat projelerde kullanılabilecek Engelli satırı.
        # Kaynak değeri olmadığı için başlangıç değeri 0 L'dir ve kullanıcı elle girer.
        satirlar.append({
            "Kullanım Yeri": "Engelli",
            "Standart Aralığı [L]": "Kaynakta yok",
            "Birim Tüketim [L]": 0.0,
            "Adet": 0,
        })

        varsayilan_df = pd.DataFrame(satirlar)
        veri_key = f"sicak_su_tablo_veri_{sicak_su_yapi_tipi}"
        editor_key = f"sicak_su_editor_{sicak_su_yapi_tipi}"

        if veri_key not in st.session_state:
            st.session_state[veri_key] = varsayilan_df.copy()

        edited_df = st.data_editor(
            st.session_state[veri_key],
            key=editor_key,
            hide_index=True,
            num_rows="fixed",
            use_container_width=True,
            column_config={
                "Kullanım Yeri": st.column_config.TextColumn("Kullanım Yeri", disabled=True),
                "Standart Aralığı [L]": st.column_config.TextColumn("Standart Aralığı [L]", disabled=True),
                "Birim Tüketim [L]": st.column_config.NumberColumn(
                    "Birim Tüketim [L]", min_value=0, step=10, format="%.0f L"
                ),
                "Adet": st.column_config.NumberColumn(
                    "Adet", min_value=0, step=1, format="%.0f"
                ),
            },
        )

        # Widget'ın kendi anahtarına yazmak Streamlit tarafından yasaktır.
        # Düzenlenen veriyi ayrı bir session-state anahtarında tutuyoruz.
        st.session_state[veri_key] = edited_df.copy()

        # Hesaplanan toplamları ayrı ve okunaklı sonuç tablosunda göster.
        sonuc_df = edited_df.copy()
        sonuc_df["Toplam [L/gün]"] = (
            pd.to_numeric(sonuc_df["Birim Tüketim [L]"], errors="coerce").fillna(0)
            * pd.to_numeric(sonuc_df["Adet"], errors="coerce").fillna(0)
        )
        sonuc_df["Birim Tüketim [L]"] = pd.to_numeric(sonuc_df["Birim Tüketim [L]"], errors="coerce").fillna(0).round(0).astype(int)
        sonuc_df["Adet"] = pd.to_numeric(sonuc_df["Adet"], errors="coerce").fillna(0).round(0).astype(int)
        sonuc_df["Toplam [L/gün]"] = pd.to_numeric(sonuc_df["Toplam [L/gün]"], errors="coerce").fillna(0).round(0).astype(int)

        aktif_sonuc_df = sonuc_df[sonuc_df["Adet"] > 0].copy()
        sicak_su_gunluk_toplam_litre = float(sonuc_df["Toplam [L/gün]"].sum())

        # BOYLER SEÇİMİ.xlsx / Sayfa1 içindeki faktörler.
        # Kullanıcı tarafından verilen hesap bağıntısı:
        # V = Kullanma Eş Zaman Faktörü x Depolama Faktörü x Toplam Tüketim
        # Sonuç emniyetli seçim için 50 L'nin bir üst katına yuvarlanır.
        # 0 L sonucu 0 L olarak korunur.
        kullanma_es_faktoru_kaynak = KULLANMA_ES_FAKTORLERI.get(sicak_su_yapi_tipi)
        depolama_faktoru_kaynak = DEPOLAMA_FAKTORLERI.get(sicak_su_yapi_tipi)
        es_zaman_faktoru = None
        konut_sayisi = None

        if sicak_su_yapi_tipi == "Apartman":
            konut_sayisi = st.selectbox(
                "Konut / daire sayısı (Excel eş zaman faktörü tablosu)",
                list(ES_ZAMAN_FAKTORLERI.keys()),
                key="sicak_su_konut_sayisi",
            )
            es_zaman_faktoru = ES_ZAMAN_FAKTORLERI[konut_sayisi]
            varsayilan_es = float(es_zaman_faktoru)
        else:
            varsayilan_es = float(kullanma_es_faktoru_kaynak or 0.0)

        # Faktör widget'larını yapı tipine göre sıfırlamak için sabit widget anahtarları
        # kullanıyoruz. Streamlit'te bir widget key'i mevcutsa, sonradan verilen
        # `value=` varsayılanı mevcut değeri değiştirmez. Bu nedenle yapı tipi veya
        # apartman daire sayısı değiştiğinde Session State'i açıkça güncelliyoruz.
        faktor_signature = (
            f"Apartman_{konut_sayisi}" if sicak_su_yapi_tipi == "Apartman"
            else sicak_su_yapi_tipi
        )
        faktor_onceki_signature = st.session_state.get("sicak_su_faktor_signature_v34")
        faktor_key_es = "sicak_su_kullanma_es_faktoru_v34"
        faktor_key_dep = "sicak_su_depolama_faktoru_v34"

        if faktor_onceki_signature != faktor_signature:
            st.session_state[faktor_key_es] = float(varsayilan_es)
            st.session_state[faktor_key_dep] = float(depolama_faktoru_kaynak or 0.0)
            st.session_state["sicak_su_faktor_signature_v34"] = faktor_signature
        else:
            # İlk çalıştırmada değerlerin mutlaka kaynak tablodan gelmesini garanti et.
            if faktor_key_es not in st.session_state:
                st.session_state[faktor_key_es] = float(varsayilan_es)
            if faktor_key_dep not in st.session_state:
                st.session_state[faktor_key_dep] = float(depolama_faktoru_kaynak or 0.0)

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            kullanma_es_faktoru = st.number_input(
                "Kullanım Eş Zaman Faktörü",
                min_value=0.0, max_value=1.0, step=0.01,
                format="%.2f", key=faktor_key_es
            )
        with col_f2:
            depolama_faktoru = st.number_input(
                "Depolama Faktörü",
                min_value=0.0, max_value=10.0, step=0.01,
                format="%.2f", key=faktor_key_dep
            )

        st.markdown("**Ortalama Ani Sıcak Su İhtiyacı:**")
        toplam_tuketim = float(sicak_su_gunluk_toplam_litre)
        es_zamanli_tuketim = toplam_tuketim * float(kullanma_es_faktoru)
        hesaplanan_boyler_hacmi = es_zamanli_tuketim * float(depolama_faktoru)
        secilen_boyler_hacmi_hesaplanan = int(math.ceil(hesaplanan_boyler_hacmi / 50.0) * 50) if hesaplanan_boyler_hacmi > 0 else 0

        # Boyler hacmi ile aşağıdaki sekonder su debisi iki yönlü senkron çalışır.
        # İlk açılışta üstte hesaplanan emniyetli hacim aşağıya aktarılır.
        # Kullanıcı aşağıdaki değeri değiştirirse üstteki emniyetli değer de aynı değere döner.
        boyler_sync_key = "boyler_ms_v39"
        prev_upper_key = "sicak_su_emniyet_hacmi_v39"
        prev_lower_key = "boyler_ms_onceki_v39"

        if boyler_sync_key not in st.session_state:
            st.session_state[boyler_sync_key] = float(secilen_boyler_hacmi_hesaplanan)
            st.session_state[prev_lower_key] = float(secilen_boyler_hacmi_hesaplanan)
            st.session_state[prev_upper_key] = float(secilen_boyler_hacmi_hesaplanan)

        mevcut_alt_deger = float(st.session_state.get(boyler_sync_key, secilen_boyler_hacmi_hesaplanan))
        onceki_ust_deger = float(st.session_state.get(prev_upper_key, secilen_boyler_hacmi_hesaplanan))
        onceki_alt_deger = float(st.session_state.get(prev_lower_key, mevcut_alt_deger))
        ust_degisti = abs(float(secilen_boyler_hacmi_hesaplanan) - onceki_ust_deger) > 1e-9
        alt_degisti = abs(mevcut_alt_deger - onceki_alt_deger) > 1e-9

        if alt_degisti and not ust_degisti:
            # Alt bölüm kullanıcı tarafından değiştirildi: üstteki emniyetli hacmi güncelle.
            secilen_boyler_hacmi = int(round(mevcut_alt_deger))
        elif ust_degisti and not alt_degisti:
            # Üstteki hesap değişti: alt bölümü yeni emniyetli hacme getir.
            st.session_state[boyler_sync_key] = float(secilen_boyler_hacmi_hesaplanan)
            secilen_boyler_hacmi = int(secilen_boyler_hacmi_hesaplanan)
        elif alt_degisti and ust_degisti:
            # Aynı turda iki değişiklik varsa kullanıcının alt bölümdeki manuel değeri önceliklidir.
            secilen_boyler_hacmi = int(round(mevcut_alt_deger))
        else:
            secilen_boyler_hacmi = int(round(mevcut_alt_deger))

        st.session_state[prev_upper_key] = float(secilen_boyler_hacmi_hesaplanan)
        st.session_state[prev_lower_key] = float(st.session_state.get(boyler_sync_key, secilen_boyler_hacmi))

        st.markdown(
            f"Kullanma Eş Zaman Faktörü = **{kullanma_es_faktoru:.2f}**"
        )
        st.markdown(
            f"Depolama Faktörü = **{depolama_faktoru:.2f}**"
        )
        st.markdown(
            f"V = {depolama_faktoru:.2f} × {kullanma_es_faktoru:.2f} × {int(round(toplam_tuketim))} = **{int(round(hesaplanan_boyler_hacmi))} L**"
        )
        st.markdown(
            f"V = **{secilen_boyler_hacmi} L** (Emniyetle)"
        )

        # --- BOYLER ISITICI KAPASİTESİ ---
        # Bu başlık PROGRAM EKRANINDA sabit kalır.
        # İstenen özel başlıklar yalnızca oluşturulan RAPORDA kullanılır.
        st.markdown("**BOYLER ISITICI KAPASİTESİ**")
        col_q1, col_q2 = st.columns(2)
        with col_q1:
            boyler_ms = st.number_input(
                "Sekonder su debisi (lt/h)",
                min_value=0.0, step=50.0,
                format="%.0f", key=boyler_sync_key
            )
            # Kullanıcının alt değeri üstteki hacimle senkronize edilsin.
            secilen_boyler_hacmi = int(round(boyler_ms))
            boyler_ts_giris = st.number_input(
                "Sekonder giriş sıcaklığı (°C)",
                min_value=0.0, max_value=100.0, step=1.0, value=10.0,
                format="%.0f", key="boyler_ts_giris_v38"
            )
            boyler_ts_cikis = st.number_input(
                "Sekonder çıkış sıcaklığı (°C)",
                min_value=0.0, max_value=100.0, step=1.0, value=60.0,
                format="%.0f", key="boyler_ts_cikis_v38"
            )
        with col_q2:
            boyler_mp = st.number_input(
                "Primer su debisi (lt/h)",
                min_value=0.0, step=50.0, value=0.0,
                format="%.0f", key="boyler_mp_v38"
            )

            # 4. TESİSTE KULLANILACAK ISI İLETİM AKIŞKANLARI bölümünde
            # seçilen Boyler Isıtma Rejimi (ör. 80/60 veya 70/50),
            # burada primer giriş/çıkış sıcaklıklarına otomatik aktarılır.
            # İlk açılışta Boyler Isıtma Rejimi 80/60 olduğundan bu iki alan da
            # doğrudan 80/60 olarak başlatılır. Kullanıcı primer değerleri elle
            # değiştirebilir; ancak üstteki Boyler Isıtma Rejimi değişirse bu
            # alanlar yeni rejime otomatik olarak senkronlanır.
            try:
                rej_metni = str(rej_boyler or "80/60").strip()
                parcalar = rej_metni.replace(",", ".").split("/", 1)
                if len(parcalar) != 2:
                    raise ValueError
                boyler_tp_giris_kaynak = float(parcalar[0].strip())
                boyler_tp_cikis_kaynak = float(parcalar[1].strip())
            except (ValueError, TypeError, AttributeError):
                boyler_tp_giris_kaynak, boyler_tp_cikis_kaynak = 80.0, 60.0

            # v52: Boyler rejimi ile primer giriş/çıkış sıcaklıklarını doğrudan
            # ve güvenli şekilde senkronize et. Eski v38/v42 session_state
            # anahtarlarını kullanmıyoruz; böylece daha önce kalmış 0/0 değerleri
            # yeni rejimin üzerine yazamaz.
            boyler_rejim_key = "boyler_primer_rejim_v52"
            boyler_tp_giris_key = "boyler_tp_giris_v52"
            boyler_tp_cikis_key = "boyler_tp_cikis_v52"
            secili_rejim = str(rej_boyler or "80/60")
            onceki_rejim = st.session_state.get(boyler_rejim_key)

            # İlk açılışta veya rejim değiştiğinde primer sıcaklıklarını seçilen
            # rejime birebir aktar. İlk açılışın varsayılanı 80/60'tır.
            if onceki_rejim is None or onceki_rejim != secili_rejim:
                st.session_state[boyler_tp_giris_key] = float(boyler_tp_giris_kaynak)
                st.session_state[boyler_tp_cikis_key] = float(boyler_tp_cikis_kaynak)
                st.session_state[boyler_rejim_key] = secili_rejim
            else:
                # Eski/boş session değerleri kalmışsa güvenli varsayılanı geri yükle.
                if boyler_tp_giris_key not in st.session_state:
                    st.session_state[boyler_tp_giris_key] = float(boyler_tp_giris_kaynak)
                if boyler_tp_cikis_key not in st.session_state:
                    st.session_state[boyler_tp_cikis_key] = float(boyler_tp_cikis_kaynak)

            boyler_tp_giris = st.number_input(
                "Primer giriş sıcaklığı (°C)",
                min_value=0.0, max_value=150.0, step=1.0,
                format="%.0f", key=boyler_tp_giris_key
            )
            boyler_tp_cikis = st.number_input(
                "Primer çıkış sıcaklığı (°C)",
                min_value=0.0, max_value=150.0, step=1.0,
                format="%.0f", key=boyler_tp_cikis_key
            )

        boyler_c = 1.0
        boyler_delta_ts = boyler_ts_cikis - boyler_ts_giris
        boyler_delta_tp = boyler_tp_giris - boyler_tp_cikis
        boyler_q_kcal_h = boyler_ms * boyler_c * boyler_delta_ts
        # 6.3.4 re-sirkülasyon hesabında seçilmiş boylerin gerçek ısı yükünü kullan.
        st.session_state["boyler_q_kcal_h_v97"] = float(boyler_q_kcal_h)

        # kCal/h -> kW dönüşümü: önce gerçek/küsüratlı değer gösterilir.
        # Ardından en yakın tam sayıya yuvarlanır ve kullanıcıya nihai Q BOYLER
        # değerini elle değiştirebileceği bir alan verilir.
        boyler_q_kw_hesaplanan = boyler_q_kcal_h * 0.001163
        boyler_q_kw_yuvarlanmis = int(math.floor(boyler_q_kw_hesaplanan + 0.5)) if boyler_q_kw_hesaplanan >= 0 else int(math.ceil(boyler_q_kw_hesaplanan - 0.5))

        boyler_q_kw_key = "boyler_q_kw_final_v63"
        boyler_q_kw_prev_key = "boyler_q_kw_hesaplanan_onceki_v63"
        boyler_q_kw_manual_key = f"{boyler_q_kw_key}__manual"
        # Otomatik Q değerini hangi girişlerin ürettiğini takip et. Böylece
        # önceki Streamlit oturumundan kalan (ör. 3 kW) değer yeni hesaba taşınmaz.
        # Kullanıcı aynı hesap üzerinde elle değiştirirse o değer korunur.
        boyler_q_kw_imza_key = "boyler_q_kw_hesap_imza_v68"
        boyler_q_kw_hesap_imzasi = (
            round(float(boyler_ms), 6),
            round(float(boyler_ts_giris), 6),
            round(float(boyler_ts_cikis), 6),
        )
        onceki_imza = st.session_state.get(boyler_q_kw_imza_key)
        hesap_girdileri_degisti = onceki_imza != boyler_q_kw_hesap_imzasi

        if boyler_q_kw_key not in st.session_state or hesap_girdileri_degisti:
            st.session_state[boyler_q_kw_key] = boyler_q_kw_yuvarlanmis
            st.session_state[boyler_q_kw_manual_key] = False
            st.session_state[boyler_q_kw_prev_key] = float(boyler_q_kw_hesaplanan)
            st.session_state[boyler_q_kw_imza_key] = boyler_q_kw_hesap_imzasi
        else:
            st.session_state[boyler_q_kw_prev_key] = float(boyler_q_kw_hesaplanan)

        def _boyler_q_kw_manuel_degisti():
            st.session_state[boyler_q_kw_manual_key] = True

        st.markdown(
            f"Q = {boyler_ms:.0f} lt/h × 1 kCal/kg.°C × ({boyler_ts_cikis:.0f}-{boyler_ts_giris:.0f}) °C"
        )
        st.markdown(
            f"Q = {boyler_q_kcal_h:.0f} kcal/h → **{boyler_q_kw_hesaplanan:.2f} kW** (hesaplanan küsüratlı değer)"
        )
        boyler_q_kw = st.number_input(
            "Nihai Q BOYLER (yuvarlanmış kW) — elle değiştirebilirsiniz",
            min_value=0,
            step=1,
            format="%d",
            key=boyler_q_kw_key,
            on_change=_boyler_q_kw_manuel_degisti,
            help=(
                f"Hesaplanan değer: {boyler_q_kw_hesaplanan:.2f} kW. "
                f"Otomatik yuvarlanan değer: {boyler_q_kw_yuvarlanmis} kW. "
                "Bu alanda değiştirdiğiniz nihai tam sayı hesap raporuna aktarılır."
            ),
        )
        st.caption(
            f"Otomatik yuvarlanan değer: {boyler_q_kw_yuvarlanmis} kW  |  "
            f"Rapor ve nihai seçim değeri: {int(boyler_q_kw)} kW"
        )
        boyler_mp_gerekli = (boyler_q_kcal_h / (boyler_c * boyler_delta_tp)) if boyler_delta_tp > 0 else 0.0

        faktor_satirlari = [{
            "Parametre": "Kullanma eş zaman faktörü",
            "Değer": f"{kullanma_es_faktoru:.2f}",
        }, {
            "Parametre": "Depolama faktörü",
            "Değer": f"{depolama_faktoru:.2f}",
        }]
        if es_zaman_faktoru is not None:
            faktor_satirlari.append({
                "Parametre": f"Excel eş zaman faktörü ({konut_sayisi} konut)",
                "Değer": f"{es_zaman_faktoru:.2f}",
            })
        st.dataframe(pd.DataFrame(faktor_satirlari), hide_index=True, use_container_width=True)

    
        # -----------------------------------------------------------------------
        # BOYLER POZ VERİTABANI - 25.175.1601 ... 25.175.1613
        # Çevre, Şehircilik ve İklim Değişikliği Bakanlığı 2026 mekanik tesisat
        # birim fiyat tariflerindeki Tek Bakır Serpantinli Dik Boyler değerleri.
        # Seçim, her boyler için gerekli hacim ve debiyi birlikte kontrol eder.
        # -----------------------------------------------------------------------
        TEK_SERPANTINLI_BOYLER_POZLARI = [
            {"poz": "25.175.1601", "hacim": 160,  "debi_80_60": 221},
            {"poz": "25.175.1602", "hacim": 200,  "debi_80_60": 272},
            {"poz": "25.175.1603", "hacim": 300,  "debi_80_60": 289},
            {"poz": "25.175.1604", "hacim": 350,  "debi_80_60": 336},
            {"poz": "25.175.1605", "hacim": 500,  "debi_80_60": 476},
            {"poz": "25.175.1606", "hacim": 600,  "debi_80_60": 521},
            {"poz": "25.175.1607", "hacim": 800,  "debi_80_60": 612},
            {"poz": "25.175.1608", "hacim": 1000, "debi_80_60": 663},
            {"poz": "25.175.1609", "hacim": 1250, "debi_80_60": 765},
            {"poz": "25.175.1610", "hacim": 1500, "debi_80_60": 867},
            {"poz": "25.175.1611", "hacim": 2000, "debi_80_60": 1088},
            {"poz": "25.175.1612", "hacim": 2500, "debi_80_60": 1309},
            {"poz": "25.175.1613", "hacim": 3000, "debi_80_60": 1479},
        ]

        def _boyler_poz_sec(hacim_toplam_litre, debi_toplam_lph, adet):
            """Her boyler için gerekli hacim ve debiyi karşılayan ilk poz."""
            adet = max(1, int(adet))
            hacim_birim = float(hacim_toplam_litre) / adet
            debi_birim = float(debi_toplam_lph) / adet

            for poz in TEK_SERPANTINLI_BOYLER_POZLARI:
                if (
                    poz["hacim"] >= hacim_birim
                    and poz["debi_80_60"] >= debi_birim
                ):
                    return poz, hacim_birim, debi_birim, True

            # Tablo sınırı aşılırsa son poz gösterilir ve uyarı verilir.
            return (
                TEK_SERPANTINLI_BOYLER_POZLARI[-1],
                hacim_birim,
                debi_birim,
                False,
            )

    # -----------------------------------------------------------------------
        # ÇİFT SERPANTİNLİ BOYLER POZ VERİTABANI - 25.175.1701 ... 25.175.1714
        # Çift Bakır Serpantinli Dik Boyler. Alt serpantin = kazan, üst serpantin = güneş.
        # -----------------------------------------------------------------------
        CIFT_SERPANTINLI_BOYLER_POZLARI = [
            {"poz": "25.175.1701", "hacim": 160,  "alt_debi": 221,  "ust_debi": 119, "alt_debi_70_50": 73,  "ust_debi_70_50": 36},
            {"poz": "25.175.1702", "hacim": 200,  "alt_debi": 272,  "ust_debi": 150, "alt_debi_70_50": 87,  "ust_debi_70_50": 48},
            {"poz": "25.175.1703", "hacim": 300,  "alt_debi": 289,  "ust_debi": 180, "alt_debi_70_50": 102, "ust_debi_70_50": 58},
            {"poz": "25.175.1704", "hacim": 350,  "alt_debi": 336,  "ust_debi": 190, "alt_debi_70_50": 123, "ust_debi_70_50": 62},
            {"poz": "25.175.1705", "hacim": 500,  "alt_debi": 476,  "ust_debi": 221, "alt_debi_70_50": 187, "ust_debi_70_50": 73},
            {"poz": "25.175.1706", "hacim": 600,  "alt_debi": 521,  "ust_debi": 261, "alt_debi_70_50": 208, "ust_debi_70_50": 93},
            {"poz": "25.175.1707", "hacim": 800,  "alt_debi": 612,  "ust_debi": 340, "alt_debi_70_50": 238, "ust_debi_70_50": 131},
            {"poz": "25.175.1708", "hacim": 1000, "alt_debi": 663, "ust_debi": 439, "alt_debi_70_50": 255, "ust_debi_70_50": 177},
            {"poz": "25.175.1709", "hacim": 1250, "alt_debi": 765, "ust_debi": 466, "alt_debi_70_50": 306, "ust_debi_70_50": 188},
            {"poz": "25.175.1710", "hacim": 1500, "alt_debi": 867, "ust_debi": 493, "alt_debi_70_50": 357, "ust_debi_70_50": 192},
            {"poz": "25.175.1711", "hacim": 2000, "alt_debi": 1088, "ust_debi": 799, "alt_debi_70_50": 442, "ust_debi_70_50": 215},
            {"poz": "25.175.1712", "hacim": 2500, "alt_debi": 1309, "ust_debi": 629, "alt_debi_70_50": 544, "ust_debi_70_50": 238},
            {"poz": "25.175.1713", "hacim": 3000, "alt_debi": 1479, "ust_debi": 697, "alt_debi_70_50": 595, "ust_debi_70_50": 286},
            {"poz": "25.175.1714", "hacim": 3000, "alt_debi": 3330, "ust_debi": 1530, "alt_debi_70_50": 0, "ust_debi_70_50": 0},
        ]


        # -----------------------------------------------------------------------
        # PLAKALI EŞANJÖR VE AKÜMÜLASYON TANKI POZ VERİTABANLARI
        # Poz açıklamalarındaki kapasite ve tek basınç kaybı değeri esas alınır; aynı değer primer ve sekonder devreye uygulanır.
        # -----------------------------------------------------------------------
        AKUMULASYON_TANKI_POZLARI = [
            {"poz": "25.175.2501", "hacim": 100},
            {"poz": "25.175.2502", "hacim": 150},
            {"poz": "25.175.2503", "hacim": 200},
            {"poz": "25.175.2504", "hacim": 300},
            {"poz": "25.175.2505", "hacim": 350},
            {"poz": "25.175.2506", "hacim": 500},
            {"poz": "25.175.2507", "hacim": 600},
            {"poz": "25.175.2508", "hacim": 800},
            {"poz": "25.175.2509", "hacim": 1000},
            {"poz": "25.175.2510", "hacim": 1250},
            {"poz": "25.175.2511", "hacim": 1500},
            {"poz": "25.175.2512", "hacim": 2000},
            {"poz": "25.175.2513", "hacim": 2500},
            # Poz 25.175.2514 için kullanılacak emniyetli kapasite: V = 3000 L
            {"poz": "25.175.2514", "hacim": 3000},
        ]

        AKUMULASYON_TANKI_TIP = (
            "Kendinden poliüretan izoleli, içi epoxy kaplı, "
            "katodik koruma donanımlı."
        )

        def _akumulasyon_tanki_sec(gerekli_hacim_litre, adet=1):
            """Kullanıcının belirlediği tank adedine göre uygun tekil tank hacmini seçer.

            Mantık: gerekli toplam akümülasyon hacmi / tank adedi = tank başına
            gerekli hacim. Bu değeri karşılayan en küçük poz seçilir. Böylece
            örneğin 2.830 L ve 19 adet için 149 L/tank gerekir ve 150 L'lik
            25.175.2502 poz seçilir; toplam 2.850 L olur.
            """
            gerekli = max(0.0, float(gerekli_hacim_litre))
            adet = max(1, int(adet))

            if gerekli <= 0:
                poz = AKUMULASYON_TANKI_POZLARI[0]
                return {
                    "poz": poz["poz"],
                    "hacim": int(poz["hacim"]),
                    "adet": adet,
                    "toplam_hacim": int(adet * poz["hacim"]),
                    "gerekli_hacim": gerekli,
                    "tip": AKUMULASYON_TANKI_TIP,
                }

            # Kullanıcının girdiği adet, toplam gerekli hacmi böler.
            # Ardından her tank için gerekli hacmi karşılayan EN KÜÇÜK poz seçilir.
            # Örnek: 14.360 L ihtiyaç ve 8 adet -> 1.795 L/tank -> 2.000 L
            # (25.175.2512) seçilir ve toplam kapasite 16.000 L olur.
            birim_gerekli = gerekli / float(adet)
            secilen = None
            for poz in AKUMULASYON_TANKI_POZLARI:
                if float(poz["hacim"]) >= birim_gerekli:
                    secilen = poz
                    break

            # İstenen adet çok az ise son poz bile yetersiz kalabilir. Bu durumda
            # son poz seçilir ve program açıkça yetersizlik durumunu bildirir.
            if secilen is None:
                secilen = AKUMULASYON_TANKI_POZLARI[-1]

            toplam = int(adet * float(secilen["hacim"]))
            yeterli = toplam >= gerekli
            return {
                "poz": secilen["poz"],
                "hacim": int(secilen["hacim"]),
                "adet": adet,
                "toplam_hacim": toplam,
                "gerekli_hacim": gerekli,
                "birim_gerekli_hacim": birim_gerekli,
                "yeterli": yeterli,
                "tip": AKUMULASYON_TANKI_TIP,
            }

        PLAKALI_ESANJOR_POZLARI = [
            {"poz": "25.220.2101", "q_kcal_h": 20000, "primer_dp_mss": 0.5},
            {"poz": "25.220.2102", "q_kcal_h": 50000, "primer_dp_mss": 1.0},
            {"poz": "25.220.2103", "q_kcal_h": 75000, "primer_dp_mss": 1.5},
            {"poz": "25.220.2104", "q_kcal_h": 100000, "primer_dp_mss": 2.0},
            {"poz": "25.220.2105", "q_kcal_h": 200000, "primer_dp_mss": 3.0},
            {"poz": "25.220.2106", "q_kcal_h": 300000, "primer_dp_mss": 3.0},
            {"poz": "25.220.2107", "q_kcal_h": 400000, "primer_dp_mss": 3.0},
            {"poz": "25.220.2108", "q_kcal_h": 500000, "primer_dp_mss": 3.0},
            {"poz": "25.220.2109", "q_kcal_h": 600000, "primer_dp_mss": 3.0},
            {"poz": "25.220.2110", "q_kcal_h": 700000, "primer_dp_mss": 4.0},
            {"poz": "25.220.2111", "q_kcal_h": 800000, "primer_dp_mss": 4.0},
            {"poz": "25.220.2112", "q_kcal_h": 900000, "primer_dp_mss": 4.0},
            {"poz": "25.220.2113", "q_kcal_h": 1000000, "primer_dp_mss": 4.0},
        ]

        def _plakali_esanjör_poz_sec(q_kcal_h):
            """Eşanjör başına gerekli ısı yükünü karşılayan en küçük pozu seçer."""
            q = max(0.0, float(q_kcal_h))
            for poz in PLAKALI_ESANJOR_POZLARI:
                if float(poz["q_kcal_h"]) >= q:
                    return poz
            return PLAKALI_ESANJOR_POZLARI[-1]

        def _cift_boyler_poz_sec(hacim_toplam_litre, debi_toplam_lph, adet):
            """Çift serpantinli boyler için otomatik poz seçimi."""
            adet = max(1, int(adet))
            hacim_birim = float(hacim_toplam_litre) / adet
            debi_birim = float(debi_toplam_lph) / adet
            for poz in CIFT_SERPANTINLI_BOYLER_POZLARI:
                toplam_serpantin_debisi = float(poz["alt_debi"]) + float(poz["ust_debi"])
                if poz["hacim"] >= hacim_birim and toplam_serpantin_debisi >= debi_birim:
                    return poz, hacim_birim, debi_birim, True
            return CIFT_SERPANTINLI_BOYLER_POZLARI[-1], hacim_birim, debi_birim, False

    # -----------------------------------------------------------------------
        # BOYLER / EŞANJÖR SEÇİMİ
        # -----------------------------------------------------------------------
        # Üstte yapılan sıcak su ihtiyacı hesabının sonuçları bu üç seçim
        # sekmesinde ortak referans olarak kullanılacaktır. Ekipman/poz seçimleri
        # daha sonra her sekmenin kendi veri tabanına bağlanabilir.
        st.markdown("### • BOYLER / EŞANJÖR SEÇİMİ")

        # Rapor için tek bir boyler/eşanjör sistemi seçilir.
        # Varsayılan: Tek Serpantinli Boyler.
        boyler_secili_tip_key = "boyler_secili_tip_v57"
        st.session_state.setdefault(
            boyler_secili_tip_key,
            "TEK SERPANTİNLİ BOYLER",
        )

        # st.tabs() ile aktif sekmeyi Python tarafından seçmek mümkün değildir.
        # Bu nedenle sekme görünümünü yatay radio ile oluşturuyoruz. Üst bölümdeki
        # seçim her rerun'da bu değeri güncellediği için alt sekme otomatik seçilir.
        _boyler_alt_sekme_labels = [
            "TEK SERPANTİNLİ BOYLER",
            "ÇİFT SERPANTİNLİ BOYLER",
            "PLAKALI EŞANJÖR",
        ]
        _boyler_alt_sekme_index = int(
            st.session_state.get("boyler_alt_sekme_v71", 0)
        )
        _boyler_alt_sekme_index = max(0, min(2, _boyler_alt_sekme_index))
        _boyler_alt_sekme = st.radio(
            "Boyler / Eşanjör tipi",
            _boyler_alt_sekme_labels,
            index=_boyler_alt_sekme_index,
            horizontal=True,
            label_visibility="collapsed",
            key="boyler_alt_sekme_v71",
        )
        _boyler_alt_sekme_index = _boyler_alt_sekme_labels.index(_boyler_alt_sekme)

        # Boyler / eşanjör Cihaz Poz No bilgisinin raporda gösterilip gösterilmeyeceği
        # kullanıcı tercihine bırakılır. Varsayılan True ile mevcut davranış korunur.
        boyler_poz_rapora_eklensin_key = "boyler_poz_rapora_eklensin_v72"
        boyler_poz_rapora_eklensin = st.checkbox(
            "Cihaz Poz Numarasını Hesap Raporunda Göster",
            value=bool(st.session_state.get(boyler_poz_rapora_eklensin_key, True)),
            key=boyler_poz_rapora_eklensin_key,
            help="İşaretli ise seçilen tek/çift serpantinli boyler veya plakalı eşanjörün Cihaz Poz No bilgisi rapora eklenir. İşaretli değilse raporda gösterilmez.",
        )

        # Üst seçim ile alt seçim arasında tam senkronizasyon.
        _boyler_secili_tip_from_alt = _boyler_alt_sekme_labels[_boyler_alt_sekme_index]
        if st.session_state.get("boyler_secili_tip_v57") != _boyler_secili_tip_from_alt:
            st.session_state["boyler_secili_tip_v57"] = _boyler_secili_tip_from_alt

        boyler_secim_ortak_bilgiler = {
            "Gerekli boyler hacmi": f"{secilen_boyler_hacmi} L",
            "Toplam günlük sıcak su tüketimi": f"{int(round(toplam_tuketim))} L/gün",
            "Hesaplanan ısı yükü": f"{int(boyler_q_kw)} kW",
            "Sekonder giriş / çıkış": f"{boyler_ts_giris:.0f} / {boyler_ts_cikis:.0f} °C",
            "Primer giriş / çıkış": f"{boyler_tp_giris:.0f} / {boyler_tp_cikis:.0f} °C",
        }

        for _tab_index in range(3):
            # Yalnızca aktif sekmenin içeriğini çiz. Böylece üst seçimle belirlenen
            # sekme gerçekten açılmış/görüntülenmiş olur.
            if _tab_index != _boyler_alt_sekme_index:
                continue
            with st.container():
                if _tab_index == 0:
                    _tab_prefix = "tek_serpantin"
                    _tab_title = "Tek Serpantinli Boyler Seçimi"
                    _tip_adi = "TEK SERPANTİNLİ BOYLER"
                elif _tab_index == 1:
                    _tab_prefix = "cift_serpantin"
                    _tab_title = "Çift Serpantinli Boyler Seçimi"
                    _tip_adi = "ÇİFT SERPANTİNLİ BOYLER"
                else:
                    _tab_prefix = "plakali_esanjör"
                    _tab_title = "Plakalı Eşanjör Seçimi"
                    _tip_adi = "PLAKALI EŞANJÖR"

                st.markdown(f"#### {_tab_title}")

                _rapor_sec_key = f"boyler_raporda_kullan_{_tab_prefix}_v58"
                if st.button(
                    "✓ BU SİSTEMİ RAPORDA KULLAN",
                    key=_rapor_sec_key,
                    use_container_width=True,
                ):
                    st.session_state[boyler_secili_tip_key] = _tip_adi
                    st.rerun()

                if st.session_state.get(boyler_secili_tip_key) == _tip_adi:
                    st.success("Bu sistem raporda kullanılacak şekilde seçildi.")
                else:
                    st.caption(
                        f"Raporda şu anda kullanılan sistem: "
                        f"{st.session_state.get(boyler_secili_tip_key, 'TEK SERPANTİNLİ BOYLER')}"
                    )
                st.caption(
                    "Seçim kriterleri, yukarıdaki sıcak su ihtiyacı ve ısı yükü "
                    "hesabından otomatik olarak alınacaktır."
                )

                # Ortak hesap sonuçlarını her sekmede göster.
                _bilgi_cols = st.columns(3)
                with _bilgi_cols[0]:
                    st.metric("Gerekli hacim", f"{secilen_boyler_hacmi} L")
                with _bilgi_cols[1]:
                    st.metric("Isı yükü", f"{int(boyler_q_kw)} kW")
                with _bilgi_cols[2]:
                    st.metric("Günlük tüketim", f"{int(round(toplam_tuketim))} L/gün")

                with st.expander("Hesap kriterlerini göster", expanded=False):
                    st.table(
                        pd.DataFrame(
                            list(boyler_secim_ortak_bilgiler.items()),
                            columns=["Hesap Kriteri", "Değer"],
                        )
                    )

                _adet_key = f"boyler_adet_{_tab_prefix}_v59"
                _adet_label = (
                    "Plakalı eşanjör adedi"
                    if _tab_prefix == "plakali_esanjör"
                    else "Boyler adedi"
                )
                _boyler_adet = st.number_input(
                    _adet_label,
                    min_value=1,
                    value=3 if _tab_prefix != "plakali_esanjör" else 2,
                    step=1,
                    key=_adet_key,
                )

                # Tek serpantinli boyler için poz seçimi:
                # toplam gerekli hacim ve toplam hesaplanan debi, girilen adet sayısına
                # bölünür; her iki şartı da sağlayan en küçük poz seçilir.
                if _tab_prefix == "tek_serpantin":
                    _secim_pozu, _gerekli_hacim_birim, _gerekli_debi_birim, _poz_yeterli = _boyler_poz_sec(
                        secilen_boyler_hacmi, boyler_ms, _boyler_adet
                    )
                    _boyler_secim_sonucu_key = "boyler_secim_sonucu_v59"
                    st.session_state[_boyler_secim_sonucu_key] = {
                        "tip": _tip_adi, "adet": int(_boyler_adet),
                        "poz": _secim_pozu["poz"], "hacim": int(_secim_pozu["hacim"]),
                        "debi": int(_secim_pozu["debi_80_60"]),
                        "gerekli_hacim_birim": _gerekli_hacim_birim,
                        "gerekli_debi_birim": _gerekli_debi_birim, "poz_yeterli": _poz_yeterli,
                    }
                    st.markdown("**Otomatik Boyler Seçimi**")
                    _c1, _c2, _c3, _c4 = st.columns(4)
                    with _c1: st.metric("Adet", int(_boyler_adet))
                    with _c2: st.metric("Boyler hacmi", f"{_secim_pozu['hacim']} L")
                    with _c3: st.metric("Boyler debisi", f"{_secim_pozu['debi_80_60']} L/h")
                    with _c4: st.metric("Cihaz Poz No", _secim_pozu["poz"])
                    st.success(f"Seçilen Cihaz Poz No: **{_secim_pozu['poz']}**  |  Boyler: **{_secim_pozu['hacim']} L**  |  Debi: **{_secim_pozu['debi_80_60']} L/h**")
                    st.caption(f"Her boyler için gerekli: {_gerekli_hacim_birim:.0f} L / {_gerekli_debi_birim:.0f} L/h")
                    if not _poz_yeterli: st.warning("Hesaplanan gereksinim 25.175.1613 pozunun kapasite/debi sınırını aşıyor.")

                elif _tab_prefix == "cift_serpantin":
                    # Çift serpantin: alt serpantin kazan, üst serpantin güneştir.
                    # Güneş rejimi Bölüm 4'teki 11. maddeden otomatik aktarılır.
                    _gunes_rejim = str(
                        st.session_state.get("rej_gunes_ist_v68", "60/40")
                    )
                    st.session_state["boyler_gunes_rejimi_v66"] = _gunes_rejim
                    st.info(
                        f"Güneş serpantini rejimi: **{_gunes_rejim} °C sıcak su** "
                        "(Bölüm 4'ten otomatik aktarılır.)"
                    )
                    _cift_poz, _cift_gerekli_hacim, _cift_gerekli_debi, _cift_yeterli = _cift_boyler_poz_sec(
                        secilen_boyler_hacmi, boyler_ms, _boyler_adet
                    )
                    _cift_sonuc = {
                        "tip": _tip_adi, "adet": int(_boyler_adet),
                        "poz": _cift_poz["poz"], "hacim": int(_cift_poz["hacim"]),
                        "alt_debi": int(_cift_poz["alt_debi"]), "ust_debi": int(_cift_poz["ust_debi"]),
                        "gerekli_hacim_birim": _cift_gerekli_hacim,
                        "gerekli_debi_birim": _cift_gerekli_debi, "poz_yeterli": _cift_yeterli,
                        "gunes_rejimi": _gunes_rejim,
                    }
                    st.session_state["boyler_cift_secim_sonucu_v66"] = _cift_sonuc
                    st.markdown("**Otomatik Çift Serpantinli Boyler Seçimi**")
                    _c1, _c2, _c3, _c4, _c5 = st.columns(5)
                    with _c1: st.metric("Adet", int(_boyler_adet))
                    with _c2: st.metric("Boyler hacmi", f"{_cift_poz['hacim']} L")
                    with _c3: st.metric("Alt serpantin", f"{_cift_poz['alt_debi']} L/h")
                    with _c4: st.metric("Üst serpantin", f"{_cift_poz['ust_debi']} L/h")
                    with _c5: st.metric("Cihaz Poz No", _cift_poz["poz"])
                    st.success(
                        f"Seçilen Cihaz Poz No: **{_cift_poz['poz']}** | "
                        f"Boyler: **{_cift_poz['hacim']} L** | "
                        f"Alt: **{_cift_poz['alt_debi']} L/h** | Üst: **{_cift_poz['ust_debi']} L/h**"
                    )
                    st.caption(
                        f"Her boyler için gerekli hacim: {_cift_gerekli_hacim:.0f} L | "
                        f"Toplam gerekli serpantin debisi: {_cift_gerekli_debi:.0f} L/h"
                    )
                    if not _cift_yeterli:
                        st.warning("Hesaplanan gereksinim 25.175.1714 pozunun sınırlarını aşıyor; özel ürün seçimi gerekir.")

                elif _tab_prefix == "plakali_esanjör":
                    # Plakalı eşanjör sistemi iki ekipmandan oluşur:
                    # 1) Akümülasyon tankı, 2) Plakalı eşanjör.
                    # AKÜMÜLASYON TANKI GEREKLİ HACMİ
                    # Kaynak doğrudan yukarıdaki "Ortalama Ani Sıcak Su İhtiyacı"
                    # hesabındaki emniyetli V değeridir (ör. V = 3500 L).
                    # Akümülasyon tankı seçim katsayısı bu değere ayrıca uygulanır:
                    # Gerekli akümülasyon hacmi = V (Emniyetle) × katsayı.
                    # Böylece katsayı 1.00 ise 3500 L, katsayı 0.50 ise 1750 L,
                    # katsayı 1.50 ise 5250 L esas alınır.
                    _akum_katsayi_key = "plakali_akumulasyon_katsayisi_v86"
                    _akum_katsayisi = st.number_input(
                        "Akümülasyon tankı seçim katsayısı",
                        min_value=0.0,
                        max_value=10.0,
                        step=0.10,
                        format="%.2f",
                        value=float(st.session_state.get(_akum_katsayi_key, 1.0)),
                        key=_akum_katsayi_key,
                        help=(
                            "Akümülasyon tankı gerekli hacmi = Ortalama Ani Sıcak Su İhtiyacı "
                            "(V, Emniyetle) × bu katsayı."
                        ),
                    )
                    _akum_kaynak_v = float(secilen_boyler_hacmi_hesaplanan)
                    _akum_gerekli = _akum_kaynak_v * float(_akum_katsayisi)

                    st.caption(
                        f"Akümülasyon hesabı: V = {_akum_kaynak_v:.0f} L (Emniyetle) × "
                        f"{_akum_katsayisi:.2f} = **{_akum_gerekli:.0f} L**"
                    )

                    # Tank adedi kullanıcı tarafından belirlenir. İlk açılışta,
                    # 3000 L poz kapasitesi üzerinden toplam ihtiyacı
                    # karşılayacak minimum adet önerilir.
                    _akum_otomatik_oneri = (
                        min(100, max(1, int(math.ceil(_akum_gerekli / 3000.0))))
                        if _akum_gerekli > 0 else 1
                    )
                    _akum_adet_key = "plakali_akumulasyon_adedi_v85"
                    if _akum_adet_key not in st.session_state:
                        st.session_state[_akum_adet_key] = int(_akum_otomatik_oneri)

                    _akum_adet = st.number_input(
                        "Akümülasyon tankı adedi",
                        min_value=1,
                        max_value=100,
                        step=1,
                        key=_akum_adet_key,
                        help=(
                            "Adedi değiştirince gerekli hacim tank adedine bölünür ve "
                            "uygun tank hacmi/poz otomatik seçilir. Tank hacmini aşağıdaki "
                            "seçim kutusundan elle değiştirebilirsiniz."
                        ),
                    )

                    # Önce otomatik tank seçimini yap. Mantık: gerekli toplam hacim / adet
                    # ve bu değeri karşılayan bir üst poz kapasitesi.
                    _akum_otomatik_secim = _akumulasyon_tanki_sec(
                        _akum_gerekli, int(_akum_adet)
                    )

                    # Tank hacmi/poz kullanıcı tarafından da elle değiştirilebilir.
                    # Adet veya gerekli hacim değiştiğinde seçim otomatik öneriye döner;
                    # kullanıcı seçim kutusunu değiştirdiğinde manuel tercih korunur.
                    _akum_poz_secenekleri = [
                        f"{poz['hacim']} L | {poz['poz']}"
                        for poz in AKUMULASYON_TANKI_POZLARI
                    ]
                    _akum_poz_degerleri = [poz["poz"] for poz in AKUMULASYON_TANKI_POZLARI]
                    _akum_secim_signature = (
                        round(float(_akum_kaynak_v), 6),
                        round(float(_akum_katsayisi), 6),
                        round(float(_akum_gerekli), 6),
                        int(_akum_adet),
                    )
                    _akum_prev_signature_key = "plakali_akumulasyon_secim_signature_v85"
                    _akum_poz_key = "plakali_akumulasyon_poz_secimi_v85"
                    _akum_otomatik_label = (
                        f"{_akum_otomatik_secim['hacim']} L | {_akum_otomatik_secim['poz']}"
                    )
                    if st.session_state.get(_akum_prev_signature_key) != _akum_secim_signature:
                        st.session_state[_akum_poz_key] = _akum_otomatik_label
                        st.session_state[_akum_prev_signature_key] = _akum_secim_signature
                    elif st.session_state.get(_akum_poz_key) not in _akum_poz_secenekleri:
                        st.session_state[_akum_poz_key] = _akum_otomatik_label

                    _akum_poz_secimi = st.selectbox(
                        "Tank hacmi / Cihaz Poz No (otomatik veya manuel)",
                        options=_akum_poz_secenekleri,
                        key=_akum_poz_key,
                        help=(
                            "Program adede göre otomatik seçim yapar. İsterseniz bu listeden "
                            "tank hacmini ve Cihaz Poz No'yu elle değiştirebilirsiniz."
                        ),
                    )
                    _akum_secilen_poz_no = _akum_poz_degerleri[
                        _akum_poz_secenekleri.index(_akum_poz_secimi)
                    ]
                    _akum_manuel_poz = next(
                        poz for poz in AKUMULASYON_TANKI_POZLARI
                        if poz["poz"] == _akum_secilen_poz_no
                    )

                    _akum_secim = {
                        "poz": _akum_manuel_poz["poz"],
                        "hacim": int(_akum_manuel_poz["hacim"]),
                        "adet": int(_akum_adet),
                        "toplam_hacim": int(_akum_adet * _akum_manuel_poz["hacim"]),
                        "gerekli_hacim": _akum_gerekli,
                        "birim_gerekli_hacim": _akum_gerekli / float(_akum_adet),
                        "yeterli": int(_akum_adet * _akum_manuel_poz["hacim"]) >= _akum_gerekli,
                        "tip": AKUMULASYON_TANKI_TIP,
                        "otomatik_poz": _akum_otomatik_secim["poz"],
                    }

                    st.markdown("**1. SICAK SU AKÜMÜLASYON TANKI SEÇİMİ**")
                    _a1, _a2, _a3, _a4 = st.columns(4)
                    with _a1: st.metric("Gerekli hacim", f"{_akum_gerekli:.0f} L")
                    with _a2: st.metric("Tank hacmi", f"{_akum_secim['hacim']} L")
                    with _a3: st.metric("Adet", int(_akum_adet))
                    with _a4: st.metric("Cihaz Poz No", _akum_secim["poz"])
                    _akum_sonuc_metni = (
                        f"Akümülasyon tankı: **{int(_akum_adet)} adet × {_akum_secim['hacim']} L** | "
                        f"Toplam: **{_akum_secim['toplam_hacim']} L** | Cihaz Poz No: **{_akum_secim['poz']}**"
                    )
                    if _akum_secim["yeterli"]:
                        st.success(_akum_sonuc_metni)
                    else:
                        st.warning(
                            _akum_sonuc_metni +
                            "\n\n⚠️ Seçilen tankların toplam kapasitesi gerekli hacmi karşılamıyor. "
                            "Tank adedini artırın veya daha büyük tank hacmi seçin."
                        )
                    st.caption(
                        f"Tank başına gerekli hacim: {_akum_secim.get('birim_gerekli_hacim', 0):.0f} L | "
                        f"Seçilen tank toplam kapasitesi: {_akum_secim['toplam_hacim']:.0f} L | "
                        f"Tank tipi: {_akum_secim.get('tip', AKUMULASYON_TANKI_TIP)}"
                    )

                    st.session_state["plakali_akumulasyon_secim_sonucu_v75"] = _akum_secim

                    st.markdown("**2. KULLANMA SICAK SU SİSTEMİ PLAKALI EŞANJÖRÜ**")
                    # Toplam ısıtma yükü, yedek hariç çalışan plakalı eşanjör adedine bölünür.
                    # Her eşanjör için gerekli kapasiteye göre poz seçilir.
                    _plaka_calisma_adet = max(1, int(_boyler_adet))
                    _plaka_yedek_adet = 1
                    _plaka_toplam_adet = _plaka_calisma_adet + _plaka_yedek_adet
                    _plaka_q_birim = float(boyler_q_kcal_h) / float(_plaka_calisma_adet)
                    _plaka_poz = _plakali_esanjör_poz_sec(_plaka_q_birim)
                    _plaka_adet = _plaka_toplam_adet
                    _plaka_sonuc = {
                        "tip": _tip_adi,
                        "adet": _plaka_toplam_adet,
                        "calisma_adet": _plaka_calisma_adet,
                        "yedek_adet": _plaka_yedek_adet,
                        "poz": _plaka_poz["poz"],
                        "q_kcal_h": int(_plaka_poz["q_kcal_h"]),
                        "q_kw": _plaka_poz["q_kcal_h"] * 0.001163,
                        "primer_dp_mss": float(_plaka_poz["primer_dp_mss"]),
                        # Poz açıklamalarında primer kayıp açıkça tanımlıdır.
                        # Sekonder kayıp ayrıca poz açıklamasında verilmediği için
                        # rapordaki proje kabulü 4 mSS olarak tutulur.
                        "sekonder_dp_mss": float(_plaka_poz["primer_dp_mss"]),
                        "q_hesap_kcal_h": float(_plaka_q_birim),
                        "q_toplam_kcal_h": float(boyler_q_kcal_h),
                        "primer_rejim": "80/60 °C sıcak su (Kazan)",
                        "sekonder_rejim": "10/60 °C sıcak su",
                    }
                    st.session_state["plakali_esanjör_secim_sonucu_v75"] = _plaka_sonuc
                    _p1, _p2, _p3, _p4 = st.columns(4)
                    with _p1: st.metric("Eşanjör başına Q", f"{_plaka_q_birim:,.0f} kcal/h".replace(",", "."))
                    with _p2: st.metric("Poz kapasitesi", f"{_plaka_poz['q_kcal_h']:,.0f} kcal/h".replace(",", "."))
                    with _p3: st.metric("Primer ΔP", f"{_plaka_poz['primer_dp_mss']:g} mSS")
                    with _p4: st.metric("Toplam adet", _plaka_toplam_adet)
                    st.success(
                        f"Seçilen Cihaz Poz No: **{_plaka_poz['poz']}** | "
                        f"Kapasite: **{_plaka_poz['q_kcal_h']:,.0f} kcal/h** | "
                        f"Toplam: **{_plaka_toplam_adet} adet** | "
                        f"Yedek: **1 adet** | "
                        f"Primer ΔP: **{_plaka_poz['primer_dp_mss']:g} mSS** | "
                        f"Sekonder ΔP: **4 mSS**".replace(",", ".")
                    )
                    st.info(
                        "Poz açıklamalarındaki kapasite ve primer basınç kaybı otomatik alınır. "
                        "Sekonder basınç kaybı poz açıklamasında ayrı bir değer olarak yer almadığı için "
                        "mevcut proje kabulü olan 4 mSS uygulanır."
                    )

                # Seçim yöntemi / manuel alanı:
                # kullanıcı tarafından verilecek kapasite ve ürün tabloları
                # üzerinden bu alanlara bağlanacaktır.
                _secim_key = f"boyler_secim_{_tab_prefix}_v56"
                _manuel_key = f"boyler_manuel_kapasite_{_tab_prefix}_v56"

                st.markdown("**Seçim**")
                _secim_modu = st.radio(
                    "Seçim yöntemi",
                    ["Otomatik hesaplanan değeri kullan", "Manuel seçim"],
                    horizontal=True,
                    key=_secim_key,
                )

                if _secim_modu == "Otomatik hesaplanan değeri kullan":
                    st.info(
                        f"Önerilen hesap değeri: **{secilen_boyler_hacmi} L** "
                        f"({secilen_boyler_hacmi / 1000:.2f} m³). "
                        "Ürün/poz tablosu eklendiğinde en uygun ürün otomatik seçilecektir."
                    )
                else:
                    st.number_input(
                        "Manuel seçilecek kapasite [L]",
                        min_value=0.0,
                        value=float(secilen_boyler_hacmi),
                        step=50.0,
                        key=_manuel_key,
                    )

                st.caption(
                    "Tek serpantinli boyler poz seçimi 25.175.1601–25.175.1613 "
                    "aralığındaki hacim ve 80/60 °C sıcak su debisi değerlerine göre "
                    "otomatik yapılır. Çift serpantinli boyler ve plakalı eşanjör "
                    "tabloları ayrı poz verileri tanımlandığında aynı mantıkla bağlanacaktır."
                )


        # Rapor için yalnızca kullanılan satırları sakla; kaynakta olmayan Engelli satırı da kullanıcı değer girdiyse aktarılır.
        for _, row in aktif_sonuc_df.iterrows():
            sicak_su_hesap_detaylari.append({
                "kullanim": str(row["Kullanım Yeri"]),
                "kaynak_aralik": str(row["Standart Aralığı [L]"]),
                "birim_degeri": float(row["Birim Tüketim [L]"]),
                "miktar": float(row["Adet"]),
                "toplam_litre": float(row["Toplam [L/gün]"]),
            })

        st.markdown("**Hesaplanan kullanım yerleri**")
        st.dataframe(
            aktif_sonuc_df[["Kullanım Yeri", "Standart Aralığı [L]", "Birim Tüketim [L]", "Adet", "Toplam [L/gün]"]],
            hide_index=True,
            use_container_width=True,
        )
        st.metric("Günlük toplam kullanma sıcak suyu ihtiyacı", f"{sicak_su_gunluk_toplam_litre:,.0f} L/gün".replace(",", "."))

    # ---------------------------------------------------------------------------
    # 6.3.4 KULLANMA SICAK SU TESİSATI RE-SİRKULASYON POMPASI SEÇİMİ
    # ---------------------------------------------------------------------------
    re_sirkulasyon_pompa_sonucu = {}
    if bolum_634_aktif:
        st.markdown('<div id="bolum_634"></div>', unsafe_allow_html=True)
        st.markdown(f"### • {_63_dinamik_baslik("rapor_bolum_634")}")
        st.caption("Q değeri seçilmiş boylerden otomatik alınır. Emniyet oranı ve hesaplanan debi kullanıcı tarafından gerektiğinde değiştirilebilir.")

        _rs_q_boyler = float(st.session_state.get("boyler_q_kcal_h_v97", 0.0))
        _rs_q_boyler_kw = _rs_q_boyler * 0.001163

        # Emniyet oranı: varsayılan %15, kullanıcı tarafından değiştirilebilir.
        re_sirk_emniyet = st.number_input(
            "Emniyet Oranı (%)",
            min_value=0.0, max_value=100.0, value=15.0, step=1.0,
            key="re_sirk_emniyet_v99",
            help="Re-sirkülasyon debisi hesabında QBOYLER × 0,05 × (1 + emniyet oranı) kullanılır. Varsayılan %15'tir."
        )

        # 1. AŞAMA: emniyet katsayısı uygulanmadan temel re-sirkülasyon debisi
        _rs_q_temsiz = (
            _rs_q_boyler * 0.05 / 5000.0
            if _rs_q_boyler > 0 else 0.0
        )
        # 2. AŞAMA: temel debiye kullanıcı tarafından girilen emniyet oranı uygulanır.
        _rs_q_hesap = _rs_q_temsiz * (1.0 + float(re_sirk_emniyet) / 100.0)

        if _rs_q_boyler <= 0:
            st.warning(f"Önce {_63_dinamik_baslik('rapor_bolum_633')} bölümünde boyler seçimi/hesabı yapılmalıdır. Re-sirkülasyon debisi seçilmiş boyler kapasitesinden otomatik alınacaktır.")
            _rs_q_temsiz = 0.0
            _rs_q_hesap = 0.0
        else:
            st.success(
                f"Seçilen boyler ısı yükü: **{_rs_q_boyler:,.0f} kcal/h ≈ {_rs_q_boyler_kw:,.0f} kW**".replace(",", ".")
            )
            st.markdown(
                f"**1. Aşama — Emniyetsiz debi:** V₀ = ({_rs_q_boyler:,.0f} × 0,05) / 5.000 = **{_rs_q_temsiz:.2f} m³/h**".replace(",", ".")
            )
            st.markdown(
                f"**2. Aşama — Emniyetli debi:** V = {_rs_q_temsiz:.2f} × (1 + %{float(re_sirk_emniyet):.0f}) = **{_rs_q_hesap:.2f} m³/h**".replace(",", ".")
            )

        # Hesaplanan emniyetli debiyi aşağıdaki pompa seçim debisi alanına otomatik aktar.
        # Kullanıcı alanı elle değiştirmişse, sonraki rerun'larda manuel değer korunur.
        _rs_prev_auto = st.session_state.get("re_sirk_q_auto_prev_v100")
        _rs_q_key = "re_sirk_q_v99"
        _rs_mevcut_q = st.session_state.get(_rs_q_key)
        if _rs_mevcut_q is None or (_rs_prev_auto is not None and abs(float(_rs_mevcut_q) - float(_rs_prev_auto)) < 1e-9):
            st.session_state[_rs_q_key] = float(_rs_q_hesap)
        st.session_state["re_sirk_q_auto_prev_v100"] = float(_rs_q_hesap)

        st.markdown("#### Pompa Seçim Debisi")
        re_sirk_q = st.number_input(
            "Re-sirkülasyon Debisi V [m³/h]",
            min_value=0.0,
            step=0.01,
            format="%.2f",
            key=_rs_q_key,
            help="Emniyetli debi otomatik olarak bu alana aktarılır. İsterseniz değeri elle değiştirebilirsiniz."
        )

        re_sirk_h = st.number_input(
            "Basma Yüksekliği H [mSS]", min_value=0.5, max_value=20.0, value=5.0, step=0.5,
            key="re_sirk_h_v99"
        )

        _rs_c3, _rs_c4 = st.columns(2)
        with _rs_c3:
            re_sirk_as = st.selectbox("Asıl Pompa Adedi", [1,2,3], index=0, key="re_sirk_as_v99")
        with _rs_c4:
            re_sirk_yedek = st.selectbox("Yedek Pompa Adedi", [1], index=0, key="re_sirk_yedek_v99")

        re_sirk_marka = st.selectbox(
            "Pompa Markası / Seçim Modu",
            ["Otomatik (Wilo + Grundfos)", "Wilo", "Grundfos"],
            index=0, key="re_sirk_marka_v99",
        )

        # ------------------------------------------------------------------
        # POZ NUMARASI SEÇİMİ — otomatik veya manuel
        # ------------------------------------------------------------------
        st.markdown("#### Cihaz Poz Numarası Seçimi")
        _rs_poz_modu = st.radio(
            "Poz seçim yöntemi",
            ["Otomatik Poz Seçimi", "Manuel Poz Seçimi"],
            horizontal=True,
            key="re_sirk_poz_modu_v99",
            label_visibility="collapsed",
        )

        _rs_otomatik_poz, _rs_otomatik_tanim, _rs_otomatik_durum = re_sirk_pompa_pozu_sec(float(re_sirk_q), float(re_sirk_h))
        if _rs_poz_modu == "Manuel Poz Seçimi":
            _rs_poz_opsiyonlari = [x["poz"] for x in RE_SIRK_POMPA_POZ_TABLOSU]
            _rs_mevcut_poz = st.session_state.get("re_sirk_poz_manuel_v99", _rs_otomatik_poz)
            if _rs_mevcut_poz not in _rs_poz_opsiyonlari:
                _rs_mevcut_poz = _rs_poz_opsiyonlari[0]
            re_sirk_poz = st.selectbox(
                "Cihaz Poz No",
                _rs_poz_opsiyonlari,
                index=_rs_poz_opsiyonlari.index(_rs_mevcut_poz),
                key="re_sirk_poz_manuel_v99",
            )
            _rs_poz_kayit = next(x for x in RE_SIRK_POMPA_POZ_TABLOSU if x["poz"] == re_sirk_poz)
            re_sirk_poz_tanim = _rs_poz_kayit["tanim"]
            re_sirk_poz_durum = "UYGUN"
        else:
            re_sirk_poz = _rs_otomatik_poz
            re_sirk_poz_tanim = _rs_otomatik_tanim
            re_sirk_poz_durum = _rs_otomatik_durum

        re_sirk_poz_rapora_aktar = st.checkbox(
            "Cihaz Poz Numarasını Hesap Raporuna Aktar",
            value=True,
            key="re_sirk_poz_rapora_aktar_v99",
        )

        if re_sirk_poz_durum == "UYGUN":
            st.success(f"✅ Seçilen Cihaz Poz No: **{re_sirk_poz}**")
        else:
            st.error("❌ HATA: Re-sirkülasyon pompası çalışma noktası otomatik poz sınırları dışındadır. Manuel poz seçimi ile müdahale edebilirsiniz.")
        st.info(f"📌 **Poz Tanımı:** {re_sirk_poz_tanim}")

        # Üretici/model program ekranında görünür; rapora aktarılmaz.
        re_sirk_model = re_sirk_uretici_sec(float(re_sirk_q), float(re_sirk_h), re_sirk_marka) if re_sirk_q > 0 else None
        re_curve = []
        if re_sirk_model:
            st.success(
                f"Üretici / Model: **{re_sirk_model['marka']} {re_sirk_model['model']}** | "
                f"Çalışma noktası: **{re_sirk_q:.2f} m³/h, {re_sirk_model['h_calisma']:.2f} mSS**"
            )
            re_curve = re_sirk_model["curve"]
            re_graph = pompa_grafigi_png(
                [x[0] for x in re_curve], [x[1] for x in re_curve],
                float(re_sirk_q), float(re_sirk_model["h_calisma"]),
                f"{re_sirk_model['marka']} {re_sirk_model['model']} - Üretici Q-H Eğrisi", anonim=False,
            )
            st.image(re_graph, caption=f"{re_sirk_model['marka']} {re_sirk_model['model']} - Pompa Performans Eğrisi", use_container_width=True)
            st.caption(f"Eğri kaynağı: {re_sirk_model['kaynak']}")
        else:
            st.warning("Seçilen Q/H noktasını karşılayan doğrulanmış Wilo/Grundfos model eğrisi veri setinde bulunamadı.")

        re_toplam_adet = int(re_sirk_as + re_sirk_yedek)
        re_sirkulasyon_pompa_sonucu = {
            "q_boyler_kcal_h": _rs_q_boyler,
            "q_boyler_kw": _rs_q_boyler_kw,
            "emniyet_orani": float(re_sirk_emniyet),
            "q_temsiz_m3h": float(_rs_q_temsiz),
            "q_hesap_m3h": float(_rs_q_hesap),
            "q_m3h": float(re_sirk_q),
            "h_mss": float(re_sirk_h),
            "asil_adet": int(re_sirk_as),
            "yedek_adet": int(re_sirk_yedek),
            "toplam_adet": re_toplam_adet,
            "adet_str": f"{re_toplam_adet} ({re_sirk_as} Asıl, {re_sirk_yedek} Yedek)",
            "poz": re_sirk_poz,
            "poz_tanim": re_sirk_poz_tanim,
            "poz_durumu": re_sirk_poz_durum,
            "poz_rapora_aktar": bool(re_sirk_poz_rapora_aktar),
            "marka": re_sirk_model["marka"] if re_sirk_model else "",
            "model": re_sirk_model["model"] if re_sirk_model else "",
            "guc_kw": float(re_sirk_model["p2_kw"]) if re_sirk_model else 0.20,
            "tip": "Frekans Kontrollü, Düz Boruya Takılabilen Tekli Tip Sirkülâsyon Pompası.",
            "pompa_curve": re_curve,
            "h_calisma": re_sirk_model["h_calisma"] if re_sirk_model else float(re_sirk_h),
            "pompa_kaynak": re_sirk_model["kaynak"] if re_sirk_model else "Doğrulanmış üretici eğrisi bulunamadı",
            "q_secim_m3h": float(re_sirk_q),
        }
        st.session_state["re_sirkulasyon_pompa_sonucu_v99"] = re_sirkulasyon_pompa_sonucu


with _t_yangin:
    st.header("7. YANGIN TESİSATI")
    st.info("Yangın tesisatı modülü bu sekme altında yer alacaktır.")

with _t_isitma:
    st.header("8. ISITMA TESİSATI")
    st.info("Isıtma tesisatı modülü bu sekme altında yer alacaktır.")

with _t_sogutma:
    st.header("9. SOĞUTMA TESİSATI")
    st.info("Soğutma tesisatı modülü bu sekme altında yer alacaktır.")

with _t_havalandirma:
    st.header("10. HAVALANDIRMA TESİSATI")
    st.info("Havalandırma tesisatı modülü bu sekme altında yer alacaktır.")

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

        # --- 6.2.3 YAĞ AYIRICI SEÇİMLERİ ---
        if bolum_623_aktif:
          doc.add_heading("6.2.3 YAĞ AYIRICI SEÇİMLERİ", level=2)
          doc.add_paragraph(
              "Bu bölüm mutfak / yemekhane kaynaklı atık sularda kullanılacak yağ ayırıcıya ilişkin "
              "genel seçim ve uygulama esaslarını kapsamaktadır."
          )

          yag_rapor_maddeleri = []
          for i, madde in enumerate(yag_ayirici_maddeleri, start=1):
            if i <= len(yag_ayirici_secimler) and yag_ayirici_secimler[i - 1]:
              yag_rapor_maddeleri.append(madde)

          if ek_yag_ayirici_notu.strip():
            for _not in ek_yag_ayirici_notu.split("\n"):
              if _not.strip():
                yag_rapor_maddeleri.append(_not.strip())

          if yag_rapor_maddeleri:
            doc.add_heading("6.2.3.1 YAĞ AYIRICI GENEL ESASLARI", level=3)
            for yam in yag_rapor_maddeleri:
              doc.add_paragraph(yam, style="List Bullet")

          doc.add_heading("6.2.3.2 YAĞ AYIRICI SEÇİMİ", level=3)
          doc.add_paragraph(
              "Yağ ayırıcı kapasitesi ve bağlantı çapı, proje kapsamında yapılacak debi ve ekipman "
              "bilgileri kesinleştirildiğinde ayrıca hesaplanarak seçilecektir."
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
          doc.add_heading("• YAĞMUR SUYU TOPLAMA HESABI", level=4)
          doc.add_paragraph(f"Seçilen İl: {_yr.get('mgm_il', '')}")
          _yr_yontem = _yr.get("yagis_yontemi", "Günlük Toplam En Yüksek Yağış Miktarı")
          if _yr_yontem == "Günlük Toplam En Yüksek Yağış Miktarı":
              if _yr.get("mgm_yagis_mm") is not None:
                  doc.add_paragraph(
                      f"Tasarım yağış verisi: {_yr_yontem}. "
                      f"{_yr.get('mgm_il', '')} ili için P = {_yr.get('mgm_yagis_mm', 0):.1f} mm "
                      f"({_yr.get('mgm_yagis_tarih', '')})."
                  )
          elif _yr_yontem == "Ortalama Aylık Yağış Miktarı":
              doc.add_paragraph(
                  f"Tasarım yağış verisi: {_yr_yontem}. "
                  f"12 aylık ortalama yağış değerlerinin aritmetik ortalaması ile P = "
                  f"{_yr.get('mgm_ortalama_aylik_yagis', 0):.1f} mm alınmıştır."
              )
          else:
              doc.add_paragraph(
                  f"Tasarım yağış verisi: {_yr_yontem}. "
                  f"{_yr.get('mgm_en_yuksek_ay', '')} ayındaki en yüksek aylık ortalama yağış değeri "
                  f"P = {_yr.get('mgm_en_yuksek_ay_yagis', 0):.1f} mm alınmıştır."
              )

          # Aylık MGM tablosu yalnızca 2. veya 3. tasarım yağış yöntemi seçildiğinde rapora eklenir.
          if _yr_yontem in ("Ortalama Aylık Yağış Miktarı", "En Yüksek Aylık Ortalama Yağış Miktarı"):
                        # Seçilen ilin 12 aylık ortalama yağış tablosu rapora eklenir.
                        _aylik_rapor = _yr.get("mgm_aylik_yagis", {}) or {}
                        if _aylik_rapor:
                            doc.add_heading("SEÇİLEN İLİN AYLIK ORTALAMA YAĞIŞ DEĞERLERİ", level=5)
                            _ay_tbl = doc.add_table(rows=1, cols=3)
                            _ay_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                            _ay_tbl.autofit = True
                            _ay_hdr = _ay_tbl.rows[0].cells
                            _ay_hdr[0].text = "AY"
                            _ay_hdr[1].text = "ORTALAMA YAĞIŞ (mm)"
                            _ay_hdr[2].text = "DURUM"
                            _en_ay = _yr.get("mgm_en_yuksek_ay", "")
                            for _ay in ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]:
                                if _ay not in _aylik_rapor: continue
                                _c = _ay_tbl.add_row().cells
                                _c[0].text = _ay
                                _c[1].text = f"{_aylik_rapor[_ay]:.1f}"
                                _c[2].text = "EN YÜKSEK AY" if _ay == _en_ay else ""
                                for _cell in _c:
                                    _cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                    for _p in _cell.paragraphs:
                                        _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                        for _r in _p.runs: _r.font.size = Pt(8.5)
                                if _ay == _en_ay:
                                    for _cell in _c:
                                        _tcPr = _cell._tc.get_or_add_tcPr()
                                        _shd = OxmlElement("w:shd"); _shd.set(qn("w:fill"), "FFF2CC"); _tcPr.append(_shd)
                                        for _p in _cell.paragraphs:
                                            for _r in _p.runs: _r.bold = True
                            doc.add_paragraph(
                                f"12 aylık ortalama değerlerin aritmetik ortalaması: "
                                f"{_yr.get('mgm_ortalama_aylik_yagis', 0):.1f} mm; "
                                f"en yüksek aylık ortalama: {_en_ay} = {_yr.get('mgm_en_yuksek_ay_yagis', 0):.1f} mm."
                            )

          # 81 il günlük maksimum yağış tablosu yalnızca 1. yöntem seçildiğinde rapora eklenir.
          if _yr_yontem == "Günlük Toplam En Yüksek Yağış Miktarı":
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

          # Hesap girdileri ve yüzde parametreleri.
          _yr_A = float(_yr.get('cati_alani', 0) or 0)
          _yr_P = float(_yr.get('yagis', 0) or 0)
          _yr_C = float(_yr.get('akis_katsayisi', 0) or 0)
          _yr_sarnic_orani = float(_yr.get('sarnic_orani', 80) or 0)
          _yr_filtre_etkinlik = float(_yr.get('filtre_etkinlik', 90) or 0)

          # Hesap girdileri, MGM kaynak bilgisinin hemen altında gösterilir.
          doc.add_paragraph(
              f"Toplama alanı: A = {_yr_A:.2f} m²; "
              f"tasarım yağış yüksekliği: P = {_yr_P:.2f} mm; "
              f"akış katsayısı: C = {_yr_C:.2f}"
          )

          doc.add_paragraph(
              f"Sarnıca alınacak yağmur suyu oranı: %{_yr_sarnic_orani:.0f}; "
              f"filtre etkinlik katsayısı: %{_yr_filtre_etkinlik:.0f}."
          )

          # Ham yağış hacmi ve iki yüzde parametresi uygulanarak sarnıca alınacak hacim.
          _yr_ham_V = float(_yr.get('ham_toplanabilir_m3', _yr_A * _yr_P * _yr_C / 1000.0) or 0)
          _yr_V = float(_yr.get('toplanabilir_m3', _yr_ham_V * _yr_sarnic_orani / 100.0 * _yr_filtre_etkinlik / 100.0) or 0)
          doc.add_paragraph("Toplanabilir yağmur suyu ve sarnıca alınacak su hesabı:")
          doc.add_paragraph("V_ham = A × P × C / 1000")
          doc.add_paragraph(
              f"V_ham = {_yr_A:.2f} × {_yr_P:.2f} × {_yr_C:.2f} / 1000 = {_yr_ham_V:.2f} m³"
          )
          doc.add_paragraph(
              f"V_sarnıç = V_ham × %{_yr_sarnic_orani:.0f} × %{_yr_filtre_etkinlik:.0f}"
          )
          doc.add_paragraph(
              f"V_sarnıç = {_yr_ham_V:.2f} × %{_yr_sarnic_orani:.0f} × %{_yr_filtre_etkinlik:.0f} = {_yr_V:.2f} m³"
          )

          doc.add_heading("• YAĞMUR SUYU DEPOSU HACİM HESABI", level=4)
          doc.add_paragraph(
              f"Yağmur suyu deposu tipi: {_yr.get('yagmur_depo_tipi', sih_yagmur_depo_tipi)}"
          )
          doc.add_paragraph(
              f"Yıllık toplam yağış: {_yr.get('mgm_yillik_yagis_mm', 0):.2f} mm "
              "(MGM aylık ortalama yağışlarının toplamı)"
          )
          doc.add_paragraph(
              f"Yıllık toplanabilir yağış hacmi: {_yr.get('yillik_toplam_hacim_m3', 0):.2f} m³/yıl"
          )
          doc.add_paragraph(
              f"Depolanacak oran: %{_yr.get('depolama_orani', 6):.0f}"
          )
          doc.add_paragraph(
              "V_yıllık = A × P_yıllık × C / 1000"
          )
          doc.add_paragraph(
              f"V_yıllık = {_yr.get('cati_alani', 0):.2f} × "
              f"{_yr.get('mgm_yillik_yagis_mm', 0):.2f} × "
              f"{_yr.get('akis_katsayisi', 0):.2f} / 1000 = "
              f"{_yr.get('yillik_toplam_hacim_m3', 0):.2f} m³/yıl"
          )
          doc.add_paragraph(
              f"V_depo = V_yıllık × %{_yr.get('depolama_orani', 6):.0f} = "
              f"{_yr.get('yillik_toplam_hacim_m3', 0):.2f} × "
              f"%{_yr.get('depolama_orani', 6):.0f} = "
              f"{_yr.get('gerekli_depo', 0):.2f} m³"
          )
          doc.add_paragraph(f"Hesaplanan gerekli depo hacmi: {_yr.get('gerekli_depo', 0):.2f} m³")
          _rapor_secilen_depo = float(_yr.get('secilen_depo', 0) or 0)
          _rapor_depo_tipi = str(_yr.get('yagmur_depo_tipi', sih_yagmur_depo_tipi) or '').strip()
          if _rapor_depo_tipi:
              doc.add_paragraph(
                  f'Seçilen "{_rapor_depo_tipi}" depo hacmi (Emniyetle): {_rapor_secilen_depo:.2f} m³'
              )
          else:
              doc.add_paragraph(
                  f"Seçilen depo hacmi (Emniyetle): {_rapor_secilen_depo:.2f} m³"
              )
          _rapor_yagmur_depo_poz = str(_yr.get("yagmur_depo_poz", "") or "").strip()
          if _yr.get("yagmur_depo_poz_rapora_eklensin", True) and _rapor_yagmur_depo_poz:
              doc.add_paragraph(f"Cihaz Poz No: {_rapor_yagmur_depo_poz}")

          doc.add_heading("• YAĞMUR SUYU FİLTRESİ SEÇİMİ", level=4)
          doc.add_paragraph(f"Filtre tipi: {_yr.get('filtre_tipi', '')}")
          doc.add_paragraph(
              f"Yağmur suyu toplama alanı: {_yr.get('cati_alani', 0):.2f} m²"
          )
          doc.add_paragraph(
              f"Filtre adedi: {_yr.get('filtre_adet', 1):.0f} adet"
          )
          doc.add_paragraph(
              f"Filtre başına düşen toplama alanı: {_yr.get('filtre_basina_alan_m2', 0):.2f} m²"
          )
          doc.add_paragraph(
              f"Tek filtre kapasitesi: {_yr.get('filtre_kapasite_m2', 0):.0f} m²/adet"
          )
          doc.add_paragraph(
              f"Toplam filtre kapasitesi: {_yr.get('filtre_toplam_kapasite_m2', _yr.get('filtre_kapasite_m2', 0)):,.0f} m²"
          )
          doc.add_paragraph(
              f"Tek filtre maksimum debisi: {_yr.get('filtre_debisi_ls', 0):.0f} L/s; "
              f"toplam maksimum debi: {_yr.get('filtre_toplam_debisi_ls', _yr.get('filtre_debisi_ls', 0)):,.0f} L/s"
          )
          if _yr.get('filtre_poz_rapora_eklensin', True):
              doc.add_paragraph(
                  f"Cihaz Poz No: {_yr.get('filtre_poz', '')}"
              )
          if _yr.get('filtre_kapasite_yetersiz', False):
              doc.add_paragraph(
                  "UYARI: Seçilen filtre tipi için mevcut en büyük poz kapasitesi, "
                  "toplama alanını karşılamamaktadır."
              )

          doc.add_heading("• TAŞMA HATTI HESABI", level=4)
          _tasma_A = _yr.get('cati_alani', 0)
          _tasma_P = _yr.get('yagis', 0)
          _tasma_C = _yr.get('akis_katsayisi', 0)
          _tasma_Vham = _yr.get('ham_toplanabilir_m3', 0)
          _tasma_t = _yr.get('sure_dk', 0)
          _tasma_Q = _yr.get('debi_m3h', 0)
          _tasma_E = _yr.get('tasma_emniyet', 0)
          _tasma_Qson = _yr.get('tasma_debisi', 0)
          _tasma_DN = _yr.get('tasma_cap', 0)
          _tasma_Q_lps = _yr.get('tasma_Q_lps', _tasma_Qson / 3.6)
          _tasma_hesaplanan_Q_lps = _yr.get('tasma_hesaplanan_Q_lps', _tasma_Qson / 3.6)
          _tasma_tasarim_Q_lps = _yr.get('tasma_tasarim_Q_lps', _tasma_Q_lps)
          _tasma_tasarim_debisi_m3h = _yr.get('tasma_tasarim_debisi_m3h', _tasma_tasarim_Q_lps * 3.6)
          _tasma_debi_sinirlandi = _yr.get('tasma_debi_sinirlandi', False)
          _tasma_maks_tasarim_Q_lps = _yr.get('tasma_maks_tasarim_Q_lps', 200.0)
          _tasma_n = _yr.get('tasma_manning_n', 0.011)
          _tasma_malzeme = _yr.get('tasma_malzeme', 'PVC')
          _tasma_egim = _yr.get('tasma_egim_yuzde', 1.0)
          _tasma_max_hiz = _yr.get('tasma_max_hiz', 3.0)
          _tasma_kapasite = _yr.get('tasma_hidrolik_kapasite_lps', 0.0)
          _tasma_hiz = _yr.get('tasma_hidrolik_hiz_ms', 0.0)
          _tasma_uygun = _yr.get('tasma_hidrolik_uygun', False)
          _tasma_hat_adedi = int(_yr.get('tasma_hat_adedi', 1) or 1)
          _tasma_hat_Q_lps = float(_yr.get('tasma_hat_Q_lps', _tasma_tasarim_Q_lps / _tasma_hat_adedi) or 0)
          _tasma_toplam_kapasite_lps = float(_yr.get('tasma_toplam_kapasite_lps', _tasma_kapasite * _tasma_hat_adedi) or 0)
          _tasma_onerilen_hat_adedi = _yr.get('tasma_onerilen_hat_adedi')
          _tasma_oneri_cap = _yr.get('tasma_oneri_cap')

          doc.add_paragraph("Taşma hattı hesabında çatıdan oluşan ham yağmur suyu hacmi, tasarım yağış süresine dağıtılmış ve ardından emniyet katsayısı uygulanmıştır.")
          doc.add_paragraph("1. Ham yağmur suyu hacmi:")
          doc.add_paragraph("Vham = A × P × C / 1000")
          doc.add_paragraph(
              f"Vham = {_tasma_A:,.2f} m² × {_tasma_P:,.2f} mm × {_tasma_C:.2f} / 1000 = {_tasma_Vham:,.2f} m³"
          )
          doc.add_paragraph("2. Tasarım yağış debisi:")
          doc.add_paragraph("Qyağış = Vham / (t / 60)")
          doc.add_paragraph(
              f"Qyağış = {_tasma_Vham:,.2f} m³ / ({_tasma_t:,.2f} / 60) = {_tasma_Q:,.2f} m³/h"
          )
          doc.add_paragraph("3. Emniyet katsayısı uygulanmış taşma debisi:")
          doc.add_paragraph("Qtaşma = Qyağış × (1 + E / 100)")
          doc.add_paragraph(
              f"Qtaşma = {_tasma_Q:,.2f} × (1 + {_tasma_E:.2f} / 100) = {_tasma_Qson:,.2f} m³/h = {_tasma_hesaplanan_Q_lps:,.2f} L/s"
          )
          doc.add_paragraph(
              f"Proje tasarım kriteri: taşma hattı tasarım debisi maksimum {_tasma_maks_tasarim_Q_lps:,.0f} L/s olarak sınırlandırılmıştır."
          )
          doc.add_paragraph(
              f"Tasarım taşma debisi = min({_tasma_hesaplanan_Q_lps:,.2f}, {_tasma_maks_tasarim_Q_lps:,.0f}) = {_tasma_tasarim_Q_lps:,.2f} L/s = {_tasma_tasarim_debisi_m3h:,.2f} m³/h"
          )
          if _tasma_debi_sinirlandi:
              doc.add_paragraph(
                  "Not: Hesaplanan taşma debisi 200 L/s üst sınırını aştığı için hidrolik ön boyutlandırmada "
                  "tasarım debisi 200 L/s alınmıştır. Hesaplanan gerçek debi ayrıca yukarıda gösterilmiştir."
              )
          doc.add_paragraph(
              f"Taşma hattı adedi: {_tasma_hat_adedi} adet; her hatta düşen tasarım debisi = "
              f"{_tasma_tasarim_Q_lps:,.2f} / {_tasma_hat_adedi} = {_tasma_hat_Q_lps:,.2f} L/s"
          )
          if _tasma_onerilen_hat_adedi is not None:
              doc.add_paragraph(
                  f"DN250 sınırı altında tasarım debisini karşılayan önerilen minimum düzen: "
                  f"{int(_tasma_onerilen_hat_adedi)} hat × DN {int(_tasma_oneri_cap)}."
              )
          doc.add_paragraph("4. Taşma hattı hidrolik kontrolü (Manning yöntemi):")
          doc.add_paragraph(
              f"Boru malzemesi: {_tasma_malzeme}; Manning katsayısı n = {_tasma_n:.3f}; boru eğimi = %{_tasma_egim:.2f}; "
              f"izin verilen maksimum hız = {_tasma_max_hiz:.2f} m/s"
          )
          doc.add_paragraph("Q gerekli = Qtaşma / 3,6")
          doc.add_paragraph(
              f"Q gerekli = {_tasma_tasarim_debisi_m3h:,.2f} m³/h / 3,6 = {_tasma_hat_Q_lps:,.2f} L/s/hat = {_tasma_hat_Q_lps/1000.0:,.4f} m³/s/hat"
          )
          doc.add_paragraph("Manning formülü:")
          doc.add_paragraph("Q = (1/n) × A × R^(2/3) × S^(1/2)")
          _rep_secili_manning = next((x for x in _yr.get('tasma_hidrolik_tablo', []) if int(x.get('dn', 0)) == int(_tasma_DN)), None)
          if _rep_secili_manning:
              _rep_D = float(_rep_secili_manning.get('dn', 0)) / 1000.0
              _rep_A = float(_rep_secili_manning.get('alan_m2', 0.0))
              _rep_R = _rep_D / 4.0
              _rep_Qm3s = float(_rep_secili_manning.get('q_manning_lps', 0.0)) / 1000.0
              _rep_V = float(_rep_secili_manning.get('hiz_ms', 0.0))
              doc.add_paragraph(
                  f"D = {_rep_D:.3f} m; A = π × D² / 4 = π × {_rep_D:.3f}² / 4 = {_rep_A:.5f} m²; "
                  f"R = D / 4 = {_rep_D:.3f} / 4 = {_rep_R:.5f} m; S = {_tasma_egim/100:.4f}."
              )
              doc.add_paragraph(
                  f"Q = (1 / {_tasma_n:.3f}) × {_rep_A:.5f} × ({_rep_R:.5f})^(2/3) × "
                  f"({_tasma_egim/100:.4f})^(1/2) = {_rep_Qm3s:.5f} m³/s = {_rep_Qm3s*1000:.2f} L/s."
              )
              doc.add_paragraph(
                  f"V = Q / A = {_rep_Qm3s:.5f} / {_rep_A:.5f} = {_rep_V:.2f} m/s; "
                  f"kabul edilen maksimum hız = {_tasma_max_hiz:.2f} m/s."
              )
          doc.add_paragraph(
              f"Seçilen minimum taşma hattı: {_tasma_hat_adedi} hat × DN {_tasma_DN}; "
              f"tek hat tasarım kapasitesi = {_tasma_kapasite:,.2f} L/s; "
              f"toplam tasarım kapasitesi = {_tasma_toplam_kapasite_lps:,.2f} L/s; "
              f"Manning hızı = {_tasma_hiz:.2f} m/s"
          )
          # Hidrolik kontrol tablosu rapor bölümünde kullanılmadan önce alınmalıdır.
          # Aksi halde DN200/DN250 kapasite satırında değişken tanımsız kalır ve
          # rapor oluşturma işlemi NameError ile durur.
          _hidrolik_tablo = _yr.get('tasma_hidrolik_tablo', [])
          if _tasma_DN in (200, 250):
              _rep_secili = next((x for x in _hidrolik_tablo if int(x.get('dn', 0)) == int(_tasma_DN)), None)
              if _rep_secili and _rep_secili.get('q_hiz_lps') is not None:
                  doc.add_paragraph(
                      f"DN {_tasma_DN} için {_tasma_max_hiz:.2f} m/s tasarım hızına göre kapasite: "
                      f"Q = A × V = {_rep_secili.get('alan_m2', 0):.5f} × {_tasma_max_hiz:.2f} "
                      f"= {_rep_secili.get('q_hiz_lps', 0):.2f} L/s."
                  )
          doc.add_paragraph(
              f"Hidrolik kontrol sonucu: {'UYGUN' if _tasma_uygun else 'YETERSİZ'}"
          )
          if _hidrolik_tablo:
              doc.add_paragraph("Kontrol edilen çaplar:")
              for _x in _hidrolik_tablo:
                  doc.add_paragraph(
                      f"DN {_x.get('dn', 0)} → tasarım kapasitesi {_x.get('q_kapasite_lps', 0):.2f} L/s; "
                      f"Manning kapasitesi {_x.get('q_manning_lps', _x.get('q_kapasite_lps', 0)):.2f} L/s; "
                      f"Manning hızı {_x.get('hiz_ms', 0):.2f} m/s; "
                      f"{'UYGUN' if _x.get('uygun') else 'YETERSİZ'}"
                  )
          doc.add_paragraph(
              "Not: Bu kontrol, taşma hattını cazibeli ve tam dolu dairesel boru kabulüyle Manning kapasitesi üzerinden ön boyutlandırır. "
              "Son proje kontrolünde gerçek kotlar, çıkış koşulu ve akış rejimi ayrıca doğrulanmalıdır."
          )

          doc.add_heading("• TAŞKAN SİFONU / KOKU KAPANI SEÇİMİ", level=4)
          doc.add_paragraph(
              "Taşkan sifonu / koku kapanı kullanılacaktır." if _yr.get("sifon") else
              "Taşkan sifonu / koku kapanı öngörülmemiştir."
          )
          if _yr.get("sifon"):
              _sifon_poz = str(_yr.get("tasma_sifonu_poz", "") or "").strip()
              _sifon_dn = int(_yr.get("tasma_sifonu_dn", 0) or 0)
              _sifon_tanim = str(_yr.get("tasma_sifonu_tanim", "") or "").strip()
              _sifon_ozellik = str(_yr.get("tasma_sifonu_ozellik", "") or "").strip()
              _sifon_adedi = int(_yr.get("tasma_sifonu_adedi", _tasma_hat_adedi) or _tasma_hat_adedi)
              doc.add_paragraph(
                  f"Hidrolik hesapta kullanılan toplam tasarım taşma debisi: {_tasma_tasarim_Q_lps:,.2f} L/s "
                  f"(üst sınır {_tasma_maks_tasarim_Q_lps:,.0f} L/s)."
              )
              doc.add_paragraph(
                  f"Taşma hattı düzeni: {_tasma_hat_adedi} paralel hat; her hat için tasarım debisi = "
                  f"{_tasma_hat_Q_lps:,.2f} L/s."
              )
              doc.add_paragraph(
                  f"Hidrolik hesap sonucu her hat için gerekli minimum taşma hattı: DN {_tasma_DN}."
              )
              if _sifon_tanim:
                  doc.add_paragraph(f"Seçilen Taşkan Sifonu: {_sifon_adedi} adet × {_sifon_tanim}")
              if _sifon_ozellik:
                  doc.add_paragraph(f"Taşkan Sifonu Özelliği: {_sifon_ozellik}")
              if _yr.get("tasma_sifonu_poz_rapora_eklensin") and _sifon_poz:
                  doc.add_paragraph(f"Taşkan Sifonu Cihaz Poz No: {_sifon_poz}")
              elif _sifon_poz:
                  doc.add_paragraph("Taşkan Sifonu Cihaz Poz No rapora eklenmemiştir.")
          if _yr.get("kanal_baglanti"):
              doc.add_paragraph("Taşma hattı kanalizasyona bağlanacaktır; geri tepme koruması sağlanacaktır.")
          else:
              doc.add_paragraph("Taşma hattı kanalizasyona bağlanmayacaktır.")
          if _yr.get("geri_tepme"):
              doc.add_paragraph("Geri tepme önleyici düzenek öngörülmüştür.")

          doc.add_heading("• AKIŞ DÜZENLEYİCİ (CAZİBE YAVAŞLATICI / SAKİNLEŞTİRİCİ GİRİŞ) SEÇİMİ", level=4)
          if _yr.get("sakin_giris"):
              doc.add_paragraph("Akış düzenleyici (cazibe yavaşlatıcı / sakinleştirici giriş) kullanılacaktır.")
              doc.add_paragraph(f"Malzeme / Özellik: {_yr.get('sakin_giris_malzeme_ozellik', '')}")
              doc.add_paragraph(f"Fonksiyonu: {_yr.get('sakin_giris_fonksiyonu', '')}")
              if _yr.get("sakin_giris_poz_rapora_eklensin"):
                  doc.add_paragraph(f"Cihaz Poz No: {_yr.get('sakin_giris_poz', '25.181.5300')}")
          else:
              doc.add_paragraph("Akış düzenleyici (cazibe yavaşlatıcı / sakinleştirici giriş) öngörülmemiştir.")

          doc.add_heading("• HAVALANDIRMA VE HAŞERE KORUMASI", level=4)
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

    # Rapor oluşturma bloğu sayfanın sonunda çalıştığı için, üstteki indirme
    # düğmesinin yeni oluşturulan dosyayı gösterebilmesi amacıyla bir kez
    # yeniden çalıştırılır. İstek zaten pop edildiğinden döngü oluşmaz.
    st.rerun()


# SAYFA SONU ANKORU
st.markdown('<div id="sayfa_sonu"></div>', unsafe_allow_html=True)
