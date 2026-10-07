"""
Mühendislik Proje Raporu Otomasyonu
Proje Yönetimi Modülü

Bu modül mevcut hesaplama ve raporlama fonksiyonlarından bağımsızdır.
Proje oluşturma, kaydetme, açma, güncelleme ve rapor sürüm takibini sağlar.
"""

from __future__ import annotations

import json
import os
import re
import urllib.parse
import requests
import shutil
import tempfile
from copy import deepcopy
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional


# -----------------------------------------------------------------------------
# AYARLAR
# -----------------------------------------------------------------------------

VARSAYILAN_PROJE_DIZINI = Path("projects")
PROJE_DOSYASI = "proje.json"
HESAPLAR_DOSYASI = "hesaplar.json"
RAPOR_DIZINI = "raporlar"
GRAFIK_DIZINI = "grafikler"


# -----------------------------------------------------------------------------
# YARDIMCI FONKSİYONLAR
# -----------------------------------------------------------------------------


def simdi_iso() -> str:
    """Yerel zamanı ISO formatında döndürür."""
    return datetime.now().isoformat(timespec="seconds")


def guvenli_dosya_adi(metin: str) -> str:
    """Klasör ve dosya adlarında kullanılabilecek güvenli metin üretir."""
    metin = str(metin).strip()
    metin = re.sub(r"[^a-zA-Z0-9çÇğĞıİöÖşŞüÜ._ -]", "", metin)
    metin = re.sub(r"\s+", "_", metin)
    metin = re.sub(r"_+", "_", metin)
    return metin.strip("._") or "proje"


def json_kaydet_guvenli(dosya_yolu: Path, veri: Dict[str, Any]) -> None:
    """JSON verisini geçici dosyaya yazıp atomik şekilde hedefe taşır."""
    dosya_yolu.parent.mkdir(parents=True, exist_ok=True)

    gecici_dosya = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".tmp",
            dir=dosya_yolu.parent,
            delete=False,
        ) as dosya:
            json.dump(veri, dosya, ensure_ascii=False, indent=2, default=str)
            dosya.flush()
            gecici_dosya = Path(dosya.name)

        gecici_dosya.replace(dosya_yolu)
    finally:
        if gecici_dosya and gecici_dosya.exists():
            gecici_dosya.unlink(missing_ok=True)


def json_oku(dosya_yolu: Path, varsayilan: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """JSON dosyasını okur; dosya yoksa varsayılan değeri döndürür."""
    if not dosya_yolu.exists():
        return deepcopy(varsayilan or {})

    with dosya_yolu.open("r", encoding="utf-8") as dosya:
        veri = json.load(dosya)

    if not isinstance(veri, dict):
        raise ValueError(f"Beklenen JSON nesnesi bulunamadı: {dosya_yolu}")

    return veri


# -----------------------------------------------------------------------------
# SUPABASE UZAK KAYIT
# -----------------------------------------------------------------------------


def _supabase_ayarlari() -> tuple[str, str]:
    """Supabase bilgilerini önce ortam değişkenlerinden, sonra Streamlit Secrets'tan alır.

    Streamlit Cloud'da Secrets çoğu kurulumda os.environ'a otomatik aktarılmadığı
    için yalnızca os.getenv kullanmak kalıcı proje kayıtlarının devre dışı kalmasına
    ve uygulama yeniden başladığında projelerin kaybolmuş gibi görünmesine neden olur.
    """
    url = str(os.getenv("SUPABASE_URL", "")).strip().rstrip("/")
    key = str(os.getenv("SUPABASE_SECRET_KEY", "")).strip()

    if not url or not key:
        try:
            import streamlit as st
            _secrets = st.secrets
            if not url:
                url = str(_secrets.get("SUPABASE_URL", "")).strip().rstrip("/")
            if not key:
                key = str(
                    _secrets.get("SUPABASE_SECRET_KEY", "")
                    or _secrets.get("SUPABASE_KEY", "")
                ).strip()
        except Exception:
            pass

    return url, key


def _supabase_aktif_mi() -> bool:
    url, key = _supabase_ayarlari()
    return bool(url and key)


def _supabase_headers() -> Dict[str, str]:
    _, key = _supabase_ayarlari()
    return {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates,return=representation",
    }


def _supabase_url(path: str) -> str:
    url, _ = _supabase_ayarlari()
    return f"{url}/rest/v1/{path.lstrip('/')}"


def _supabase_hata(response: requests.Response) -> str:
    try:
        veri = response.json()
        if isinstance(veri, dict):
            return str(veri.get("message") or veri.get("error") or veri)
        return str(veri)
    except Exception:
        return response.text[:500] or f"HTTP {response.status_code}"


def _supabase_listele() -> List[Dict[str, Any]]:
    if not _supabase_aktif_mi():
        return []
    response = requests.get(
        _supabase_url("projects"),
        params={"select": "project_id,project_name,created_at,updated_at", "order": "project_name.asc"},
        headers=_supabase_headers(),
        timeout=15,
    )
    if not response.ok:
        raise RuntimeError(f"Supabase proje listesi alınamadı: {_supabase_hata(response)}")
    veri = response.json()
    return veri if isinstance(veri, list) else []


def _supabase_getir(proje_id: str) -> Optional[Dict[str, Any]]:
    if not _supabase_aktif_mi():
        return None
    response = requests.get(
        _supabase_url("projects"),
        params={"project_id": f"eq.{proje_id}", "select": "*", "limit": "1"},
        headers=_supabase_headers(),
        timeout=15,
    )
    if not response.ok:
        raise RuntimeError(f"Supabase proje açılamadı: {_supabase_hata(response)}")
    veri = response.json()
    return veri[0] if veri else None


def _supabase_kaydet(proje: "ProjeVerisi") -> None:
    if not _supabase_aktif_mi():
        return
    payload = {
        "project_id": proje.proje_id,
        "project_name": proje.proje_adi,
        "data": proje.sozluk(),
        "updated_at": proje.guncelleme_tarihi or simdi_iso(),
    }
    response = requests.post(
        _supabase_url("projects"),
        headers=_supabase_headers(),
        json=payload,
        timeout=20,
    )
    if not response.ok:
        raise RuntimeError(f"Supabase proje kaydedilemedi: {_supabase_hata(response)}")


def _supabase_sil(proje_id: str) -> None:
    if not _supabase_aktif_mi():
        return
    response = requests.delete(
        _supabase_url("projects"),
        params={"project_id": f"eq.{proje_id}"},
        headers=_supabase_headers(),
        timeout=15,
    )
    if not response.ok:
        raise RuntimeError(f"Supabase proje silinemedi: {_supabase_hata(response)}")


# -----------------------------------------------------------------------------
# VERİ MODELİ
# -----------------------------------------------------------------------------


@dataclass
class ProjeVerisi:
    """Projeye ait tüm bölümlerin ortak veri taşıyıcısı."""

    proje_id: str
    proje_adi: str
    sirket_adi: str = "FUGA MEKANİK MÜHENDİSLİK MÜŞAVİRLİK İNŞ.SAN.TİC.LTD.ŞTİ"
    is_veren: str = ""
    adres: str = ""
    rapor_turu: str = "MEKANİK TESİSAT UYGULAMA PROJESİ HESAP RAPORU"
    hazirlayan: str = "Mehmet Küçük"
    mmo_no: str = "109913"
    rapor_tarihi: str = ""

    genel: Dict[str, Any] = field(default_factory=dict)
    standartlar: Dict[str, Any] = field(default_factory=dict)
    proje_kapsami: Dict[str, Any] = field(default_factory=dict)
    isi_iletim_akiskanlari: Dict[str, Any] = field(default_factory=dict)
    iklim: Dict[str, Any] = field(default_factory=dict)
    sihhi_tesisat: Dict[str, Any] = field(default_factory=dict)
    modul_07: Dict[str, Any] = field(default_factory=dict)
    modul_08: Dict[str, Any] = field(default_factory=dict)
    modul_09: Dict[str, Any] = field(default_factory=dict)
    hesap_sonuclari: Dict[str, Any] = field(default_factory=dict)

    olusturma_tarihi: str = field(default_factory=simdi_iso)
    guncelleme_tarihi: str = field(default_factory=simdi_iso)
    son_rapor_tarihi: str = ""

    def sozluk(self) -> Dict[str, Any]:
        """Dataclass nesnesini JSON'a uygun sözlüğe çevirir."""
        return asdict(self)

    @classmethod
    def sozlukten_olustur(cls, veri: Dict[str, Any]) -> "ProjeVerisi":
        """Sözlükten ProjeVerisi nesnesi oluşturur."""
        alanlar = set(cls.__dataclass_fields__.keys())
        uygun_veri = {anahtar: deger for anahtar, deger in veri.items() if anahtar in alanlar}
        return cls(**uygun_veri)


# -----------------------------------------------------------------------------
# PROJE YÖNETİCİSİ
# -----------------------------------------------------------------------------


class ProjeYoneticisi:
    """Projelerin klasör, veri ve rapor sürümü yönetimini sağlar."""

    def __init__(self, ana_dizin: str | Path = VARSAYILAN_PROJE_DIZINI):
        self.ana_dizin = Path(ana_dizin)
        self.ana_dizin.mkdir(parents=True, exist_ok=True)

    def proje_dizini(self, proje_id: str) -> Path:
        """Proje ID'sine göre proje klasörünü döndürür."""
        return self.ana_dizin / guvenli_dosya_adi(proje_id)

    def proje_dosyasi(self, proje_id: str) -> Path:
        return self.proje_dizini(proje_id) / PROJE_DOSYASI

    def hesaplar_dosyasi(self, proje_id: str) -> Path:
        return self.proje_dizini(proje_id) / HESAPLAR_DOSYASI

    def rapor_dizini(self, proje_id: str) -> Path:
        return self.proje_dizini(proje_id) / RAPOR_DIZINI

    def grafik_dizini(self, proje_id: str) -> Path:
        return self.proje_dizini(proje_id) / GRAFIK_DIZINI

    def proje_var_mi(self, proje_id: str) -> bool:
        return self.proje_dosyasi(proje_id).exists()

    def proje_olustur(
        self,
        proje_adi: str,
        proje_id: Optional[str] = None,
        **bilgiler: Any,
    ) -> ProjeVerisi:
        """Yeni proje oluşturur ve ilk kayıt dosyasını yazar."""
        proje_adi = str(proje_adi).strip()
        if not proje_adi:
            raise ValueError("Proje adı boş bırakılamaz.")

        proje_id = guvenli_dosya_adi(proje_id or proje_adi)

        if self.proje_var_mi(proje_id):
            raise FileExistsError(f"Bu proje zaten mevcut: {proje_id}")

        proje = ProjeVerisi(
            proje_id=proje_id,
            proje_adi=proje_adi,
            rapor_tarihi=bilgiler.pop("rapor_tarihi", datetime.now().strftime("%d.%m.%Y")),
            **bilgiler,
        )

        self._klasorleri_hazirla(proje_id)
        self.kaydet(proje)
        return proje

    def _klasorleri_hazirla(self, proje_id: str) -> None:
        """Proje için gerekli klasörleri oluşturur."""
        proje_dizini = self.proje_dizini(proje_id)
        (proje_dizini / RAPOR_DIZINI).mkdir(parents=True, exist_ok=True)
        (proje_dizini / GRAFIK_DIZINI).mkdir(parents=True, exist_ok=True)

    def kaydet(
        self,
        proje: ProjeVerisi,
        hesap_sonuclari: Optional[Dict[str, Any]] = None,
    ) -> ProjeVerisi:
        """Proje verisini ve varsa hesap sonuçlarını güvenli şekilde kaydeder."""
        if not proje.proje_id:
            raise ValueError("Proje ID boş olamaz.")

        if not proje.proje_adi:
            raise ValueError("Proje adı boş olamaz.")

        self._klasorleri_hazirla(proje.proje_id)
        proje.guncelleme_tarihi = simdi_iso()

        json_kaydet_guvenli(self.proje_dosyasi(proje.proje_id), proje.sozluk())

        if hesap_sonuclari is not None:
            proje.hesap_sonuclari = deepcopy(hesap_sonuclari)
            json_kaydet_guvenli(
                self.hesaplar_dosyasi(proje.proje_id),
                proje.hesap_sonuclari,
            )
        elif not self.hesaplar_dosyasi(proje.proje_id).exists():
            json_kaydet_guvenli(
                self.hesaplar_dosyasi(proje.proje_id),
                proje.hesap_sonuclari,
            )

        # Streamlit Cloud kalıcı disk sağlamadığı için Supabase etkinse
        # aynı projenin tamamını uzak veritabanına da kaydet.
        if _supabase_aktif_mi():
            _supabase_kaydet(proje)

        return proje

    def ac(self, proje_id: str) -> ProjeVerisi:
        """Projeyi önce Supabase'den, bulunamazsa yerel önbellekten açar."""
        proje_id = guvenli_dosya_adi(proje_id)

        if _supabase_aktif_mi():
            uzak = _supabase_getir(proje_id)
            if uzak and isinstance(uzak.get("data"), dict):
                proje = ProjeVerisi.sozlukten_olustur(uzak["data"])
                proje.proje_id = proje_id
                proje.proje_adi = str(uzak.get("project_name") or proje.proje_adi)
                self._klasorleri_hazirla(proje_id)
                json_kaydet_guvenli(self.proje_dosyasi(proje_id), proje.sozluk())
                json_kaydet_guvenli(self.hesaplar_dosyasi(proje_id), proje.hesap_sonuclari)
                return proje

        dosya = self.proje_dosyasi(proje_id)
        if not dosya.exists():
            raise FileNotFoundError(f"Proje bulunamadı: {proje_id}")

        proje = ProjeVerisi.sozlukten_olustur(json_oku(dosya))
        proje.proje_id = proje_id
        hesaplar = json_oku(self.hesaplar_dosyasi(proje_id), {})
        proje.hesap_sonuclari = hesaplar
        return proje

    def guncelle(self, proje: ProjeVerisi, **degisiklikler: Any) -> ProjeVerisi:
        """Açılmış proje üzerinde alan bazlı değişiklik yapar ve kaydeder."""
        alanlar = set(ProjeVerisi.__dataclass_fields__.keys())

        for anahtar, deger in degisiklikler.items():
            if anahtar not in alanlar:
                raise KeyError(f"Geçersiz proje alanı: {anahtar}")
            if anahtar in {"proje_id", "olusturma_tarihi"}:
                raise ValueError(f"Bu alan güncellenemez: {anahtar}")
            setattr(proje, anahtar, deger)

        return self.kaydet(proje)

    def hesap_sonuclarini_kaydet(
        self,
        proje: ProjeVerisi,
        hesap_sonuclari: Dict[str, Any],
    ) -> ProjeVerisi:
        """Hesap sonuçlarını proje verisinden ayrı JSON dosyasına kaydeder."""
        proje.hesap_sonuclari = deepcopy(hesap_sonuclari)
        proje.guncelleme_tarihi = simdi_iso()

        json_kaydet_guvenli(
            self.hesaplar_dosyasi(proje.proje_id),
            proje.hesap_sonuclari,
        )
        json_kaydet_guvenli(self.proje_dosyasi(proje.proje_id), proje.sozluk())
        return proje

    def rapor_kaydet(
        self,
        proje: ProjeVerisi,
        rapor_bytes: bytes,
        dosya_uzantisi: str = ".docx",
        rapor_etiketi: str = "Rapor",
    ) -> Path:
        """Oluşturulan raporu eski raporları silmeden sürümlü kaydeder."""
        if not isinstance(rapor_bytes, (bytes, bytearray)):
            raise TypeError("rapor_bytes bytes veya bytearray olmalıdır.")

        uzanti = dosya_uzantisi if dosya_uzantisi.startswith(".") else f".{dosya_uzantisi}"
        rapor_dizini = self.rapor_dizini(proje.proje_id)
        rapor_dizini.mkdir(parents=True, exist_ok=True)

        mevcutlar = list(rapor_dizini.glob(f"{guvenli_dosya_adi(rapor_etiketi)}_v*{uzanti}"))
        sonraki_surum = len(mevcutlar) + 1
        tarih = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        dosya_adi = (
            f"{guvenli_dosya_adi(rapor_etiketi)}_v{sonraki_surum:02d}_{tarih}{uzanti}"
        )
        hedef = rapor_dizini / dosya_adi

        with hedef.open("wb") as dosya:
            dosya.write(rapor_bytes)

        proje.son_rapor_tarihi = simdi_iso()
        proje.guncelleme_tarihi = simdi_iso()
        json_kaydet_guvenli(self.proje_dosyasi(proje.proje_id), proje.sozluk())
        return hedef

    def proje_listesi(self) -> List[Dict[str, Any]]:
        """Yerel önbellek ile Supabase projelerini birleştirir."""
        birlesik: Dict[str, Dict[str, Any]] = {}

        try:
            for item in self._yerel_proje_listesi():
                birlesik[str(item["proje_id"])] = item
        except Exception:
            pass

        if _supabase_aktif_mi():
            uzaklar = _supabase_listele()
            for item in uzaklar:
                pid = str(item.get("project_id", "")).strip()
                if not pid:
                    continue
                birlesik[pid] = {
                    "proje_id": pid,
                    "proje_adi": str(item.get("project_name") or pid),
                    "olusturma_tarihi": str(item.get("created_at") or ""),
                    "guncelleme_tarihi": str(item.get("updated_at") or ""),
                    "son_rapor_tarihi": "",
                }

        return sorted(birlesik.values(), key=lambda x: str(x.get("proje_adi", "")).lower())

    def _yerel_proje_listesi(self) -> List[Dict[str, Any]]:
        liste: List[Dict[str, Any]] = []
        if not self.ana_dizin.exists():
            return liste
        for proje_dizini in sorted(self.ana_dizin.iterdir()):
            if not proje_dizini.is_dir():
                continue
            proje_dosyasi = proje_dizini / PROJE_DOSYASI
            if not proje_dosyasi.exists():
                continue
            try:
                veri = json_oku(proje_dosyasi)
                liste.append({
                    "proje_id": veri.get("proje_id", proje_dizini.name),
                    "proje_adi": veri.get("proje_adi", proje_dizini.name),
                    "olusturma_tarihi": veri.get("olusturma_tarihi", ""),
                    "guncelleme_tarihi": veri.get("guncelleme_tarihi", ""),
                    "son_rapor_tarihi": veri.get("son_rapor_tarihi", ""),
                })
            except (OSError, json.JSONDecodeError, ValueError):
                continue
        return liste

    def rapor_listesi(self, proje_id: str) -> List[Path]:
        """Proje için oluşturulmuş tüm raporları tarihe göre listeler."""
        dizin = self.rapor_dizini(proje_id)
        if not dizin.exists():
            return []
        return sorted(dizin.iterdir(), key=lambda yol: yol.stat().st_mtime, reverse=True)

    def proje_sil(self, proje_id: str, onay: bool = False) -> None:
        """Projeyi tamamen siler. Güvenlik için açık onay ister."""
        if not onay:
            raise PermissionError("Projeyi silmek için onay=True verilmelidir.")

        dizin = self.proje_dizini(proje_id)
        yerel_var = dizin.exists()
        uzak_var = False
        if _supabase_aktif_mi():
            uzak_var = _supabase_getir(guvenli_dosya_adi(proje_id)) is not None

        if not yerel_var and not uzak_var:
            raise FileNotFoundError(f"Proje bulunamadı: {proje_id}")

        if yerel_var:
            shutil.rmtree(dizin)
        if _supabase_aktif_mi():
            _supabase_sil(guvenli_dosya_adi(proje_id))


# -----------------------------------------------------------------------------
# BASİT TEST / ÖRNEK KULLANIM
# -----------------------------------------------------------------------------


if __name__ == "__main__":
    yonetici = ProjeYoneticisi("projects_test")

    test_proje_id = "ornek_proje"
    if yonetici.proje_var_mi(test_proje_id):
        print(f"'{test_proje_id}' zaten mevcut; test oluşturma atlandı.")
    else:
        proje = yonetici.proje_olustur(
            proje_adi="Örnek Proje",
            proje_id=test_proje_id,
            is_veren="Örnek İşveren",
            adres="Ankara",
        )
        yonetici.hesap_sonuclarini_kaydet(
            proje,
            {
                "test_hesabi": {
                    "durum": "tamamlandı",
                    "deger": 123.45,
                }
            },
        )
        print("Test projesi oluşturuldu:", proje.proje_id)

    acilan = yonetici.ac(test_proje_id)
    print("Açılan proje:", acilan.proje_adi)
    print("Kayıtlı projeler:", yonetici.proje_listesi())
