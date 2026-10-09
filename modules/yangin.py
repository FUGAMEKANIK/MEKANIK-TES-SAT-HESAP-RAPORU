# Yangın modülü aynı app.py çalıştırması içinde iki kez çağrılırsa
# Streamlit aynı widget key'lerini ikinci kez oluşturmasın.
# Her yeni Streamlit rerun'ında app.py globals sıfırlandığı için normal çalışma etkilenmez.
if not globals().get("_YANGIN_FRAGMENT_EXECUTED", False):
    globals()["_YANGIN_FRAGMENT_EXECUTED"] = True
    with _t_yangin:
        # -----------------------------------------------------------------------
        # 7. YANGIN TESİSATI
        # Bu aşamada yalnızca bölüm/alt bölüm altyapısı oluşturulmuştur.
        # Gerçek hesap, seçim ve poz verileri sonraki adımlarda eklenecektir.
        # -----------------------------------------------------------------------
        st.header("7. YANGIN TESİSATI")
        st.caption(
            "Yangın tesisatı bölümü modüler olarak hazırlanmıştır. "
            "Hesap ve ekipman seçimleri 7.1'den itibaren adım adım eklenecektir."
        )

        _yangin_711_maddeleri = [
            "Can ve mal güvenliğinin sağlanması",
            "Yangının kontrol altına alınması ve söndürülmesi",
            "Yangın tesisatının koruma kapsamı",
            "Yangın suyu ihtiyacının karşılanması gereksinimleri",
            "Yangın dolabı sistemi tasarımı ve hesaplamaları",
            "Hidrant sistemi tasarımı ve hesaplamaları",
            "Sprinkler sistemi tasarımı ve hesaplamaları",
            "Gazlı söndürme sistemi tasarımı ve hesaplamaları",
            "Köpüklü söndürme sistemleri tasarımı ve hesaplamaları",
            "Davlumbaz söndürme sistemleri tasarımı ve hesaplamaları",
            "Merdiven basınçlandırma sistemleri tasarımı ve hesaplamaları",
            "Asansör basınçlandırma sistemleri tasarımı ve hesaplamaları",
            "Duman tahliye sistemleri tasarımı ve hesaplamaları",
            "Yangın pompa grubunun tasarımı ve hesaplamaları",
            "Yangın suyu depolama sisteminin tasarım ve hesaplamaları",
            "Yangın tesisatının bina otomasyonu ve diğer tesisat tasarımlarıyla ilişkisi",
            "Yangın tesisatının işletme ve bakım esasları",
        ]
        _yangin_712_maddeleri = [
            "Yangın tesisatı tasarımında yapının kullanım amacı, yapı özellikleri, kullanıcı profili ve kullanım yoğunluğunun dikkate alınması",
            "Yapının yangın güvenliği açısından değerlendirilmesinde bina yüksekliği, yapı alanı, kat sayısı ve kullanım şeklinin dikkate alınması",
            "Yangın tesisatı tasarımının yapının yangın riskini ve yangının yayılma ihtimalini azaltacak şekilde oluşturulması",
            "Yangın tehlike sınıfının belirlenmesi ve yangın söndürme sistemlerinin tasarım kriterlerinin belirlenen sınıfa göre oluşturulması",
            "Binada bulunan farklı kullanım alanlarının yangın riski bakımından ayrı ayrı değerlendirilmesi",
            "Yangın bölmeleri ve yangına dayanıklı yapı elemanlarının yangın tesisatı tasarımı üzerindeki etkilerinin değerlendirilmesi",
            "Yangın söndürme sistemlerinin korunacak alanın tamamında yeterli koruma sağlayacak şekilde yerleşiminin yapılması",
            "Yangın dolaplarının erişilebilirlik, kullanım kolaylığı ve yangına ilk müdahale şartları dikkate alınarak konumlandırılması",
            "Hidrantların bina çevresindeki yangınlara müdahaleyi sağlayacak şekilde uygun konumlarda düzenlenmesi",
            "Sprinkler sisteminin gerekli olduğu yapılarda sprinkler başlıklarının korunacak alanın geometrisi, kullanım amacı ve yangın tehlike sınıfına uygun olarak yerleştirilmesi",
            "Sprinkler tasarım alanının ve tasarım yoğunluğunun belirlenen yangın tehlike sınıfına göre oluşturulması",
            "Yangın dolabı, hidrant ve sprinkler sistemlerinin aynı yangın suyu kaynağından beslenmesi durumunda eş zamanlı çalışma şartlarının değerlendirilmesi",
            "Yangın tesisatının gerekli tasarım debisini karşılayacak şekilde boyutlandırılması", "Yangın tesisatının kritik kullanım noktasında gerekli basıncı sağlayacak şekilde tasarlanması",
            "Bina kot farklarının yangın tesisatı üzerindeki statik basınç etkisinin hesaplarda dikkate alınması", "Boru çaplarının gerekli debi ve basınç kayıpları dikkate alınarak belirlenmesi",
            "Yangın tesisatı boru şebekesinde sürtünme ve lokal basınç kayıplarının hidrolik hesaplarda dikkate alınması", "Yangın pompa grubunun sistemin tasarım debisi ve gerekli basma yüksekliğini karşılayacak şekilde seçilmesi",
            "Ana yangın pompası, yedek yangın pompası ve jokey pompanın sistemin çalışma prensibine uygun olarak birlikte değerlendirilmesi", "Yangın suyu deposunun sistemin hesaplanan su ihtiyacını karşılayabilecek kullanılabilir hacme göre belirlenmesi",
            "Yangın suyu deposunda yangın söndürme amacıyla ayrılan su rezervinin diğer kullanım amaçları nedeniyle tüketilmesinin önlenmesi", "Yangın tesisatının su kaynağı, depo, pompa grubu ve dağıtım şebekesi arasındaki hidrolik ilişkinin birlikte değerlendirilmesi",
            "Yangın tesisatının elektrik beslemesi ve otomatik çalışma sisteminin yangın anında sistemin sürekliliğini sağlayacak şekilde değerlendirilmesi", "Yangın pompa grubunun manuel ve otomatik çalışma koşullarının proje kapsamında belirlenmesi",
            "Yangın tesisatında kullanılacak ekipmanların çalışma basıncı, debi kapasitesi ve bağlantı şartlarına uygunluğunun kontrol edilmesi", "Yangın tesisatının diğer mekanik tesisatlarla kesiştiği noktalarda koordinasyonun sağlanması",
            "Yangın tesisatı borularının mimari, statik ve elektrik tesisatlarıyla koordineli şekilde geçirilmesi", "Boru askı ve desteklerinin boru ağırlığı, dolu boru ağırlığı ve işletme şartları dikkate alınarak değerlendirilmesi",
            "Yangın tesisatında bakım, test, boşaltma ve devre dışı bırakma işlemlerinin yapılabilmesi için gerekli vana, drenaj ve test bağlantılarının oluşturulması", "Yangın tesisatının devreye alınması sonrasında test ve kontrol işlemlerinin yapılabilmesine imkân verecek şekilde tasarlanması",
            "Yangın tesisatı tasarımında kullanılan bütün kabullerin ve hesap parametrelerinin proje raporunda açıkça gösterilmesi", "Yangın tesisatı tasarımında kullanılan standart ve yönetmeliklerin proje başında tanımlanması",
            "Yangın tesisatı tasarım kriterlerinin depo, pompa, hidrant, sprinkler ve hidrolik hesap modüllerine aktarılabilecek şekilde kayıt altına alınması",
        ]
        # -------------------------------------------------------------------
        # 7.1.3 YANGIN TESİSATI STANDARTLARI
        # Standartlar iki ayrı sekmede tutulur: TS / TS EN ve NFPA.
        # Her standart bağımsız olarak seçilebilir ve seçilenler rapora aktarılır.
        # -------------------------------------------------------------------
        _yangin_713_ts_standartlari = [
            "Binaların Yangından Korunması Hakkında Yönetmelik (BYKHY)",
            "TS EN 12845 — Sabit yangın söndürme sistemleri — Otomatik sprinkler sistemleri — Tasarım, kurulum ve bakım",
            "TS EN 671-1 — Sabit yangın söndürme sistemleri — Hortum sistemleri — Bölüm 1: Yarı sert hortumlu hortum makaraları",
            "TS EN 671-2 — Sabit yangın söndürme sistemleri — Hortum sistemleri — Bölüm 2: Yassı hortumlu hortum sistemleri",
            "TS EN 14384 — Yerüstü hidrantları",
            "TS EN 14339 — Yeraltı hidrantları",
            "TS EN 13565-1 — Sabit yangın söndürme sistemleri — Köpük sistemleri — Bölüm 1: Bileşenler",
            "TS EN 13565-2 — Sabit yangın söndürme sistemleri — Köpük sistemleri — Bölüm 2: Tasarım, yapım ve bakım",
            "TS EN 15004 serisi — Sabit yangın söndürme sistemleri — Gazlı söndürme sistemleri",
            "TS EN 54 serisi — Yangın algılama ve yangın alarm sistemleri",
            "TS EN 12259 serisi — Sabit yangın söndürme sistemleri — Sprinkler sistemleri bileşenleri",
            "TS EN 17451 — Sabit yangın söndürme sistemleri — Sprinkler sistemlerinin kurulumu — Proje, montaj ve kabul hususları",
            "TS EN 14816 — Sabit yangın söndürme sistemleri — Su püskürtme sistemleri — Tasarım, kurulum ve bakım",
            "TS EN 14972 serisi — Sabit yangın söndürme sistemleri — Su sisi sistemleri",
            "TS EN 13501 serisi — Yapı mamulleri ve yapı elemanlarının yangın sınıflandırması",
        ]

        _yangin_713_nfpa_standartlari = [
            "NFPA 10 — Standard for Portable Fire Extinguishers (Taşınabilir Yangın Söndürücüler Standardı)",
            "NFPA 11 — Standard for Low-, Medium-, and High-Expansion Foam (Düşük, Orta ve Yüksek Genleşmeli Köpük Standardı)",
            "NFPA 12 — Standard on Carbon Dioxide Extinguishing Systems (Karbondioksitli Söndürme Sistemleri Standardı)",
            "NFPA 12A — Standard on Halon 1301 Fire Extinguishing Systems (Halon 1301 Yangın Söndürme Sistemleri Standardı)",
            "NFPA 13 — Standard for the Installation of Sprinkler Systems (Sprinkler Sistemlerinin Kurulumu Standardı)",
            "NFPA 13D — Standard for the Installation of Sprinkler Systems in One- and Two-Family Dwellings and Manufactured Homes (Tek ve İki Ailelik Konutlar ve İmal Edilmiş Konutlarda Sprinkler Sistemlerinin Kurulumu Standardı)",
            "NFPA 13R — Standard for the Installation of Sprinkler Systems in Low-Rise Residential Occupancies (Alçak Katlı Konut Yapılarında Sprinkler Sistemlerinin Kurulumu Standardı)",
            "NFPA 14 — Standard for the Installation of Standpipe and Hose Systems (Yangın Düşey Boru ve Hortum Sistemlerinin Kurulumu Standardı)",
            "NFPA 15 — Standard for Water Spray Fixed Systems for Fire Protection (Yangından Korunma İçin Sabit Su Püskürtme Sistemleri Standardı)",
            "NFPA 16 — Standard for the Installation of Foam-Water Sprinkler and Foam-Water Spray Systems (Köpük-Su Sprinkler ve Köpük-Su Püskürtme Sistemlerinin Kurulumu Standardı)",
            "NFPA 17 — Standard for Dry Chemical Extinguishing Systems (Kuru Kimyevi Tozlu Söndürme Sistemleri Standardı)",
            "NFPA 17A — Standard for Wet Chemical Extinguishing Systems (Islak Kimyasal Söndürme Sistemleri Standardı)",
            "NFPA 18 — Standard on Wetting Agents (Islatıcı Maddeler Standardı)",
            "NFPA 18A — Standard on Water Additives for Fire Control and Vapor Mitigation (Yangın Kontrolü ve Buhar Azaltımı İçin Su Katkı Maddeleri Standardı)",
            "NFPA 20 — Standard for the Installation of Stationary Pumps for Fire Protection (Yangından Korunma İçin Sabit Yangın Pompalarının Kurulumu Standardı)",
            "NFPA 22 — Standard for Water Tanks for Private Fire Protection (Özel Yangın Korunması İçin Su Depoları Standardı)",
            "NFPA 24 — Standard for the Installation of Private Fire Service Mains and Their Appurtenances (Özel Yangın Servis Boru Şebekeleri ve Bağlantı Elemanlarının Kurulumu Standardı)",
            "NFPA 25 — Standard for the Inspection, Testing, and Maintenance of Water-Based Fire Protection Systems (Su Bazlı Yangından Korunma Sistemlerinin Muayene, Test ve Bakım Standardı)",
            "NFPA 72 — National Fire Alarm and Signaling Code (Ulusal Yangın Alarm ve Sinyalizasyon Kodu)",
            "NFPA 750 — Standard on Water Mist Fire Protection Systems (Su Sisi Yangından Korunma Sistemleri Standardı)",
            "NFPA 770 — Standard on Hybrid (Water and Inert Gas) Fire-Extinguishing Systems (Hibrit Su ve İnert Gazlı Yangın Söndürme Sistemleri Standardı)",
            "NFPA 2001 — Standard on Clean Agent Fire Extinguishing Systems (Temiz Gazlı Söndürme Sistemleri Standardı)",
            "NFPA 2010 — Standard for Fixed Aerosol Fire-Extinguishing Systems (Sabit Aerosol Yangın Söndürme Sistemleri Standardı)",
        ]

        # -------------------------------------------------------------------
        # 7.2.1 / 7.2.2 - BYKHY EK-1/B BİNA KULLANIMI / TEHLİKE SINIFI
        # Kaynak: Binaların Yangından Korunması Hakkında Yönetmelik Kılavuzu,
        # Ek-1/B Orta Tehlike Kullanım Alanları.
        # -------------------------------------------------------------------
        _ek1b_basliklari = ["Orta Tehlike -1", "Orta Tehlike -2", "Orta Tehlike -3", "Orta Tehlike -4"]
        # BYKHY Kılavuzu Ek-1/B — Orta Tehlike Kullanım Alanları.
        # Hücre metinleri Bakanlık kılavuzundaki tablo esas alınarak korunmuştur.
        _ek1b_satirlari = [
            {"tur": "Cam ve seramikler", "hücreler": ["", "", "Cam fabrikaları", ""]},
            {"tur": "Kimyasallar", "hücreler": ["Çimento işleri", "Fotoğraf laboratuvarları, fotoğraf film fabrikaları", "Boyama işlemleri, sabun fabrikaları", "Mum ve balmumu fabrikaları, kibrit fabrikaları, boyahaneler"]},
            {"tur": "Mühendislik", "hücreler": ["Metal levha üretimi", "Otomotiv fabrikaları, tamirhaneleri", "Elektronik fabrikaları, buzdolabı ve çamaşır makinesi fabrikaları", ""]},
            {"tur": "Yiyecek ve içecekler", "hücreler": ["Mezbahalar, mandıralar", "Fırınlar, bisküvi, çikolata, şekerleme imalathaneleri, bira fabrikaları", "Hayvan yemi fabrikaları, meyve kurutma, suyu çıkarılmış sebze ve çorba fabrikaları, şeker imalathaneleri, tahıl değirmenleri", "Alkol damıtma"]},
            {"tur": "Çeşitli", "hücreler": ["Hastaneler, oteller, konutlar, lokantalar, kütüphaneler (kitap depoları hariç), okullar, bürolar", "Fizik laboratuvarları, çamaşırhaneler, otoparklar, müzeler", "Radyo ve televizyon yayınevleri, tren istasyonları, tesisat odaları", "Sinemalar, tiyatrolar, konser salonları, tütün fabrikaları"]},
            {"tur": "Kâğıt", "hücreler": ["Cilthaneler, mukavva fabrikaları, kâğıt fabrikaları, baskı işleri ve matbaalar", "", "", "Atık kâğıt işletmeleri"]},
            {"tur": "Lastik ve plastik", "hücreler": ["Kablo fabrikaları, plastik döküm ve plastik eşya (köpük plastik hariç), kauçuk eşya fabrikaları, sentetik lif (akrilik hariç) fabrikaları, vulkanize fabrikaları", "", "Halat fabrikaları", ""]},
            {"tur": "Dükkânlar ve ofisler", "hücreler": ["Bilgisayara veri işleme ofisleri (veri saklama odaları, hariç)", "Büyük mağazalar, alışveriş merkezleri", "", "Sergi salonları"]},
            {"tur": "Tekstiller ve konfeksiyon", "hücreler": ["Deri eşya fabrikaları", "Halı fabrikaları (kauçuk ve köpük plastik hariç), kumaş ve giysi fabrikaları, fiber levha fabrikaları, ayakkabı imalathaneleri, triko (örgü), ev tekstili (bez) fabrikaları, yatak, şilte fabrikaları (köpük plastik hariç), dikim ve dokuma atölyeleri, yün ve yünlü kumaş atölyeleri", "", "Pamuk iplikhanesi, keten ve kenevir hazırlama tesisleri"]},
            {"tur": "Kereste ve tahta", "hücreler": ["Ahşap işleri fabrikaları, mobilya fabrikaları (köpük plastikler hariç), mobilya mağazaları, koltuk, kanepe ve benzeri döşemelerinin (plastik köpük hariç) imalathaneleri", "Odun talaşı fabrikaları, yonga levha fabrikaları, kontrplak levhaları", "", ""]},
        ]

        # Seçim ekranında aynı tablo hücresindeki farklı kullanım alanlarının
        # birbirinden bağımsız seçilebilmesi için hücreler ayrıca tekil kullanım
        # alanlarına ayrılır. Her kayıt yine aynı Ek-1/B hücresine bağlıdır.
        _ek1b_bina_secenekleri = [
            ("Cam fabrikaları", "Cam ve seramikler", 2),
            ("Çimento işleri", "Kimyasallar", 0),
            ("Fotoğraf laboratuvarları, fotoğraf film fabrikaları", "Kimyasallar", 1),
            ("Boyama işlemleri, sabun fabrikaları", "Kimyasallar", 2),
            ("Mum ve balmumu fabrikaları, kibrit fabrikaları, boyahaneler", "Kimyasallar", 3),
            ("Metal levha üretimi", "Mühendislik", 0),
            ("Otomotiv fabrikaları, tamirhaneleri", "Mühendislik", 1),
            ("Elektronik fabrikaları, buzdolabı ve çamaşır makinesi fabrikaları", "Mühendislik", 2),
            ("Mezbahalar, mandıralar", "Yiyecek ve içecekler", 0),
            ("Fırınlar, bisküvi, çikolata, şekerleme imalathaneleri, bira fabrikaları", "Yiyecek ve içecekler", 1),
            ("Hayvan yemi fabrikaları, meyve kurutma, suyu çıkarılmış sebze ve çorba fabrikaları, şeker imalathaneleri, tahıl değirmenleri", "Yiyecek ve içecekler", 2),
            ("Alkol damıtma", "Yiyecek ve içecekler", 3),
            ("Hastaneler", "Çeşitli", 0),
            ("Oteller", "Çeşitli", 0),
            ("Konutlar", "Çeşitli", 0),
            ("Lokantalar", "Çeşitli", 0),
            ("Kütüphaneler (kitap depoları hariç)", "Çeşitli", 0),
            ("Okullar", "Çeşitli", 0),
            ("Bürolar", "Çeşitli", 0),
            ("Fizik laboratuvarları", "Çeşitli", 1),
            ("Çamaşırhaneler", "Çeşitli", 1),
            ("Otoparklar", "Çeşitli", 1),
            ("Müzeler", "Çeşitli", 1),
            ("Radyo ve televizyon yayınevleri", "Çeşitli", 2),
            ("Tren istasyonları", "Çeşitli", 2),
            ("Tesisat odaları", "Çeşitli", 2),
            ("Sinemalar", "Çeşitli", 3),
            ("Tiyatrolar", "Çeşitli", 3),
            ("Konser salonları", "Çeşitli", 3),
            ("Tütün fabrikaları", "Çeşitli", 3),
            ("Cilthaneler", "Kâğıt", 0),
            ("Mukavva fabrikaları", "Kâğıt", 0),
            ("Kâğıt fabrikaları", "Kâğıt", 0),
            ("Baskı işleri ve matbaalar", "Kâğıt", 0),
            ("Atık kâğıt işletmeleri", "Kâğıt", 3),
            ("Kablo fabrikaları", "Lastik ve plastik", 0),
            ("Plastik döküm ve plastik eşya (köpük plastik hariç)", "Lastik ve plastik", 0),
            ("Kauçuk eşya fabrikaları", "Lastik ve plastik", 0),
            ("Sentetik lif (akrilik hariç) fabrikaları", "Lastik ve plastik", 0),
            ("Vulkanize fabrikaları", "Lastik ve plastik", 0),
            ("Halat fabrikaları", "Lastik ve plastik", 2),
            ("Bilgisayara veri işleme ofisleri (veri saklama odaları, hariç)", "Dükkânlar ve ofisler", 0),
            ("Büyük mağazalar", "Dükkânlar ve ofisler", 1),
            ("Alışveriş merkezleri", "Dükkânlar ve ofisler", 1),
            ("Sergi salonları", "Dükkânlar ve ofisler", 3),
            ("Deri eşya fabrikaları", "Tekstiller ve konfeksiyon", 0),
            ("Halı fabrikaları (kauçuk ve köpük plastik hariç)", "Tekstiller ve konfeksiyon", 1),
            ("Kumaş ve giysi fabrikaları", "Tekstiller ve konfeksiyon", 1),
            ("Fiber levha fabrikaları", "Tekstiller ve konfeksiyon", 1),
            ("Ayakkabı imalathaneleri", "Tekstiller ve konfeksiyon", 1),
            ("Triko (örgü)", "Tekstiller ve konfeksiyon", 1),
            ("Ev tekstili (bez) fabrikaları", "Tekstiller ve konfeksiyon", 1),
            ("Yatak, şilte fabrikaları (köpük plastik hariç)", "Tekstiller ve konfeksiyon", 1),
            ("Dikim ve dokuma atölyeleri", "Tekstiller ve konfeksiyon", 1),
            ("Yün ve yünlü kumaş atölyeleri", "Tekstiller ve konfeksiyon", 1),
            ("Pamuk iplikhanesi, keten ve kenevir hazırlama tesisleri", "Tekstiller ve konfeksiyon", 3),
            ("Ahşap işleri fabrikaları", "Kereste ve tahta", 0),
            ("Mobilya fabrikaları (köpük plastikler hariç)", "Kereste ve tahta", 0),
            ("Mobilya mağazaları", "Kereste ve tahta", 0),
            ("Koltuk, kanepe ve benzeri döşemelerinin (plastik köpük hariç) imalathaneleri", "Kereste ve tahta", 0),
            ("Odun talaşı fabrikaları", "Kereste ve tahta", 1),
            ("Yonga levha fabrikaları", "Kereste ve tahta", 1),
            ("Kontrplak levhaları", "Kereste ve tahta", 1),
        ]
        _ek1b_bina_kayitlari = []
        for _etiket, _tur, _kolon in _ek1b_bina_secenekleri:
            _ek1b_bina_kayitlari.append({
                "etiket": _etiket,
                "kullanim_turu": _tur,
                "sinif": _ek1b_basliklari[_kolon],
                "kolon": _kolon,
                "satir": _tur,
            })

        # Eski tekli seçim anahtarını koruyarak yeni çoklu seçim yapısına geçiş.
        _eski_secim = st.session_state.get("yangin_721_ek1b_secim", "")
        st.session_state.setdefault("yangin_721_ek1b_secimler", [])
        st.session_state.setdefault("yangin_721_ek1c_secimler", [])
        if not isinstance(st.session_state.get("yangin_721_ek1b_secimler"), list):
            st.session_state["yangin_721_ek1b_secimler"] = []
        if not isinstance(st.session_state.get("yangin_721_ek1c_secimler"), list):
            st.session_state["yangin_721_ek1c_secimler"] = []
        if not st.session_state["yangin_721_ek1b_secimler"] and _eski_secim:
            _eski_kayit = next((x for x in _ek1b_bina_kayitlari if x["etiket"] == _eski_secim), None)
            if _eski_kayit:
                st.session_state["yangin_721_ek1b_secimler"] = [_eski_kayit["etiket"]]

        # BYKHY Ek-1/C — Yüksek Tehlike Kullanım Alanları.
        # Kaynak kılavuzdaki Ek-1/C tablosu 4 tehlike sütunundan oluşur.
        _ek1c_basliklari = [
            "Yüksek Tehlike -1", "Yüksek Tehlike -2",
            "Yüksek Tehlike -3", "Yüksek Tehlike -4"
        ]
        _ek1c_satirlari = [
            {"hücreler": ["Döşemelik kumaş ve muşamba fabrikaları; kumaş ve muşamba yer döşemeleri imalatı", "Aydınlatma fişeği fabrikaları", "Selüloz nitrat fabrikaları", "Havai fişek fabrikaları"]},
            {"hücreler": ["Boya, renklendirici (ahşap renklendirici ve koruyucuları-pnoteks) ve vernik imalatı", "Plastik köpük ve sünger imalathaneleri, lastik köpük eşyaları", "", ""]},
            {"hücreler": ["Yapay kauçuk, reçine, lamba isi ve terebentin imalatı", "Katran damıtma", "", ""]},
            {"hücreler": ["Talaş fabrikaları; odun yünü imalatı", "Otobüs ambarı, yüklü kamyonlar ve vagonlar; otobüsler, yüksüz kamyonlar ve demiryolu vagonları için depolar", "", ""]},
        ]
        # ÖNEMLİ: Her kullanım alanı, tablodaki GERÇEK satır ve sütun koordinatıyla
        # tanımlanır. Önceki sürümde burada liste sırası "satır" kabul edildiği için
        # örneğin Havai fişek fabrikaları seçildiğinde yanlış satırdaki hücre sarıya
        # boyanabiliyordu.
        _ek1c_bina_kayitlari = [
            {"etiket": "Döşemelik kumaş ve muşamba fabrikaları; kumaş ve muşamba yer döşemeleri imalatı", "sinif": "Yüksek Tehlike -1", "kolon": 0, "satir": 0},
            {"etiket": "Aydınlatma fişeği fabrikaları", "sinif": "Yüksek Tehlike -2", "kolon": 1, "satir": 0},
            {"etiket": "Selüloz nitrat fabrikaları", "sinif": "Yüksek Tehlike -3", "kolon": 2, "satir": 0},
            {"etiket": "Havai fişek fabrikaları", "sinif": "Yüksek Tehlike -4", "kolon": 3, "satir": 0},
            {"etiket": "Boya, renklendirici (ahşap renklendirici ve koruyucuları-pnoteks) ve vernik imalatı", "sinif": "Yüksek Tehlike -1", "kolon": 0, "satir": 1},
            {"etiket": "Plastik köpük ve sünger imalathaneleri, lastik köpük eşyaları", "sinif": "Yüksek Tehlike -2", "kolon": 1, "satir": 1},
            {"etiket": "Yapay kauçuk, reçine, lamba isi ve terebentin imalatı", "sinif": "Yüksek Tehlike -1", "kolon": 0, "satir": 2},
            {"etiket": "Katran damıtma", "sinif": "Yüksek Tehlike -2", "kolon": 1, "satir": 2},
            {"etiket": "Talaş fabrikaları; odun yünü imalatı", "sinif": "Yüksek Tehlike -1", "kolon": 0, "satir": 3},
            {"etiket": "Otobüs ambarı, yüklü kamyonlar ve vagonlar; otobüsler, yüksüz kamyonlar ve demiryolu vagonları için depolar", "sinif": "Yüksek Tehlike -2", "kolon": 1, "satir": 3},
        ]

        st.session_state.setdefault("yangin_721_manuel", False)
        st.session_state.setdefault("yangin_721_manuel_sinif", _ek1b_basliklari[0])
        _tum_yangin_tehlike_siniflari = ["Düşük Tehlike"] + _ek1b_basliklari + _ek1c_basliklari
        st.session_state.setdefault("rapor_bolum_721", True)
        st.session_state.setdefault("rapor_bolum_722", True)

        def _ek1b_secili_kayitlar():
            _secimler = st.session_state.get("yangin_721_ek1b_secimler", []) or []
            return [x for x in _ek1b_bina_kayitlari if x["etiket"] in _secimler]

        def _ek1c_secili_kayitlar():
            _secimler = st.session_state.get("yangin_721_ek1c_secimler", []) or []
            return [x for x in _ek1c_bina_kayitlari if x["etiket"] in _secimler]

        def _yangin_721_secili_kayitlar():
            return _ek1b_secili_kayitlar() + _ek1c_secili_kayitlar()

        def _ek1b_secili_kayit():
            _kayitlar = _ek1b_secili_kayitlar()
            return _kayitlar[0] if _kayitlar else None

        def _yangin_tehlike_sinif_sirasi(_sinif):
            _s = str(_sinif or "")
            if _s.startswith("Düşük"):
                return 0
            if _s.startswith("Orta Tehlike"):
                try: return 1 + int(_s.split("-")[-1].strip())
                except Exception: return 1
            if _s.startswith("Yüksek Tehlike"):
                try: return 5 + int(_s.split("-")[-1].strip())
                except Exception: return 5
            return -1

        def _ek1b_otomatik_sinif(_kayitlar):
            if not _kayitlar:
                return "Belirlenemedi"
            return max((_x["sinif"] if "sinif" in _x else _ek1b_basliklari[int(_x["kolon"])] for _x in _kayitlar), key=_yangin_tehlike_sinif_sirasi)

        def _ek1b_html_tablo(_secili_kayitlar):
            _sec_hucreleri = {(x["satir"], int(x["kolon"])) for x in (_secili_kayitlar or [])}
            _html = '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">'
            _html += '<tr><th style="border:1px solid #777;padding:7px;background:#e6e6e6">KULLANIM TÜRÜ</th>'
            for _h in _ek1b_basliklari:
                _html += f'<th style="border:1px solid #777;padding:7px;background:#e6e6e6">{_h}</th>'
            _html += '</tr>'
            for _satir in _ek1b_satirlari:
                _html += f'<tr><td style="border:1px solid #777;padding:7px;font-weight:700">{_satir["tur"]}</td>'
                for _j, _metin in enumerate(_satir["hücreler"]):
                    _bg = '#fff2cc' if ((_satir["tur"], _j) in _sec_hucreleri) else '#ffffff'
                    _html += f'<td style="border:1px solid #777;padding:7px;background:{_bg};vertical-align:top">{_metin or ""}</td>'
                _html += '</tr>'
            _html += '</table></div>'
            return _html

        def _ek1c_html_tablo(_secili_kayitlar):
            _sec_hucreleri = {(int(x["satir"]), int(x["kolon"])) for x in (_secili_kayitlar or [])}
            _html = '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">'
            _html += '<tr>'
            for _h in _ek1c_basliklari:
                _html += f'<th style="border:1px solid #777;padding:7px;background:#e6e6e6">{_h}</th>'
            _html += '</tr>'
            for _i, _satir in enumerate(_ek1c_satirlari):
                _html += '<tr>'
                for _j, _metin in enumerate(_satir["hücreler"]):
                    _bg = '#fff2cc' if ((_i, _j) in _sec_hucreleri) else '#ffffff'
                    _html += f'<td style="border:1px solid #777;padding:7px;background:{_bg};vertical-align:top">{_metin or ""}</td>'
                _html += '</tr>'
            _html += '</table></div>'
            return _html

        _yangin_71_gruplari = [
            ("7.1.1 YANGIN TESİSATININ AMACI VE KAPSAMI", _yangin_711_maddeleri, "yangin_711_secimler", "rapor_bolum_711"),
            ("7.1.2 YANGIN TESİSATI TASARIM ESASLARI", _yangin_712_maddeleri, "yangin_712_secimler", "rapor_bolum_712"),
        ]

        _yangin_713_standartlari = _yangin_713_ts_standartlari + _yangin_713_nfpa_standartlari
        st.session_state.setdefault("yangin_713_secimler", [True] * len(_yangin_713_standartlari))
        st.session_state.setdefault("rapor_bolum_713", True)
        for _baslik71, _maddeler71, _key71, _rrk71 in _yangin_71_gruplari:
            st.session_state.setdefault(_key71, [True] * len(_maddeler71))
            st.session_state.setdefault(_rrk71, True)

        def _yangin_71_toplu_sec(_key71, _maddeler71, _deger):
            vals = [bool(_deger)] * len(_maddeler71)
            st.session_state[_key71] = vals
            for i, v in enumerate(vals):
                st.session_state[f"{_key71}_item_{i}"] = v

        def _yangin_71_rapor_sync(_rrk, _ui_key):
            st.session_state[_rrk] = bool(st.session_state.get(_ui_key, True))

        _yangin_bolumleri = [
            ("bolum_71", "rapor_bolum_71", "7.1 YANGIN TESİSATI GENEL ESASLARI", [
                ("bolum_711", "rapor_bolum_711", "7.1.1 Yangın Tesisatının Amacı ve Kapsamı"),
                ("bolum_712", "rapor_bolum_712", "7.1.2 Yangın Tesisatı Tasarım Esasları"),
                ("bolum_713", "rapor_bolum_713", "7.1.3 Yangın Tesisatı Standartları"),
            ]),
            ("bolum_72", "rapor_bolum_72", "7.2 YANGIN TEHLİKE SINIFI VE TASARIM KRİTERLERİ", [
                ("bolum_721", "rapor_bolum_721", "7.2.1 Bina Kullanım Amacı"),
                ("bolum_722", "rapor_bolum_722", "7.2.2 Yangın Tehlike Sınıfı"),
                ("bolum_723", "rapor_bolum_723", "7.2.3 Yangın Bölmeleri"),
                ("bolum_724", "rapor_bolum_724", "7.2.4 Tasarım Kriterleri"),
                ("bolum_725", "rapor_bolum_725", "7.2.5 Tasarım Debisi"),
            ]),
            ("bolum_73", "rapor_bolum_73", "7.3 BİNA İÇİ HORTUM SİSTEMİ TASARIMI VE HESAPLAMALARI", []),
            ("bolum_74", "rapor_bolum_74", "7.4 HİDRANT SİSTEMİ TASARIMI VE HESAPLAMALARI", []),
            ("bolum_75", "rapor_bolum_75", "7.5 SPRİNKLER (YAĞMURLAMA) SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_76", "rapor_bolum_76", "7.6 GAZLI SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_77", "rapor_bolum_77", "7.7 KÖPÜKLÜ SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_78", "rapor_bolum_78", "7.8 DAVLUMBAZ SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_79", "rapor_bolum_79", "7.9 MERDİVEN BASINÇLANDIRMA SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_710", "rapor_bolum_710", "7.10 ASANSÖR BASINÇLANDIRMA SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_7110", "rapor_bolum_7110", "7.11 DUMAN TAHLİYE SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_7120", "rapor_bolum_7120", "7.12 YANGIN SUYU DEPOLAMA SİSTEMİ TASARIM VE HESAPLAMALARI", []),
            ("bolum_7130", "rapor_bolum_7130", "7.13 YANGIN POMPA GRUBU TASARIM VE HESAPLAMALARI", []),
            ("bolum_7140", "rapor_bolum_7140", "7.14 YANGIN TESİSATI HİDROLİK HESAPLARI", []),
            ("bolum_7150", "rapor_bolum_7150", "7.15 YANGIN TESİSATI SONUÇ TABLOSU", []),
        ]

        for _key, _rapor_key, _baslik, _altlar in _yangin_bolumleri:
            st.session_state.setdefault(_key, True)
            st.session_state.setdefault(_rapor_key, True)

            if not st.session_state.get(_key, True):
                continue

            if _baslik.startswith(("7.3 ", "7.4 ", "7.5 ", "7.6 ", "7.7 ", "7.8 ", "7.9 ", "7.10 ", "7.11 ", "7.12 ", "7.13 ", "7.14 ", "7.15 ")):
                st.markdown(
                    f'<div style="font-size:23px; font-weight:800; font-style:italic; '
                    f'color:#1F4E79; margin:16px 0 10px 0; text-transform:uppercase;">'
                    f'{_baslik.upper()}</div>',
                    unsafe_allow_html=True,
                )
                st.checkbox(
                    "Bu bölüm rapora eklensin",
                    key=f"{_rapor_key}_ui",
                    on_change=lambda _rk=_rapor_key, _uk=f"{_rapor_key}_ui": st.session_state.__setitem__(
                        _rk, bool(st.session_state.get(_uk, True))
                    ),
                )
            else:
                st.subheader(_baslik)

            if _baslik.startswith("7.1 "):
                # 7.1 alt maddeleri tek bir hiyerarşi içinde açılır/kapanır.
                # 7.1.1 ve 7.1.2 madde seçimleri; 7.1.3 ise standart kütüphanesi içerir.
                st.caption("7.1 alt maddeleri ayrı ayrı açılıp kapatılabilir. İşaretli maddeler rapora aktarılır.")

                for _baslik71, _maddeler71, _key71, _rrk71 in _yangin_71_gruplari:
                    with st.expander(_baslik71, expanded=False):
                        c1, c2, c3 = st.columns([1, 1, 2])
                        with c1:
                            # Bu butonların anahtarları proje/session verisine kaydedilmez.
                            # Eski sürümlerden kalmış aynı anahtarlar varsa Streamlit,
                            # st.button oluşturulurken StreamlitValueAssignmentNotAllowedError
                            # verebiliyor. Bu nedenle buton anahtarlarını her çalıştırmada temizliyoruz.
                            # Bu iki toplu seçim düğmesi proje verisine ait değildir.
                            # Eski kayıtlı projelerde widget anahtarları session_state içine
                            # taşınabildiği için explicit key kullanmıyoruz; Streamlit burada
                            # düğmeleri delta konumlarına göre benzersiz olarak tanımlar.
                            _71_tum_sec_key = f"yangin_71_{_key71}_tum_sec_v3"
                            _71_tum_kaldir_key = f"yangin_71_{_key71}_tum_kaldir_v3"
                            if st.button(
                                f"✓ TÜMÜNÜ SEÇ — {_key71}",
                                key=_71_tum_sec_key,
                                use_container_width=True,
                            ):
                                _yangin_71_toplu_sec(_key71, _maddeler71, True)
                                st.rerun()
                        with c2:
                            if st.button(
                                f"✕ TÜMÜNÜ KALDIR — {_key71}",
                                key=_71_tum_kaldir_key,
                                use_container_width=True,
                            ):
                                _yangin_71_toplu_sec(_key71, _maddeler71, False)
                                st.rerun()
                        with c3:
                            _ui_rapor_key = f"yangin_71_rapor_ui_{_rrk71}"
                            st.session_state.setdefault(_ui_rapor_key, bool(st.session_state.get(_rrk71, True)))
                            st.checkbox(
                                "Bu alt bölüm rapora eklensin",
                                key=_ui_rapor_key,
                                on_change=_yangin_71_rapor_sync,
                                args=(_rrk71, _ui_rapor_key),
                            )
                        vals = list(st.session_state.get(_key71, [True] * len(_maddeler71)))
                        for i, madde in enumerate(_maddeler71):
                            item_key = f"{_key71}_item_{i}"
                            st.session_state.setdefault(item_key, bool(vals[i]))
                            st.checkbox(f"• {madde}", key=item_key)
                            vals[i] = bool(st.session_state.get(item_key, vals[i]))
                        st.session_state[_key71] = vals

                # 7.1.3 standart kütüphanesi de aynı seviyede açılır/kapanır.
                with st.expander("7.1.3 YANGIN TESİSATI STANDARTLARI", expanded=False):
                    _ui_713_rapor_key = "yangin_713_rapor_ui"
                    st.session_state.setdefault(_ui_713_rapor_key, bool(st.session_state.get("rapor_bolum_713", True)))
                    st.checkbox(
                        "7.1.3 bölümü rapora eklensin",
                        key=_ui_713_rapor_key,
                        on_change=_yangin_71_rapor_sync,
                        args=("rapor_bolum_713", _ui_713_rapor_key),
                    )
                    _ts_tab, _nfpa_tab = st.tabs(["TS / TS EN STANDARTLARI", "NFPA STANDARTLARI"])
                    for _tab, _prefix, _liste in [
                        (_ts_tab, "ts", _yangin_713_ts_standartlari),
                        (_nfpa_tab, "nfpa", _yangin_713_nfpa_standartlari),
                    ]:
                        with _tab:
                            _b1, _b2 = st.columns(2)
                            _713_tum_sec_key = f"yangin_713_{_prefix}_tum_sec_v3"
                            _713_tum_kaldir_key = f"yangin_713_{_prefix}_tum_kaldir_v3"
                            if _b1.button("✓ TÜMÜNÜ SEÇ", key=_713_tum_sec_key, use_container_width=True):
                                for _i in range(len(_liste)):
                                    st.session_state[f"yangin_713_{_prefix}_{_i}"] = True
                                st.rerun()
                            if _b2.button("✕ TÜMÜNÜ KALDIR", key=_713_tum_kaldir_key, use_container_width=True):
                                for _i in range(len(_liste)):
                                    st.session_state[f"yangin_713_{_prefix}_{_i}"] = False
                                st.rerun()
                            for _i, _std in enumerate(_liste):
                                st.session_state.setdefault(f"yangin_713_{_prefix}_{_i}", True)
                                st.checkbox(_std, key=f"yangin_713_{_prefix}_{_i}")
                    _ts_vals = [bool(st.session_state.get(f"yangin_713_ts_{_i}", True)) for _i in range(len(_yangin_713_ts_standartlari))]
                    _nfpa_vals = [bool(st.session_state.get(f"yangin_713_nfpa_{_i}", True)) for _i in range(len(_yangin_713_nfpa_standartlari))]
                    st.session_state["yangin_713_secimler"] = _ts_vals + _nfpa_vals

            elif _baslik.startswith("7.2"):
                # 7.2.1 - Ek-1/B üzerinden bina/kullanım alanı seçimi ve otomatik tehlike sınıfı
                _alt721 = next((x for x in _altlar if x[0] == "bolum_721"), None)
                _alt722 = next((x for x in _altlar if x[0] == "bolum_722"), None)
                _k721, _r721, _b721 = _alt721 if _alt721 else ("bolum_721", "rapor_bolum_721", "7.2.1 Bina Kullanım Amacı")
                _k722, _r722, _b722 = _alt722 if _alt722 else ("bolum_722", "rapor_bolum_722", "7.2.2 Yangın Tehlike Sınıfı")
                st.session_state.setdefault(_k721, True)
                st.session_state.setdefault(_k722, True)
                st.session_state.setdefault(_r721, True)
                st.session_state.setdefault(_r722, True)

                # ------------------------------------------------------------------
                # GENEL BİNA BİLGİLERİ
                # Bu bilgiler 7.3 ve sonraki yangın tesisatı modüllerinin ortak
                # veri kaynağıdır. Şimdilik yalnızca program ekranında kullanılır;
                # rapor üretimine aktarılmaz.
                # ------------------------------------------------------------------
                with st.expander("GENEL BİNA BİLGİLERİ", expanded=True):
                    st.markdown(
                        '<div style="font-size:16px;font-weight:700;font-style:italic;">'
                        'Yangın Yönetmeliğine Göre Ortak Bina Verileri</div>', 
                        unsafe_allow_html=True,
                    )
                    st.caption(
                        "Bu bilgiler, aşağıdaki yangın tesisatı bölümlerinde yapılacak "
                        "ön değerlendirme ve hesaplamalarda ortak veri olarak kullanılacaktır. "
                        "Şimdilik rapora aktarılmaz."
                    )

                    with st.expander("ⓘ BİNA YÜKSEKLİĞİ VE YAPI YÜKSEKLİĞİ TANIMI / ŞEKLİ", expanded=False):
                        st.markdown(
                            "**Bina yüksekliği:** Binanın kot aldığı noktadan saçak seviyesine kadar olan mesafedir."
                        )
                        st.markdown(
                            "**Yapı yüksekliği:** Bodrum katlar, asma katlar ve çatı arası piyesler dâhil olmak üzere, yapının inşa edilen bütün katlarının toplam yüksekliğidir."
                        )
                        # Şema dosyası GitHub/Streamlit ortamında isim veya uzantı farkı
                        # nedeniyle bulunamazsa, aynı klasördeki uygun PNG dosyasını da ara.
                        _sema_kok = Path(__file__).resolve().parent.parent
                        _sema_adaylari = [
                            _sema_kok / "bina_yuksekligi_yapi_yuksekligi_sema_opt.png",
                            *_sema_kok.glob("bina_yuksekligi_yapi_yuksekligi_sema*.png"),
                        ]
                        _sema_yolu = next((p for p in _sema_adaylari if p.is_file()), None)
                        _sema_verisi = _sema_yolu.read_bytes() if _sema_yolu else None
                        if _sema_verisi:
                            st.image(_sema_verisi, caption="Bina yüksekliği ve yapı yüksekliği — şematik gösterim", use_container_width=True)
                        else:
                            st.warning("Bina yüksekliği / yapı yüksekliği şeması bulunamadı. PNG dosyasının app.py ile aynı kökte olduğundan emin olun.")
                        st.caption(
                            "Not: Şema açıklayıcı amaçlıdır. Projede ölçü alınırken yürürlükteki mevzuat tanımları ve ilgili kotlar esas alınmalıdır."
                        )

                    _b1, _b2, _b3 = st.columns(3)
                    with _b1:
                        st.number_input(
                            "Toplam yapı / kapalı kullanım alanı (m²)",
                            min_value=0.0, step=10.0,
                            key="yangin_genel_toplam_alan_m2",
                            help="Yangın tesisatı değerlendirmelerinde kullanılacak toplam yapı / kapalı kullanım alanını m² olarak giriniz.",
                        )
                        st.number_input(
                            "Kat sayısı", min_value=0, step=1,
                            key="yangin_genel_kat_sayisi",
                            help="Binanın toplam kat sayısını giriniz. İlgili yangın güvenliği değerlendirmelerinde kullanılacaktır.",
                        )
                        st.number_input(
                            "Bodrum kat sayısı", min_value=0, step=1,
                            key="yangin_genel_bodrum_kat_sayisi",
                            help="Binada bulunan bodrum kat adedini giriniz.",
                        )
                    with _b2:
                        st.number_input(
                            "Bina yüksekliği (m)", min_value=0.0, step=0.10,
                            key="yangin_genel_bina_yuksekligi_m",
                            help="BYKHY tanımına göre binanın kot aldığı noktadan saçak seviyesine kadar olan mesafedir. Değeri metre olarak giriniz. ⓘ Ayrıntılı şekil aşağıdaki bilgi panelindedir.",
                        )
                        st.number_input(
                            "Yapı yüksekliği (m)", min_value=0.0, step=0.10,
                            key="yangin_genel_yapi_yuksekligi_m",
                            help="BYKHY tanımına göre bodrum katlar, asma katlar ve çatı arası piyesler dâhil yapının inşa edilen bütün katlarının toplam yüksekliğidir. Değeri metre olarak giriniz. ⓘ Ayrıntılı şekil aşağıdaki bilgi panelindedir.",
                        )
                        st.number_input(
                            "Merdiven kovası yüksekliği (m)", min_value=0.0, step=0.10,
                            key="yangin_genel_merdiven_kovasi_yuksekligi_m",
                            help="Merdiven kovası / merdiven basınçlandırması değerlendirmesinde kullanılacak yaklaşık yüksekliği metre olarak giriniz.",
                        )
                    with _b3:
                        st.number_input(
                            "Toplam kişi sayısı", min_value=0, step=1,
                            key="yangin_genel_kisi_sayisi",
                            help="Binada bulunan veya hesapta esas alınacak toplam kişi sayısını giriniz.",
                        )
                        st.number_input(
                            "Otopark sayısı / araç kapasitesi", min_value=0, step=1,
                            key="yangin_genel_otopark_arac_sayisi",
                            help="Kapalı otoparkta bulunabilecek toplam araç kapasitesini giriniz.",
                        )
                        st.number_input(
                            "Kapalı otopark alanı (m²)", min_value=0.0, step=10.0,
                            key="yangin_genel_kapali_otopark_alan_m2",
                            help="Kapalı otoparkın yangın güvenliği değerlendirmesinde kullanılacak toplam alanını m² olarak giriniz.",
                        )

                    _b4, _b5, _b6 = st.columns(3)
                    with _b4:
                        st.number_input(
                            "Yatak sayısı (varsa)", min_value=0, step=1,
                            key="yangin_genel_yatak_sayisi",
                            help="Hastane, otel, yurt vb. yapılarda varsa toplam yatak sayısını giriniz.",
                        )
                    with _b5:
                        st.number_input(
                            "İmar planlama / yerleşim alanı (m²)", min_value=0.0, step=10.0,
                            key="yangin_genel_imar_alani_m2",
                            help="Proje kapsamında kullanılacak yerleşim / imar alanını m² olarak giriniz.",
                        )
                    with _b6:
                        st.checkbox(
                            "Acil durum asansörü bulunuyor",
                            key="yangin_genel_acil_durum_asansoru",
                            help="Binada acil durum asansörü bulunup bulunmadığını belirtiniz.",
                        )

                _genel_alan = float(st.session_state.get("yangin_genel_toplam_alan_m2", 0.0) or 0.0)
                _kat = int(st.session_state.get("yangin_genel_kat_sayisi", 0) or 0)
                _bodrum = int(st.session_state.get("yangin_genel_bodrum_kat_sayisi", 0) or 0)
                _bina_h = float(st.session_state.get("yangin_genel_bina_yuksekligi_m", 0.0) or 0.0)
                _yapi_h = float(st.session_state.get("yangin_genel_yapi_yuksekligi_m", 0.0) or 0.0)
                _merdiven_h = float(st.session_state.get("yangin_genel_merdiven_kovasi_yuksekligi_m", 0.0) or 0.0)
                _otopark_arac = int(st.session_state.get("yangin_genel_otopark_arac_sayisi", 0) or 0)
                _otopark_alan = float(st.session_state.get("yangin_genel_kapali_otopark_alan_m2", 0.0) or 0.0)
                _yatak = int(st.session_state.get("yangin_genel_yatak_sayisi", 0) or 0)
                _imar_alan = float(st.session_state.get("yangin_genel_imar_alani_m2", 0.0) or 0.0)
                _acil_asansor = bool(st.session_state.get("yangin_genel_acil_durum_asansoru", False))

                if st.session_state.get(_k721, True):
                    with st.expander("7.2.1 BİNA KULLANIM AMACI", expanded=True):
                        _c721a, _c721b = st.columns([3, 1])
                        with _c721a:
                            st.multiselect(
                                "BYKHY Ek-1/B'ye göre bina / kullanım alanı (birden fazla seçilebilir)",
                                [x["etiket"] for x in _ek1b_bina_kayitlari],
                                key="yangin_721_ek1b_secimler",
                                placeholder="Ek-1/B kullanım alanlarını seçiniz",
                            )
                            st.multiselect(
                                "BYKHY Ek-1/C'ye göre yüksek tehlike kullanım alanı (birden fazla seçilebilir)",
                                [x["etiket"] for x in _ek1c_bina_kayitlari],
                                key="yangin_721_ek1c_secimler",
                                placeholder="Ek-1/C kullanım alanlarını seçiniz",
                            )
                        with _c721b:
                            _ui_721_rapor_key = "yangin_721_rapor_ui"
                            st.session_state.setdefault(_ui_721_rapor_key, bool(st.session_state.get("rapor_bolum_721", True)))
                            st.checkbox(
                                "7.2.1 rapora eklensin",
                                key=_ui_721_rapor_key,
                                on_change=lambda: st.session_state.__setitem__(
                                    "rapor_bolum_721", bool(st.session_state.get("yangin_721_rapor_ui", True))
                                ),
                            )

                        _secili_kayitlar = _yangin_721_secili_kayitlar()
                        _otomatik = _ek1b_otomatik_sinif(_secili_kayitlar)
                        # Ek-1/B tablosu ekranı gereksiz yere uzatmasın; üst bölümlerdeki
                        # açılır/kapanır yapı ile aynı mantıkta, varsayılan olarak kapalı gösterilir.
                        with st.expander("📋 BYKHY EK-1/B TABLOSUNU GÖSTER / GİZLE", expanded=False):
                            st.markdown("**BYKHY Ek-1/B — Orta Tehlike Kullanım Alanları**")
                            st.markdown(_ek1b_html_tablo(_ek1b_secili_kayitlar()), unsafe_allow_html=True)
                            st.caption("Tablo, Bakanlık BYKHY Kılavuzu Ek-1/B'deki kullanım türleri ve Orta Tehlike sınıfları esas alınarak gösterilmektedir.")
                        _secili_c = _ek1c_secili_kayitlar()
                        with st.expander("📋 BYKHY EK-1/C TABLOSUNU GÖSTER / GİZLE", expanded=False):
                            st.markdown("**BYKHY Ek-1/C — Yüksek Tehlike Kullanım Alanları**")
                            st.markdown(_ek1c_html_tablo(_secili_c), unsafe_allow_html=True)
                            st.caption("Kaynak: Binaların Yangından Korunması Hakkında Yönetmelik Kılavuzu, Ek-1/C — Yüksek Tehlike Kullanım Alanları.")
                        if _secili_kayitlar:
                            st.success(f"Seçilen kullanım alanı sayısı: **{len(_secili_kayitlar)}** — en yüksek otomatik yangın tehlike sınıfı: **{_otomatik}**")

                        st.checkbox(
                            "Yangın tehlike sınıfına manuel müdahale et",
                            key="yangin_721_manuel",
                        )
                        if st.session_state.get("yangin_721_manuel", False):
                            st.selectbox(
                                "Manuel yangın tehlike sınıfı",
                                _tum_yangin_tehlike_siniflari,
                                key="yangin_721_manuel_sinif",
                            )

                if st.session_state.get(_k722, True):
                    with st.expander("7.2.2 YANGIN TEHLİKE SINIFI", expanded=True):
                        _secili_kayitlar = _yangin_721_secili_kayitlar()
                        _otomatik = _ek1b_otomatik_sinif(_secili_kayitlar)
                        _manuel = bool(st.session_state.get("yangin_721_manuel", False))
                        _etkin = st.session_state.get("yangin_721_manuel_sinif", _otomatik) if _manuel else _otomatik
                        _c722a, _c722b = st.columns([3, 1])
                        with _c722a:
                            st.metric("Uygulanacak Yangın Tehlike Sınıfı", _etkin)
                        with _c722b:
                            _ui_722_rapor_key = "yangin_722_rapor_ui"
                            st.session_state.setdefault(_ui_722_rapor_key, bool(st.session_state.get("rapor_bolum_722", True)))
                            st.checkbox(
                                "7.2.2 rapora eklensin",
                                key=_ui_722_rapor_key,
                                on_change=lambda: st.session_state.__setitem__(
                                    "rapor_bolum_722", bool(st.session_state.get("yangin_722_rapor_ui", True))
                                ),
                            )
                        st.write(f"**Ek-1/B + Ek-1/C otomatik sonucu (en yüksek seçilen sınıf):** {_otomatik}")

                        if _secili_kayitlar:
                            st.markdown("**Çoklu seçilen kullanım alanları ve yangın tehlike sınıfları:**")
                            for _kayit in _secili_kayitlar:
                                st.markdown(f"- {_kayit['etiket']} — **{_kayit['sinif']}**")

                        st.markdown(
                            f'<div style="background-color:#FFF2CC; border:1px solid #D6B656; '
                            f'padding:8px 12px; border-radius:4px; margin-top:8px;">'
                            f'<b>Seçilen yangın tehlike sınıfı:</b> {_etkin}'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        st.session_state["yangin_722_etkin_sinif"] = _etkin
                        st.session_state["yangin_722_otomatik_sinif"] = _otomatik
                        st.session_state["yangin_722_secim_kaynagi"] = "Manuel" if _manuel else "BYKHY Ek-1/B"
                # --------------------------------------------------------------
                # GENEL YANGIN SİSTEMİ GEREKSİNİMLERİ — SADECE PROGRAM EKRANI
                # Yönetmelik ön değerlendirmesi. Rapor veri akışına aktarılmaz.
                # --------------------------------------------------------------
                _secili_kullanim = [x["etiket"] for x in _yangin_721_secili_kayitlar()]
                _kullanim_metni = " ".join(_secili_kullanim).lower()
                _etkin_sinif = st.session_state.get("yangin_722_etkin_sinif", "Belirlenemedi")
                _yuksek_bina = _yapi_h > 30.50
                _konut = "konut" in _kullanim_metni

                # Yangın dolabı — Madde 94/1-b-1
                _dolap_gerekce = []
                if _yuksek_bina:
                    _dolap_gerekce.append("Yapı yüksekliği 30,50 m eşiğini aşıyor")
                if _genel_alan > 2000:
                    _dolap_gerekce.append("Toplam kapalı kullanım alanı 2000 m² eşiğini aşıyor")
                if any(k in _kullanim_metni for k in ["imalathane", "atölye", "depo", "otel", "motel", "sağlık", "toplanma", "eğitim"] ) and _genel_alan > 1000:
                    _dolap_gerekce.append("Kullanım türü ve 1000 m² eşiği birlikte sağlanıyor")
                _dolap = bool(_dolap_gerekce)

                # Sprinkler — Madde 96/2'deki temel eşikler.
                _spr_gerekce = []
                if (_yapi_h > 30.50 and not _konut):
                    _spr_gerekce.append("Konut dışı bina ve yapı yüksekliği 30,50 m üzeri")
                if (_konut and _yapi_h > 51.50):
                    _spr_gerekce.append("Konut ve yapı yüksekliği 51,50 m üzeri")
                if _otopark_alan > 600 or _otopark_arac > 10:
                    _spr_gerekce.append("Kapalı otopark için yönetmelik eşiği sağlanıyor")
                if _yatak > 200 or ("otel" in _kullanim_metni and _yatak > 100):
                    _spr_gerekce.append("Yataklı kullanım için yönetmelik eşiği sağlanıyor")
                if (any(k in _kullanim_metni for k in ["mağaza", "alışveriş", "ticaret", "eğlence"]) and _genel_alan > 2000):
                    _spr_gerekce.append("Katlı mağaza/alışveriş/ticaret/eğlence alanı 2000 m² üzeri")
                if any(k in _kullanim_metni for k in ["parlayıcı", "kolay alevlen", "yanıcı madde"]) and _genel_alan > 1000:
                    _spr_gerekce.append("Kolay alevlenici/parlayıcı madde kullanımında 1000 m² eşiği")
                _sprinkler = bool(_spr_gerekce)

                # Merdiven basınçlandırma — Madde 89. Merdiven yüksekliği girilmişse
                # onu, girilmemişse yapı yüksekliğini yalnızca ön değerlendirme için kullan.
                _merdiven_esas_h = _merdiven_h if _merdiven_h > 0 else _yapi_h
                _bas_gerekce = []
                if not _konut and _merdiven_esas_h > 30.50:
                    _bas_gerekce.append("Konut dışı bina ve merdiven kovası yüksekliği 30,50 m üzeri")
                if _bodrum > 4:
                    _bas_gerekce.append("Bodrum kat sayısı 4'ten fazla")
                if _konut and _yapi_h > 51.50:
                    _bas_gerekce.append("Konut ve yapı yüksekliği 51,50 m üzeri")
                _merdiven_bas = bool(_bas_gerekce)

                # Acil durum asansörü kuyusu basınçlandırması — Madde 89/4.
                _asansor_bas = _acil_asansor

                # Duman tahliye — Madde 86/3.
                _duman_gerekce = []
                if _otopark_alan > 2000:
                    _duman_gerekce.append("Kapalı otopark alanı 2000 m² üzeri")
                _duman = bool(_duman_gerekce)

                # Hidrant — Madde 95/7. Alan bilgisi yoksa karar verilemez.
                _hidrant = _imar_alan > 5000
                _hidrant_belirsiz = _imar_alan <= 0

                def _gereksinim_satiri(_baslik, _durum, _aciklama):
                    _bg = "#C6EFCE" if _durum == "GEREKLİ" else ("#FFF2CC" if _durum == "EK BİLGİ GEREKLİ" else "#E7E6E6")
                    return (
                        f'<div style="background:{_bg};border:1px solid #999;padding:9px 12px;'
                        f'margin:4px 0;border-radius:5px;">'
                        f'<b>{_baslik}</b> &nbsp; <b>{_durum}</b><br>'
                        f'<span style="font-size:13px;">{_aciklama}</span></div>'
                    )

                st.markdown(
                    '<div style="font-size:18px;font-weight:700;font-style:italic;">'
                    'YANGIN YÖNETMELİĞİNE GÖRE SİSTEM GEREKSİNİMLERİ ÖN DEĞERLENDİRMESİ'
                    '</div>', unsafe_allow_html=True
                )
                st.caption(
                    "Bu tablo yalnızca program içi ön değerlendirmedir. Kesin proje kararı; "
                    "ilgili tüm kullanım alanları, yangın bölmeleri, mimari veriler ve hidrolik "
                    "hesaplar tamamlandıktan sonra verilecektir."
                )

                st.markdown(_gereksinim_satiri(
                    "YANGIN DOLABI", "GEREKLİ" if _dolap else "GEREKMEYEBİLİR",
                    "; ".join(_dolap_gerekce) if _dolap_gerekce else "Mevcut genel bina verileriyle Madde 94 kapsamındaki zorunluluk eşiği görülmedi."
                ), unsafe_allow_html=True)
                st.markdown(_gereksinim_satiri(
                    "SPRİNKLER (YAĞMURLAMA)", "GEREKLİ" if _sprinkler else "GEREKMEYEBİLİR",
                    "; ".join(_spr_gerekce) if _spr_gerekce else "Madde 96 kapsamındaki girilmiş eşiklerden biri şu an sağlanmış görünmüyor."
                ), unsafe_allow_html=True)
                st.markdown(_gereksinim_satiri(
                    "HİDRANT", "EK BİLGİ GEREKLİ" if _hidrant_belirsiz else ("GEREKLİ" if _hidrant else "GEREKMEYEBİLİR"),
                    "İmar planlama / yerleşim alanı bilgisi girilmedi. 5000 m² eşiği kontrol edilemiyor." if _hidrant_belirsiz else ("İmar planlama alanı 5000 m² üzeri." if _hidrant else "İmar planlama alanı 5000 m² eşiğini aşmıyor."),
                ), unsafe_allow_html=True)
                st.markdown(_gereksinim_satiri(
                    "MERDİVEN BASINÇLANDIRMA", "GEREKLİ" if _merdiven_bas else "GEREKMEYEBİLİR",
                    "; ".join(_bas_gerekce) if _bas_gerekce else "Madde 89 kapsamındaki girilmiş eşiklerden biri şu an sağlanmış görünmüyor."
                ), unsafe_allow_html=True)
                st.markdown(_gereksinim_satiri(
                    "ACİL DURUM ASANSÖRÜ KUYUSU BASINÇLANDIRMA", "GEREKLİ" if _asansor_bas else "GEREKMEYEBİLİR",
                    "Acil durum asansörü bulunduğu işaretlendi." if _asansor_bas else "Acil durum asansörü bulunuyor olarak işaretlenmedi."
                ), unsafe_allow_html=True)
                st.markdown(_gereksinim_satiri(
                    "DUMAN TAHLİYE", "GEREKLİ" if _duman else "GEREKMEYEBİLİR",
                    "; ".join(_duman_gerekce) if _duman_gerekce else "Madde 86/3 kapsamındaki 2000 m² kapalı otopark eşiği girilen verilerde sağlanmıyor."
                ), unsafe_allow_html=True)
                st.markdown(_gereksinim_satiri(
                    "YANGIN SUYU DEPOSU / POMPA", "DEĞERLENDİRİLECEK",
                    "Sulu söndürme sistemlerinin kesin debi, basınç ve süre değerleri belirlendiğinde Madde 91-92 kapsamında ayrıca hesaplanacaktır."
                ), unsafe_allow_html=True)
                st.markdown(_gereksinim_satiri(
                    "GAZLI / KÖPÜKLÜ / DAVLUMBAZ SÖNDÜRME", "RİSKE GÖRE DEĞERLENDİRİLECEK",
                    "Bu sistemler yalnızca genel bina alanı ve yükseklikten otomatik olarak kesinleştirilmeyecek; ilgili kullanım/risk bilgileri ilgili bölümlerde ayrıca sorgulanacaktır."
                ), unsafe_allow_html=True)

            elif _baslik.startswith("7.3 "):
                # ------------------------------------------------------------------
                # 7.3 YANGIN DOLABI SİSTEMİ
                # 7.2 GENEL BİNA BİLGİLERİ ve seçilen etkin yangın tehlike sınıfı
                # bu bölümün ortak veri kaynağıdır.
                # ------------------------------------------------------------------
                with st.expander("7.3 BİNA İÇİ HORTUM SİSTEMİ TASARIMI VE HESAPLAMALARI", expanded=True):
                    st.markdown(
                        '<div style="font-size:22px; font-weight:800; font-style:italic; color:#1F4E79;">'
                        "BİNA İÇİ HORTUM SİSTEMİ TASARIMI VE HESAPLAMALARI</div>",
                        unsafe_allow_html=True,
                    )
                    _ui_73_rapor_key = "yangin_73_rapor_ui"
                    st.session_state.setdefault(_ui_73_rapor_key, bool(st.session_state.get("rapor_bolum_73", True)))
                    st.checkbox(
                        "7.3 rapora eklensin",
                        key=_ui_73_rapor_key,
                        on_change=lambda: st.session_state.__setitem__(
                            "rapor_bolum_73", bool(st.session_state.get("yangin_73_rapor_ui", True))
                        ),
                    )

                    st.markdown(
                        '<div style="font-size:18px; font-weight:800; font-style:italic; color:#1F4E79; margin-top:12px;">'
                        "7.3.1 BİNA İÇİ HORTUM SİSTEMİ YÖNETMELİK ESASLARI VE TASARIM KRİTERLERİ</div>",
                        unsafe_allow_html=True,
                    )
                    _yd_esaslari = [
                        ("1", "Yangın dolabı yapılması", "Yüksek binalarda; toplam kapalı kullanım alanı 1000 m²’den büyük imalathane, atölye, depo, otel, motel, sağlık, toplanma amaçlı ve eğitim binalarında ve kapalı kullanım alanı 2000 m²’den büyük binalarda yangın dolabı yapılması zorunludur."),
                        ("2", "Yangın dolaplarının yerleşimi", "Yangın dolapları her katta ve yangın duvarları ile ayrılmış her bölümde, aralarındaki uzaklık 30 m’yi geçmeyecek şekilde düzenlenir. Yağmurlama sistemi ve katlarda itfaiye su alma ağzı bulunması hâlinde bu mesafe 45 m’ye kadar çıkarılabilir."),
                        ("3", "Yerleşim yeri ve erişilebilirlik", "Dolapların mümkün olduğunca koridor çıkışları ve merdiven sahanlıkları yakınına, kolay görülebilecek ve acil durumda kolay erişilebilecek yerlere yerleştirilmesi esastır."),
                        ("4", "Dolap ve kabin özellikleri", "Dolap veya kabin, gerekli yangın söndürme cihazlarının yerleştirilmesine izin verecek büyüklükte olmalı; hortum ve cihazların yangın sırasında kullanımını zorlaştırmamalı ve yalnızca yangın söndürme amacıyla kullanılmalıdır."),
                        ("5", "Yuvarlak yarı-sert hortumlu dolaplar", "Hortum serme ve bağlama konusunda eğitimli personel veya itfaiye görevlisi bulunmayan yapılarda TS EN 671-1’e uygun yuvarlak yarı-sert hortumlu yangın dolapları kullanılır. Hortum TS EN 694’e uygun, çapı 25 mm ve uzunluğu en fazla 30 m olmalıdır."),
                        ("6", "Yuvarlak yarı-sert hortumlu dolaplarda debi ve basınç", "İçinde itfaiye su alma ağzı bulunmayan yuvarlak yarı-sert hortumlu yangın dolaplarında tasarım debisi 100 L/dak ve lüle girişindeki tasarım basıncı 400 kPa olmalıdır. Lüle giriş basıncı 700 kPa’ı aşarsa basınç düşürücü kullanılır."),
                        ("7", "Yassı hortumlu yangın dolapları", "Yetişmiş yangın söndürme görevlisi bulundurulması gereken yapılarda TS EN 671-2’ye uygun yassı hortumlu dolaplar kullanılabilir. Hortum anma çapı 50 mm’yi, uzunluğu 20 m’yi geçmemelidir. Tasarım debisi 400 L/dak ve lüle girişindeki basınç 600 kPa olmalıdır."),
                        ("8", "Yassı hortumlu dolaplarda basınç kontrolü", "Yassı hortumlu sistemlerde lüle girişindeki basıncın 900 kPa’ı aşması hâlinde basınç düşürücü kullanılır."),
                        ("9", "Periyodik bakım", "Yangın dolapları ve hortum makara sistemlerinin TS EN 671-3’te belirtilen periyodik bakımları bina sahibi, yönetici veya sorumlu bina yetkilisi tarafından yaptırılmalıdır."),
                        ("10", "Kat bağlantı vanaları ve itfaiye su alma ağızları", "İtfaiye su alma hattı kapsamında, yüksek binalar ile kat alanı 1.000 m²’den fazla olan alışveriş merkezleri, otoparklar ve benzeri yerlerde ıslak veya kuru sabit boru sistemi üzerinde itfaiye personeli ve eğitilmiş personelin kullanımına imkân sağlayan bağlantı ağızları bırakılır. Bu bağlantı ağızları kaçış merdiveni veya yangın güvenlik holü gibi korunmuş mekânlarda düzenlenir. Bir boyutu 60 m’yi geçen katlarda yangın dolabı ve itfaiye su alma ağzı yapılması gerekir. (BYKHY Madde 94, 1/a/1)"),
                        ("11", "İtfaiye su alma ağzına erişim mesafesi", "Herhangi bir noktadan itfaiye su alma ağzına olan mesafe 60 m’den fazla olamaz. Sabit boru tesisatı üzerindeki bütün hortum bağlantıları itfaiyenin kullandığı normlarda Storz tip 50 mm veya 65 mm çapında olur. (BYKHY Madde 94, 1/a/2-3)"),
                        ("12", "İtfaiye su verme ağzı", "Yüksek binalarda veya bina oturma alanı 1.000 m²’den büyük binalarda veya cephe genişliği 75 m’yi aşan binalarda, itfaiyenin sisteme dışarıdan su basabilmesi için sulu yangın söndürme sistemlerine en az 100 mm nominal çapında itfaiye su verme bağlantısı yapılır. İtfaiye su verme bağlantısında iki adet 65 mm Storz tip rakor ve çek valf bulunur. İtfaiye araçlarının bağlantı ağzına ulaşma mesafesi 18 m’den fazla olamaz. (BYKHY Madde 97, 1)"),
                        ("13", "Yangın dolapları ve hortum bağlantı muslukları", "Yangın dolapları ve hortum bağlantı muslukları TS 2217’e uygun olacaktır."),
                        ("14", "Yangın dolaplarında kullanılacak hortum ve yerleşim yüksekliği", "Yangın dolapları içerisinde hortum olarak; kirlenme, çürüme ve küflenme göz önüne alınarak kurutmaya da gerek olmayan, basınca dayanıklı 1” çapında kauçuk hortumlar kullanılacaktır. Yangın dolapları yerden 80 cm ila 120 cm yükseklikte olacak şekilde yerleştirilecektir."),
                        ("15", "Yangın dolabı gövde ve su giriş tipi", "Yangın dolapları 1.5 mm DKP saçtan mamul, makaradan gövde içerisine su girişi yapabilecek tipte olacaktır."),
                        ("16", "Yangın dolabı boru sistemi", "Yangın dolapları sprinkler sistemi kolonlarından hat alınarak değil bağımsız boru sistemi olarak tasarlanmıştır."),
                        ("17", "İtfaiye su alma ağızlarının tesis edilmesi ve özellikleri", "Binada yangının büyümesi durumunda itfaiyenin ve eğitilmiş personelin yangına müdahale edebilmesi için kaçış merdiven yuvaları içerisinde, ayrıca itfaiye su alma ağızları da tesis edilecektir. İtfaiye su alma ağızları DIN normlarına uygun vanalı, 2½” çapında olacak ve her bağlantı ağzında 2½” x 2” ara rakor ve zincirli kapakları takılı halde bulunacaktır. İtfaiye su alma ağızlarına, gerektiğinde B tipi (110’luk) veya C tipi (85’lik) yassı hortum takılarak yangına müdahale edilebilecektir."),
                        ("18", "İtfaiye su alma ağızlarının bağlantısı ve basıncı", "İtfaiye su alma ağızlarının bağlantısı doğrudan yangın kollektöründen yapılacaktır. Yangın merdiven yuvaları içinde yer alacak olan riser kolonlarına yerden yaklaşık 1.0-1.2 m yükseklikte olacak şekilde itfaiye su alma ağızları bağlantısı yapılacaktır. İtfaiye su alma ağızlarında yassı hortum ucundaki lans girişinde, akış halinde basınç 6 bar olacaktır. Ve 9 barı geçmemelidir."),
                    ]
                    for _no, _baslik_yd, _metin_yd in _yd_esaslari:
                        if _no == "•":
                            st.markdown(f"**• {_baslik_yd}:**")
                            st.write(_metin_yd)
                        else:
                            st.markdown(f"**{_no}. {_baslik_yd}**")
                            st.write(_metin_yd)

                    # Kullanıcının çap tablosundan ÖNCE/Sonra kendi maddelerini ekleyebilmesi
                    # için kalıcı oturum listeleri. Rapor aynı session-state verisini kullanır.
                    st.session_state.setdefault("yangin_731_maddeler_cap_oncesi", [])
                    st.session_state.setdefault("yangin_731_maddeler_cap_sonrasi", [])

                    def _maddeleri_goster_ve_ekle(_liste_key, _yer_etiketi):
                        _liste = st.session_state[_liste_key]
                        for _idx, _madde in enumerate(_liste, start=19):
                            st.markdown(f"**{_idx}. {_madde.get('baslik','')}**")
                            st.write(_madde.get('metin',''))

                        # ÖNEMLİ: Burada st.form/st.form_submit_button kullanmıyoruz.
                        # Proje geri yükleme mekanizması eski Streamlit widget durumlarını
                        # session_state'e taşıyabildiği için form submit widget'ında
                        # StreamlitValueAssignmentNotAllowedError oluşabiliyordu.
                        # Bunun yerine açık ve benzersiz widget anahtarları kullanıyoruz.
                        # Anahtarların sonundaki _button ifadesi ana uygulamanın proje
                        # kaydında bunları widget olarak filtrelemesini de sağlar.
                        _nonce_key = f"{_liste_key}_form_nonce"
                        st.session_state.setdefault(_nonce_key, 0)
                        _nonce = st.session_state[_nonce_key]
                        _b_key = f"{_liste_key}_baslik_{_nonce}_button"
                        _m_key = f"{_liste_key}_metin_{_nonce}_button"
                        _s_key = f"{_liste_key}_submit_{_nonce}_button"

                        with st.expander(f"➕ {_yer_etiketi} yeni madde ekle", expanded=False):
                            _b = st.text_input(
                                f"Madde başlığı — {_yer_etiketi}",
                                key=_b_key,
                            )
                            _m = st.text_area(
                                f"Madde açıklaması — {_yer_etiketi}",
                                height=90,
                                key=_m_key,
                            )
                            _submit_madde = st.button(
                                f"Maddeyi ekle — {_yer_etiketi}",
                                key=_s_key,
                                type="primary",
                            )

                        if _submit_madde:
                            if _b.strip() and _m.strip():
                                _liste.append({"baslik": _b.strip(), "metin": _m.strip()})
                                st.session_state[_liste_key] = _liste
                                # Yeni nonce yeni widget anahtarları üretir; eski input
                                # değerlerini silmeye/üzerine yazmaya gerek kalmaz.
                                st.session_state[_nonce_key] = _nonce + 1
                                st.rerun()
                            else:
                                st.warning("Madde başlığı ve açıklaması birlikte girilmelidir.")

                    st.markdown("**ÇAP TABLOSUNDAN ÖNCE EK MADDELER**")
                    _maddeleri_goster_ve_ekle("yangin_731_maddeler_cap_oncesi", "Çap tablosundan önce")

                    st.markdown("**YANGIN DOLAP SİSTEMİ İÇİN KULLANILAN ÇAP TABLOSU**")
                    st.table([
                        {"Dolap Sayısı": "(1) Dolap", "Boru Çapı": '2”'},
                        {"Dolap Sayısı": "2 ve daha fazla dolap", "Boru Çapı": '2½”'},
                    ])

                    st.markdown("**ÇAP TABLOSUNDAN SONRA EK MADDELER**")
                    _maddeleri_goster_ve_ekle("yangin_731_maddeler_cap_sonrasi", "Çap tablosundan sonra")


                    # --------------------------------------------------------------
                    # 7.3.2 - BYKHY EK-8/C
                    # --------------------------------------------------------------
                    st.markdown(
                        '<div style="font-size:18px; font-weight:800; font-style:italic; color:#1F4E79; margin-top:14px;">'
                        "7.3.2 BİNA İÇİ HORTUM SİSTEMİ TASARIM DEBİLERİ TESPİTİ</div>",
                        unsafe_allow_html=True,
                    )
                    st.write(
                        "Binaların Yangından Korunması Hakkında Yönetmelik Ek-8/C, bina tehlike sınıfına "
                        "göre yangın dolabı ve hidrant sistemi için ilâve edilecek su ihtiyaçlarını belirler. "
                        "Aşağıdaki değerler 7.2 bölümünde seçilen etkin yangın tehlike sınıfına göre otomatik seçilir."
                    )

                    # Ek-8/C tablosundaki hidrant değerleri bu bölümde gösterilmez ve
                    # hesaplanmaz. Hidrant sistemi 7.4 bölümünde ayrıca geliştirilecektir.
                    _ek8c_verileri = [
                        ("Düşük tehlike", 100, 30),
                        ("Orta Tehlike-1-2", 100, 60),
                        ("Orta Tehlike-3-4", 100, 60),
                        ("Yüksek Tehlike", 200, 90),
                    ]
                    _etkin_73 = str(st.session_state.get("yangin_722_etkin_sinif", "")).strip()
                    if not _etkin_73:
                        _etkin_73 = str(st.session_state.get("yangin_722_otomatik_sinif", "")).strip()

                    def _ek8c_grup_esle(_sinif):
                        _s = str(_sinif or "").lower().replace("–", "-").replace(" ", "")
                        if _s.startswith("düşük"):
                            return "Düşük tehlike"
                        if _s.startswith("ortatehlike-1") or _s.startswith("ortatehlike-2"):
                            return "Orta Tehlike-1-2"
                        if _s.startswith("ortatehlike-3") or _s.startswith("ortatehlike-4"):
                            return "Orta Tehlike-3-4"
                        if _s.startswith("yüksek"):
                            return "Yüksek Tehlike"
                        return ""

                    _ek8c_secili_grup = _ek8c_grup_esle(_etkin_73)
                    _ek8c_kayit = next((x for x in _ek8c_verileri if x[0] == _ek8c_secili_grup), None)

                    st.markdown("**BYKHY Ek-8/C — Yangın Dolapları İçin İlâve Edilecek Su İhtiyacı**")
                    _h = st.columns([2.5, 1.8, 1.1])
                    for _c, _txt in zip(_h, ["Bina Tehlike Sınıfı", "İlave Yangın Dolabı Debisi (L/dak)", "Süre (dak)"]):
                        _c.markdown(f"**{_txt}**")
                    for _grup, _q_dolap, _sure in _ek8c_verileri:
                        _sec = _grup == _ek8c_secili_grup
                        _bg = "#FFF2CC" if _sec else "#FFFFFF"
                        _cols = st.columns([2.5, 1.8, 1.1])
                        for _c, _val in zip(_cols, [_grup, f'{_q_dolap:,}'.replace(',', '.'), _sure]):
                            _c.markdown(
                                f'<div style="background-color:{_bg}; border:1px solid #D9D9D9; padding:7px 8px; min-height:32px;">{_val}</div>',
                                unsafe_allow_html=True,
                            )

                    if _ek8c_kayit:
                        _q_dolap = int(_ek8c_kayit[1])
                        _sure = int(_ek8c_kayit[2])

                        # Bu değerler sonraki 7.12 yangın suyu deposu hesabının
                        # doğrudan veri kaynağıdır. Hidrant değerleri 7.4'e bırakılır.
                        st.session_state["yangin_73_ek8c_grup"] = _ek8c_kayit[0]
                        st.session_state["yangin_73_ek8c_yangin_dolabi_debisi_ldak"] = _q_dolap
                        st.session_state["yangin_73_ek8c_yangin_dolabi_suresi_dak"] = _sure
                        st.session_state["yangin_73_secili_yangin_suyu_debisi_ldak"] = _q_dolap

                        st.markdown(
                            f'<div style="background-color:#FFF2CC; border:1px solid #D6B656; padding:10px 12px; border-radius:4px; margin-top:10px;">'
                            f'<b>7.2’den otomatik seçilen değer:</b> {_ek8c_kayit[0]} → Yangın dolabı debisi: <b>{_q_dolap} L/dak</b> → Süre: <b>{_sure} dk</b>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        st.markdown(
                            f'<div style="background-color:#FFF2CC; border:1px solid #D6B656; padding:10px 12px; border-radius:4px; margin-top:8px;">'
                            f'<b>Seçilen yangın suyu debisi:</b> <b>{_q_dolap} L/dak</b>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        st.caption(
                            "Not: Bu bölümde yalnızca yangın dolabı için seçilen Ek-8/C değeri kullanılmaktadır. "
                            "Hidrant sistemi ve hidrant debisi 7.4 HİDRANT SİSTEMİ TASARIMI VE HESAPLAMALARI bölümünde ayrıca geliştirilecektir. "
                            "Seçilen yangın dolabı debisi ileride yangın suyu deposu kapasite hesabına aktarılacaktır."
                        )
                    else:
                        st.warning("7.2.2’de henüz geçerli bir yangın tehlike sınıfı seçilmediği için yangın dolabı debisi otomatik seçilemedi.")

                    st.caption("Kaynak: Binaların Yangından Korunması Hakkında Yönetmelik — Ek-8/C: Yangın Dolapları ve Hidrant Sistemi İçin İlâve Edilecek Su İhtiyaçları.")
                    st.markdown(
                        '<div style="font-size:15px; margin-top:14px; padding:8px 0;"><i>'
                        'Kısaltmalar: BYKHY – Binaların Yangından Korunması Hakkında Yönetmelik; '
                        'L/dak – litre/dakika; kPa – kilopaskal</i></div>',
                        unsafe_allow_html=True,
                    )

            elif _baslik.startswith("7.4 "):
                with st.expander("7.4 HİDRANT SİSTEMİ TASARIMI VE HESAPLAMALARI", expanded=True):
                    st.markdown(
                        '<div style="font-size:22px; font-weight:800; font-style:italic; color:#1F4E79;">'
                        "HİDRANT SİSTEMİ TASARIMI VE HESAPLAMALARI</div>",
                        unsafe_allow_html=True,
                    )
                    _ui_74_rapor_key = "yangin_74_rapor_ui"
                    st.session_state.setdefault(_ui_74_rapor_key, bool(st.session_state.get("rapor_bolum_74", True)))
                    st.checkbox(
                        "7.4 rapora eklensin",
                        key=_ui_74_rapor_key,
                        on_change=lambda: st.session_state.__setitem__(
                            "rapor_bolum_74", bool(st.session_state.get("yangin_74_rapor_ui", True))
                        ),
                    )

                    st.markdown(
                        '<div style="font-size:18px; font-weight:800; font-style:italic; color:#1F4E79; margin-top:12px;">'
                        "7.4.1 HİDRANT SİSTEMİ YÖNETMELİK ESASLARI VE TASARIM KRİTERLERİ</div>",
                        unsafe_allow_html=True,
                    )
                    _hidrant_esaslari = [
                        ("1", "Hidrant sisteminin amacı ve yerleşimi", "Yapıların yangından korunmasında, ilk müdahalede söndürülemeyen yangınlara dışarıdan müdahale edebilmek için mümkün olduğunca yapının veya binanın bütün çevresini kapsayacak şekilde hidrant sistemi tesis edilir. Hidrantların itfaiye araçlarının kolay yanaşabileceği ve bağlantı yapabileceği şekilde düzenlenmesi gerekir. (BYKHY Madde 95, 1)"),
                        ("2", "Hidrant sistemi tasarım debisi ve basıncı", "Hidrant sistemi dizayn debisi en az 1.900 L/dak olacak şekilde tasarlanır. Debi, binanın tehlike sınıfına göre artırılır. Hidrant çıkışında 700 kPa basınç olması gerekir. (BYKHY Madde 95, 2)"),
                        ("3", "Hidrantlar arası uzaklık", "Hidrantlar arası uzaklık çok riskli bölgelerde 50 m, riskli bölgelerde 100 m, orta riskli bölgelerde 125 m ve az riskli bölgelerde 150 m alınır. (BYKHY Madde 95, 3)"),
                        ("4", "Hidrantların bina çevresindeki konumu", "Normal şartlarda hidrantlar, korunan binalardan ortalama 5 ilâ 15 m kadar uzağa yerleştirilir. (BYKHY Madde 95, 4)"),
                        ("5", "Hidrant besleme borusu çapı", "Hidrant sistemine suyu sağlayan boru donanımında ring sistemi mevcut değil ise kullanılabilecek en düşük boru çapı 100 mm olacak şekilde ve hidrolik hesaba göre belirlenir. (BYKHY Madde 95, 5)"),
                        ("6", "Hidrant tipi ve hat kesme vanaları", "Sistemde kullanılacak hidrantların ilgili Türk Standartlarına uygun yerüstü yangın hidrantı olması gerekir. Hidrant yenilenmesi ve bakım işlemlerini kolaylaştırmak amacıyla uygun noktalarda yeraltı veya yerüstü yahut her iki tip hat kesme vanaları temin ve tesis edilir. (BYKHY Madde 95, 6)"),
                        ("7", "Yerleşim alanlarında dış hidrant sistemi", "İçerisinde her türlü kullanım alanı bulunan ve genel yerleşim alanlarından ayrı olarak planlanan yerleşim alanlarında yapılacak binaların taban alanları toplamının 5.000 m²’den büyük olması halinde dış hidrant sistemi yapılması mecburidir. Yönetmeliğin 7’nci maddesinin on ikinci fıkrası kapsamındaki alanlarda da dış hidrant sistemi yapılır. (BYKHY Madde 95, 7)"),
                        ("8", "İtfaiye araçlarının ulaşamadığı yerleşim alanları", "İtfaiye araçlarının giremediği veya manevra yapamadığı, ulaşım imkânı olmayan yerleşim mahallerinde uygun yerlere yerüstü yangın hidrantları veya pompa ile teçhiz edilmiş yeterli kapasitede yangın havuzları ve sarnıçları yapılır. (BYKHY Madde 95, 8)"),
                        ("9", "Bina çevresinde hidrant yerleşimi", "Bina çevresinde meydana gelebilecek yangınlara müdahale edilebilmesi, dışarıdan içeriye hortum serilerek su verilebilmesi ve itfaiye geldiği zaman su alabilmesi için hidrant sistemi kurulacaktır. Bina girişlerine, köşe başlarına ve açık otopark çevresine yakın yerlere hidrant yerleşimi yapılacaktır."),
                        ("10", "Hidrant sistemi ring hattı", "Hidrant sistemi için sulu söndürme sistemleri kollektöründen ayrı bir hat alınacak ve bina çevresinde ring sistemi oluşturulacaktır."),
                        ("11", "Hidrantların bina girişleri ve köşe başlarındaki yerleşimi", "Hidrantlar bina girişlerine ve köşe başlarına yakın olmak üzere yerleştirilecektir."),
                        ("12", "Hidrant sistemi boru malzemesi", "Sistemde yüksek yoğunluklu polietilen borular veya ductile borular kullanılacaktır."),
                        ("13", "Boru hatlarının toprak altı döşeme derinliği", "Sistemde kullanılan borular hem mekanik hasarları önlemek hem de donmaya karşı tedbir almak amacı ile en az 100 cm derinlikte toprak altına yerleştirilecektir."),
                        ("14", "Yerüstü hidrantları ve kesme vanaları", "Sistemde kullanılacak hidrantlar yer üstü yangın hidrantları olacaktır. Hidrant sisteminde, hidrant yenilenmesini ve bakım işlemlerinin yapılmasını kolaylaştıracak şekilde her hidrantın girişinde kesme vanaları yerleştirilecektir."),
                        ("15", "Kuru tip hidrant özellikleri", "İtfaiye elemanları ve eğitilmiş personelin kullanımına olanak verebilmesi için, çıkış ağızları 2 x 2½” iki çıkış ağızlı, anma boyutu 4” olan kuru tip hidrantlar monte edilecektir. (TS 2821/1)."),
                        ("16", "Hidrant otomatik boşaltma düzeni", "Hidrantların otomatik çalışan bir boşaltma düzeni olacak, hidrant vanası açıkken bu sistem kapalı olup, hidrant vanası kapatıldığında gövde de kalan su otomatik olarak boşalacaktır."),
                        ("17", "Hidrant ayağı ve drenaj düzeni", "Hidrant ayakları taş veya beton düz bir zemin üzerine oturtulacak olup, hidrant gövdesindeki suyun drenajı için kullanılacak otomatik tahliye donanımının çevresi küçük çakıl taşları ile doldurulacaktır."),
                        ("18", "Hortum bağlantı ağızlarının yüksekliği", "Hidrant üzerindeki hortum bağlantı ağızlarının yerden yüksekliği en az 45 cm olacak şekilde yerleştirilecektir."),
                        ("19", "Boru hatlarının flushing işlemi", "Yapım esnasında, boru içerisinde kalan yabancı malzemeleri temizlemek için su ile flushing yapılacaktır. Flushing esnasında, borular içerisindeki hız en az 3 m/s olacaktır."),
                    ]
                    for _no, _baslik_h, _metin_h in _hidrant_esaslari:
                        st.markdown(f"**{_no}. {_baslik_h}**")
                        st.write(_metin_h)

                    st.session_state.setdefault("yangin_741_maddeler_cap_oncesi", [])
                    st.session_state.setdefault("yangin_741_maddeler_cap_sonrasi", [])

                    def _hidrant_maddeleri_goster_ve_ekle(_liste_key, _yer_etiketi):
                        _liste = st.session_state[_liste_key]
                        _bas_no = 20
                        if _liste_key.endswith("sonrasi"):
                            _bas_no = 20 + len(st.session_state.get("yangin_741_maddeler_cap_oncesi", []))
                        for _idx, _madde in enumerate(_liste, start=_bas_no):
                            st.markdown(f"**{_idx}. {_madde.get('baslik','')}**")
                            st.write(_madde.get('metin',''))

                        # 7.3.1 ile aynı güvenli widget yaklaşımı: form submit yerine
                        # benzersiz st.button kullanıyoruz. Böylece eski proje dosyalarından
                        # kalan widget state'i form submit widget'ına yazılamıyor.
                        _nonce_key = f"{_liste_key}_form_nonce"
                        st.session_state.setdefault(_nonce_key, 0)
                        _nonce = st.session_state[_nonce_key]
                        _b_key = f"{_liste_key}_baslik_{_nonce}_button"
                        _m_key = f"{_liste_key}_metin_{_nonce}_button"
                        _s_key = f"{_liste_key}_submit_{_nonce}_button"

                        with st.expander(f"➕ {_yer_etiketi} yeni madde ekle", expanded=False):
                            _b = st.text_input(
                                f"Madde başlığı — {_yer_etiketi}",
                                key=_b_key,
                            )
                            _m = st.text_area(
                                f"Madde açıklaması — {_yer_etiketi}",
                                height=90,
                                key=_m_key,
                            )
                            _submit_madde = st.button(
                                f"Maddeyi ekle — {_yer_etiketi}",
                                key=_s_key,
                                type="primary",
                            )

                        if _submit_madde:
                            if _b.strip() and _m.strip():
                                _liste.append({"baslik": _b.strip(), "metin": _m.strip()})
                                st.session_state[_liste_key] = _liste
                                st.session_state[_nonce_key] = _nonce + 1
                                st.rerun()
                            else:
                                st.warning("Madde başlığı ve açıklaması birlikte girilmelidir.")

                    st.markdown("**HİDRANT MADDELERİ — EK MADDELER (8. MADDEDEN SONRA)**")
                    _hidrant_maddeleri_goster_ve_ekle("yangin_741_maddeler_cap_oncesi", "8. maddeden sonra")

                    # 7.4.2 - HİDRANT SİSTEMİ TASARIM DEBİSİ TESPİTİ
                    # BYKHY Ek-8/C
                    # --------------------------------------------------------------
                    st.markdown(
                        '<div style="font-size:18px; font-weight:800; font-style:italic; color:#1F4E79; margin-top:18px;">'
                        "7.4.2 HİDRANT SİSTEMİ TASARIM DEBİSİ TESPİTİ</div>",
                        unsafe_allow_html=True,
                    )
                    st.write(
                        "Binaların Yangından Korunması Hakkında Yönetmelik Ek-8/C'de, bina tehlike sınıfına göre "
                        "hidrant sistemi için ilave edilecek su ihtiyaçları belirlenmiştir. Aşağıdaki tablo, 7.2 bölümünde "
                        "seçilen etkin yangın tehlike sınıfına göre otomatik olarak seçim yapar. Ayrıca Madde 95/2 gereği "
                        "hidrant sistemi dizayn debisinin en az 1.900 L/dak olması ve hidrant çıkışında 700 kPa basınç "
                        "sağlanması gerektiği dikkate alınır."
                    )

                    _ek8c_hidrant_verileri = [
                        ("Düşük tehlike", 400, 30),
                        ("Orta Tehlike-1-2", 400, 60),
                        ("Orta Tehlike-3-4", 1000, 60),
                        ("Yüksek Tehlike", 1500, 90),
                    ]
                    _etkin_74 = str(st.session_state.get("yangin_722_etkin_sinif", "")).strip()
                    if not _etkin_74:
                        _etkin_74 = str(st.session_state.get("yangin_722_otomatik_sinif", "")).strip()

                    def _ek8c_hidrant_grup_esle(_sinif):
                        _s = str(_sinif or "").lower().replace("–", "-").replace(" ", "")
                        if _s.startswith("düşük"):
                            return "Düşük tehlike"
                        if _s.startswith("ortatehlike-1") or _s.startswith("ortatehlike-2"):
                            return "Orta Tehlike-1-2"
                        if _s.startswith("ortatehlike-3") or _s.startswith("ortatehlike-4"):
                            return "Orta Tehlike-3-4"
                        if _s.startswith("yüksek"):
                            return "Yüksek Tehlike"
                        return ""

                    _ek8c_hidrant_secili_grup = _ek8c_hidrant_grup_esle(_etkin_74)
                    _ek8c_hidrant_kayit = next(
                        (x for x in _ek8c_hidrant_verileri if x[0] == _ek8c_hidrant_secili_grup),
                        None,
                    )

                    st.markdown("**BYKHY Ek-8/C — Hidrant Sistemi İçin İlâve Edilecek Su İhtiyacı**")
                    _hh = st.columns([2.5, 1.8, 1.1])
                    for _c, _txt in zip(_hh, [
                        "Bina Tehlike Sınıfı", "İlave Hidrant Debisi (L/dak)", "Süre (dak)"
                    ]):
                        _c.markdown(f"**{_txt}**")
                    for _grup, _q_hidrant, _sure in _ek8c_hidrant_verileri:
                        _sec = _grup == _ek8c_hidrant_secili_grup
                        _bg = "#FFF2CC" if _sec else "#FFFFFF"
                        _cols = st.columns([2.5, 1.8, 1.1])
                        for _c, _val in zip(
                            _cols,
                            [_grup, f'{_q_hidrant:,}'.replace(',', '.'), _sure],
                        ):
                            _c.markdown(
                                f'<div style="background-color:{_bg}; border:1px solid #D9D9D9; padding:7px 8px; min-height:32px;">{_val}</div>',
                                unsafe_allow_html=True,
                            )

                    if _ek8c_hidrant_kayit:
                        _q_hidrant = int(_ek8c_hidrant_kayit[1])
                        _sure_hidrant = int(_ek8c_hidrant_kayit[2])
                        st.session_state["yangin_74_ek8c_grup"] = _ek8c_hidrant_kayit[0]
                        st.session_state["yangin_74_ek8c_hidrant_debisi_ldak"] = _q_hidrant
                        st.session_state["yangin_74_ek8c_hidrant_suresi_dak"] = _sure_hidrant

                        st.markdown(
                            f'<div style="background-color:#FFF2CC; border:1px solid #D6B656; padding:10px 12px; border-radius:4px; margin-top:10px;">'
                            f'<b>7.2’den otomatik seçilen değer:</b> {_ek8c_hidrant_kayit[0]} → Hidrant ilave debisi: <b>{_q_hidrant} L/dak</b> → Süre: <b>{_sure_hidrant} dk</b>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        st.markdown(
                            f'<div style="background-color:#FFF2CC; border:1px solid #D6B656; padding:10px 12px; border-radius:4px; margin-top:8px;">'
                            f'<b>Seçilen hidrant ilave debisi:</b> <b>{_q_hidrant} L/dak</b>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                    else:
                        st.warning("7.2.2’de henüz geçerli bir yangın tehlike sınıfı seçilmediği için hidrant ilave debisi otomatik seçilemedi.")

                    st.caption(
                        "Not: Ek-8/C değerleri hidrant sistemi için ilave edilecek su ihtiyacını gösterir. "
                        "Hidrant sisteminin genel tasarımında BYKHY Madde 95/2'de belirtilen en az 1.900 L/dak dizayn debisi "
                        "ve 700 kPa hidrant çıkış basıncı ayrıca dikkate alınacaktır."
                    )
                    st.markdown(
                        '<div style="font-size:15px; margin-top:14px; padding:8px 0;"><i>'
                        'Kısaltmalar: BYKHY – Binaların Yangından Korunması Hakkında Yönetmelik; '
                        'L/dak – litre/dakika; kPa – kilopaskal</i></div>',
                        unsafe_allow_html=True,
                    )

            elif _baslik.startswith("7.5 "):
                # --------------------------------------------------------------
                # 7.5 SPRİNKLER (YAĞMURLAMA) SİSTEMİ
                # --------------------------------------------------------------
                with st.expander("7.5 SPRİNKLER (YAĞMURLAMA) SİSTEMİ TASARIM VE HESAPLAMALARI", expanded=True):
                    st.markdown(
                        '<div style="font-size:22px; font-weight:800; font-style:italic; color:#1F4E79;">'
                        "SPRİNKLER (YAĞMURLAMA) SİSTEMİ TASARIM VE HESAPLAMALARI</div>",
                        unsafe_allow_html=True,
                    )
                    _ui_75_rapor_key = "yangin_75_rapor_ui"
                    st.session_state.setdefault(_ui_75_rapor_key, bool(st.session_state.get("rapor_bolum_75", True)))
                    st.checkbox(
                        "7.5 rapora eklensin",
                        key=_ui_75_rapor_key,
                        on_change=lambda: st.session_state.__setitem__(
                            "rapor_bolum_75", bool(st.session_state.get(_ui_75_rapor_key, True))
                        ),
                    )

                    # ----------------------------------------------------------
                    # 7.5.1 - BYKHY MADDE 96
                    # ----------------------------------------------------------
                    st.markdown(
                        '<div style="font-size:18px; font-weight:800; font-style:italic; color:#1F4E79; margin-top:12px;">'
                        "7.5.1 SPRİNKLER (YAĞMURLAMA) SİSTEMİ YÖNETMELİK ESASLARI VE TASARIM KRİTERLERİ</div>",
                        unsafe_allow_html=True,
                    )
                    # 7.2 ile bağlantı: bina/kullanım seçimi ve etkin yangın tehlike sınıfı
                    # burada yeniden seçtirilmez; 7.2'deki mevcut session-state verisi kullanılır.
                    _secili_kullanim_75 = [
                        str(x.get("etiket", "")).strip()
                        for x in _yangin_721_secili_kayitlar()
                        if str(x.get("etiket", "")).strip()
                    ]
                    _kullanim_metni_75 = " ".join(_secili_kullanim_75).lower()
                    _etkin_sinif_75 = str(st.session_state.get("yangin_722_etkin_sinif", "")).strip()
                    if not _etkin_sinif_75:
                        _etkin_sinif_75 = str(st.session_state.get("yangin_722_otomatik_sinif", "")).strip()
                    _yapi_h_75 = float(st.session_state.get("yangin_genel_yapi_yuksekligi_m", 0.0) or 0.0)
                    _otopark_alan_75 = float(st.session_state.get("yangin_genel_kapali_otopark_alan_m2", 0.0) or 0.0)
                    _otopark_arac_75 = int(st.session_state.get("yangin_genel_otopark_arac_sayisi", 0) or 0)

                    # Seçilen kullanım alanına karşılık gelen Madde 96/2 ifadesi.
                    # Buradaki sarı vurgu “zorunluluk kesinleşti” anlamına gelmez;
                    # 7.2'de seçilen kullanım alanının Madde 96/2'deki karşılığını gösterir.
                    _spr_75_vurgu = []
                    if "konut" in _kullanim_metni_75:
                        _spr_75_vurgu.append("yapı yüksekliği 51,50 m’yi geçen konutlarda")
                    else:
                        _spr_75_vurgu.append("yapı yüksekliği 30,50 m’den fazla olan konut haricindeki bütün binalarda")
                    if any(k in _kullanim_metni_75 for k in ["otopark"]):
                        _spr_75_vurgu.append("ilgili kapalı otoparklarda")
                    if any(k in _kullanim_metni_75 for k in ["otel", "yurt", "pansiyon", "misafirhane"]):
                        _spr_75_vurgu.append("belirli büyüklükteki otel, yurt, pansiyon ve misafirhanelerde")
                    if any(k in _kullanim_metni_75 for k in ["büyük mağaza", "alışveriş", "ticaret", "eğlence", "toplanma"]):
                        _spr_75_vurgu.append("toplam alanı 2000 m²’nin üzerinde olan katlı mağaza, alışveriş, ticaret, eğlence ve toplanma yerlerinde")

                    st.markdown(
                        '<div style="background-color:#FFF2CC; border:1px solid #D6B656; padding:10px 12px; border-radius:4px; margin:8px 0 12px 0;">'
                        f'<b>7.2’den gelen bina/kullanım alanı:</b> {", ".join(_secili_kullanim_75) if _secili_kullanim_75 else "Seçilmedi"}<br>'
                        f'<b>7.2’den gelen etkin yangın tehlike sınıfı:</b> {_etkin_sinif_75 or "Belirlenemedi"}<br>'
                        f'<b>Madde 96/2’de ilişkili ifade:</b> {", ".join(_spr_75_vurgu) if _spr_75_vurgu else "Belirlenemedi"}'
                        '</div>',
                        unsafe_allow_html=True,
                    )

                    _sprinkler_esaslari = [
                        ("1", "Yağmurlama sisteminin amacı ve kapsamı", "Yağmurlama sisteminin amacı; yangına erken tepki verilmesini sağlamak, yangını kontrol altına almak ve söndürmek için belirli bir süre içerisinde tasarım alanı üzerine belirlenen miktarda su boşaltmaktır. Sistem; alarm verilmesi ve itfaiyenin çağrılması gibi acil durum fonksiyonlarını da aktif hâle getirebilir. Yağmurlama sistemi; yağmurlama başlıkları, borular, bağlantı parçaları ve askılar, tesisat kontrol vanaları, alarm zilleri, akış göstergeleri, su pompaları ve acil durum güç kaynağı gibi elemanlardan meydana gelir. Yağmurlama sistemi elemanlarının TS EN 12259’a uygun olması şarttır. (BYKHY Madde 96, 1)"),
                        ("2", "Otomatik yağmurlama sistemi yapılması gereken yerler", "Yapı yüksekliği 30,50 m’den fazla olan konut haricindeki bütün binalarda; yapı yüksekliği 51,50 m’yi geçen konutlarda; ilgili kapalı otoparklarda; belirli büyüklükteki otel, yurt, pansiyon ve misafirhanelerde; toplam alanı 2000 m²’nin üzerinde olan katlı mağaza, alışveriş, ticaret, eğlence ve toplanma yerlerinde; toplam alanı 1000 m²’den fazla olan kolay alevlenici ve parlayıcı madde üretilen veya bulundurulan yapılarda otomatik yağmurlama sistemi kurulması mecburidir. (BYKHY Madde 96, 2)"),
                        ("3", "Yağmurlama yapılmayabilecek mahaller", "Yanıcı malzeme içermeyen ve yanıcı malzeme depolanmayan ıslak hacimlere, yanıcı malzeme ihtiva etmeyen ve yangına dirençli yapı elemanları ile ayrılan yangın merdiveni yuvalarına, asansör kuyusuna ve gazlı, kuru toz, su sprey ve benzeri diğer otomatik söndürme sistemleri ile korunan mahallere yağmurlama sistemi yapılmayabilir. (BYKHY Madde 96, 3)"),
                        ("4", "Yağmurlama yapılmayacak mahaller", "Su ile genişleyen veya reaksiyona girerek yangının büyümesine sebep olabilecek maddelerin bulunduğu mahallere yağmurlama sistemi yapılmaz. (BYKHY Madde 96, 4)"),
                        ("5", "Tasarım standardı ve sprinkler başlıklarının yerleşimi", "Yağmurlama sistemi tasarımı TS EN 12845’e göre yapılır. Yağmurlama başlıklarının yerleştirilmesinde kullanım alanının tehlike sınıfı ve yağmurlama başlığının koruma alanı dikkate alınır. Düşük Tehlike ve Orta Tehlike-1 kullanım alanlarında bir adet standart yağmurlama başlığı en çok 21 m² alanı koruyacak şekilde yerleştirilebilir. (BYKHY Madde 96, 5)"),
                        ("6", "Deprem bölgelerinde boru tesisatının korunması", "Birinci ve ikinci derece deprem bölgelerinde, sismik hareketlere karşı ana kolonların herhangi bir yöne sürüklenmemesi için dört yollu destek kullanılır. 65 mm ve daha büyük nominal çaplı boruların katlardan ana dağıtım borularına bağlanmasında esnek bağlantılar, boruların tavanlara tutturulmasında iki yollu enlemesine ve boylamasına sabitleme askı elemanları kullanılır. Dilatasyon geçişlerinde her üç yönde hareketi karşılayacak detaylar uygulanır. (BYKHY Madde 96, 6)"),
                        ("7", "Yangın zonlarında kontrol ve test düzeni", "Yağmurlama sistemi ana besleme borusu birden fazla yangın zonuna hitap ediyor ise her bir zon veya kolon hattına akış anahtarları, test ve drenaj vanası ve izleme anahtarlı hat kesme vanası konulur. (BYKHY Madde 96, 7)"),
                        ("8", "Yedek yağmurlama başlıkları", "Muhtemel küçük çaplı yangınlarda yağmurlama başlığının patlaması veya birkaçının hasara uğraması hâlinde hemen değiştirilir. Yangın güvenlik sisteminin sürekliliği için 6 adetten az olmamak kaydıyla sistemin büyüklüğüne göre yeterli miktarda yedek yağmurlama başlığı ve başlığın değiştirilmesi için özel anahtarlar bulundurulur. (BYKHY Madde 96, 8)"),
                        ("9", "Kesme vanaları ve vanaların açık tutulması", "Yağmurlama sistemini besleyen borular üzerinde kesme vanaları bulunur. Boru hatlarında bulunan vanaların, bölgesel kontrol vanalarının ve su kaynağı ile yağmurlama sistemi arasında bulunan bütün vanaların devamlı açık kalmasını sağlayacak tedbirler alınır. (BYKHY Madde 96, 9)"),
                        ("10", "Basınç düşürücü vana ve manometreler", "Sistemde basınç düşürücü vana kullanılması hâlinde, her bir basınç düşürücü vananın önüne ve arkasına birer adet manometre konulur. (BYKHY Madde 96, 10)"),
                    ]
                    for _no, _baslik_s, _metin_s in _sprinkler_esaslari:
                        st.markdown(f"**{_no}. {_baslik_s}**")
                        if _no == "2" and _spr_75_vurgu:
                            # Madde metninin sadece 7.2'de seçilen kullanım alanıyla
                            # ilişkili kısmını sarı vurgula. Diğer metin aynen korunur.
                            _html_s = _metin_s
                            for _ifade in _spr_75_vurgu:
                                _html_s = _html_s.replace(
                                    _ifade,
                                    f'<mark style="background-color:#FFF2CC; padding:2px 4px; border-radius:3px; font-weight:700;">{_ifade}</mark>'
                                )
                            st.markdown(_html_s, unsafe_allow_html=True)
                        else:
                            st.write(_metin_s)

                    st.caption("Sarı vurgular, 7.2’de seçilen bina/kullanım alanının Madde 96/2’deki karşılığını gösterir. Sprinkler zorunluluğunun kesin değerlendirilmesinde bina yüksekliği, alan, otopark ve diğer yönetmelik koşulları ayrıca dikkate alınır.")
                    st.caption("Kaynak: Binaların Yangından Korunması Hakkında Yönetmelik — Madde 96: Yağmurlama sistemi.")

                    # ----------------------------------------------------------
                    # 7.5.2 - BYKHY EK-8/B
                    # ----------------------------------------------------------
                    st.markdown(
                        '<div style="font-size:18px; font-weight:800; font-style:italic; color:#1F4E79; margin-top:18px;">'
                        "7.5.2 SPRİNKLER (YAĞMURLAMA) SİSTEMİ TASARIM DEĞERLERİ TESPİTİ</div>",
                        unsafe_allow_html=True,
                    )
                    st.write(
                        "BYKHY Ek-8/B'ye göre yağmurlama sistemi tasarım yoğunluğu ve tasarım alanı, bina tehlike sınıfına göre aşağıdaki tablodan belirlenir. "
                        "Depolama alanları ve farklı özellikteki kullanım alanları için TS EN 12845 esas alınır."
                    )

                    _ek8b_sprinkler_verileri = [
                        ("Düşük Tehlike", "2,25", "84", "—"),
                        ("Orta Tehlike-1", "5,0", "72", "90"),
                        ("Orta Tehlike-2", "5,0", "144", "180"),
                        ("Orta Tehlike-3", "5,0", "216", "270"),
                        ("Orta Tehlike-4", "5,0", "360", "—"),
                        ("Yüksek Tehlike-1", "7,7", "260", "325"),
                        ("Yüksek Tehlike-2", "10,0", "260", "325"),
                        ("Yüksek Tehlike-3", "12,5", "260", "325"),
                        ("Yüksek Tehlike-4", "Yoğun su", "—", "—"),
                    ]
                    _etkin_75 = str(st.session_state.get("yangin_722_etkin_sinif", "")).strip()
                    if not _etkin_75:
                        _etkin_75 = str(st.session_state.get("yangin_722_otomatik_sinif", "")).strip()
                    _s75 = _etkin_75.lower().replace("–", "-").replace(" ", "")

                    def _ek8b_sprinkler_esle(_sinif):
                        _s = str(_sinif or "").lower().replace("–", "-").replace(" ", "")
                        for _kayit in _ek8b_sprinkler_verileri:
                            _etiket = _kayit[0].lower().replace("–", "-").replace(" ", "")
                            if _s == _etiket:
                                return _kayit[0]
                        return ""

                    _ek8b_sprinkler_secili = _ek8b_sprinkler_esle(_etkin_75)
                    _t75h = st.columns([2.0, 1.6, 1.55, 1.55])
                    for _c, _txt in zip(_t75h, [
                        "Tehlike Sınıfı", "Tasarım Yoğunluğu (mm/dak)",
                        "Koruma Alanı — Islak veya Ön Etkili (m²)",
                        "Koruma Alanı — Kuru veya Değişken (m²)"
                    ]):
                        _c.markdown(f"**{_txt}**")
                    for _grup, _yog, _islak, _kuru in _ek8b_sprinkler_verileri:
                        _sec = _grup == _ek8b_sprinkler_secili
                        _bg = "#FFF2CC" if _sec else "#FFFFFF"
                        _cols = st.columns([2.0, 1.6, 1.55, 1.55])
                        for _c, _val in zip(_cols, [_grup, _yog, _islak, _kuru]):
                            _c.markdown(
                                f'<div style="background-color:{_bg}; border:1px solid #D9D9D9; padding:7px 8px; min-height:32px;">{_val}</div>',
                                unsafe_allow_html=True,
                            )

                    if _ek8b_sprinkler_secili:
                        _sec75 = next(x for x in _ek8b_sprinkler_verileri if x[0] == _ek8b_sprinkler_secili)
                        st.session_state["yangin_75_ek8b_tehlike_sinifi"] = _sec75[0]
                        st.session_state["yangin_75_ek8b_tasarim_yogunlugu"] = _sec75[1]
                        st.session_state["yangin_75_ek8b_islak_on_etkili_alan"] = _sec75[2]
                        st.session_state["yangin_75_ek8b_kuru_degisken_alan"] = _sec75[3]
                        st.markdown(
                            f'<div style="background-color:#FFF2CC; border:1px solid #D6B656; padding:10px 12px; border-radius:4px; margin-top:10px;">'
                            f'<b>7.2’den otomatik seçilen değer:</b> {_sec75[0]} → Tasarım yoğunluğu: <b>{_sec75[1]} mm/dak</b> → Islak/ön etkili alan: <b>{_sec75[2]} m²</b> → Kuru/değişken alan: <b>{_sec75[3]} m²</b>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                    else:
                        st.warning("7.2.2’de henüz geçerli bir yangın tehlike sınıfı seçilmediği için Ek-8/B tasarım değeri otomatik seçilemedi.")

                    # ----------------------------------------------------------
                    # 7.5.2 - SPRİNKLER TASARIM DEBİSİ + EK-8/A SU DEPOSU ÖN HESABI
                    # Q = tasarım yoğunluğu × seçilen koruma alanı.
                    # Sprinkler debisi için süre kullanılmaz. Su deposu ön hacmi,
                    # BYKHY Ek-8/A tablosundan yapı yüksekliği / h aralığına göre
                    # otomatik seçilir.
                    # ----------------------------------------------------------
                    _spr_sistem_tipi_75 = st.selectbox(
                        "Sprinkler sistemi tipi (Ek-8/A)",
                        ["Islak / Ön etkili", "Kuru / Alternatif"],
                        index=0 if st.session_state.get("yangin_75_sistem_tipi", "Islak / Ön etkili") == "Islak / Ön etkili" else 1,
                        key="yangin_75_sistem_tipi",
                        help="Ek-8/A su deposu ön hesabında kullanılacak sistem tipidir. Yönetmelik tablosundaki grup eşleştirmesine göre otomatik depo hacmi belirlenir.",
                    )

                    if _ek8b_sprinkler_secili:
                        _sec75_debi = next((x for x in _ek8b_sprinkler_verileri if x[0] == _ek8b_sprinkler_secili), None)
                        if _sec75_debi:
                            try:
                                _yog75_num = float(str(_sec75_debi[1]).replace(',', '.'))
                                _alan75_islak = float(str(_sec75_debi[2]).replace(',', '.')) if str(_sec75_debi[2]).strip() not in {'—','-',''} else 0.0
                                _alan75_kuru = float(str(_sec75_debi[3]).replace(',', '.')) if str(_sec75_debi[3]).strip() not in {'—','-',''} else 0.0

                                # Sprinkler tasarım debisinde koruma alanı, seçilen sistem tipine
                                # göre alınır. Islak/ön etkili sistemde ıslak alan; kuru/alternatif
                                # sistemde kuru alan kullanılır. Artık iki alanın maksimumu alınmaz.
                                _spr_tip_norm = str(_spr_sistem_tipi_75 or '').lower()
                                if 'kuru' in _spr_tip_norm or 'alternatif' in _spr_tip_norm:
                                    _alan75 = _alan75_kuru
                                    _alan75_etiket = 'Kuru / Alternatif koruma alanı'
                                else:
                                    _alan75 = _alan75_islak
                                    _alan75_etiket = 'Islak / Ön etkili koruma alanı'

                                if _alan75 > 0 and _yog75_num > 0:
                                    _q75_spr = _yog75_num * _alan75
                                    st.session_state["yangin_75_sprinkler_debisi_ldak"] = float(_q75_spr)
                                    st.session_state["yangin_75_sprinkler_debisi_kaynagi"] = (
                                        "7.5.2 Ek-8/B tasarım yoğunluğu × " + _alan75_etiket
                                    )
                                    st.markdown(
                                        f'<div style="background-color:#FFF2CC; border:2px solid #D6B656; padding:10px 12px; border-radius:4px; margin-top:8px;">'
                                        f'<b>7.5.2 SPRİNKLER TASARIM DEBİSİ:</b> <b>{_q75_spr:.2f} L/dak</b><br>'
                                        f'{_yog75_num:g} mm/dak × {_alan75:g} m² = <b>{_q75_spr:.2f} L/dak</b><br>'
                                        f'<span style="font-size:13px;">{_alan75_etiket}</span>'
                                        f'</div>',
                                        unsafe_allow_html=True,
                                    )
                                else:
                                    st.session_state["yangin_75_sprinkler_debisi_ldak"] = 0.0
                                    st.warning(
                                        f"{_ek8b_sprinkler_secili} için {_spr_sistem_tipi_75} sistemine ait geçerli bir koruma alanı bulunmadığından sprinkler debisi hesaplanamadı."
                                    )
                            except (TypeError, ValueError):
                                st.session_state["yangin_75_sprinkler_debisi_ldak"] = 0.0
                                st.warning("Sprinkler tasarım yoğunluğu veya koruma alanı sayısal olarak okunamadı.")
                    else:
                        st.session_state["yangin_75_sprinkler_debisi_ldak"] = 0.0
                        st.warning("7.2'den geçerli bir yangın tehlike sınıfı gelmediği için sprinkler tasarım debisi hesaplanamadı.")

                    # ----------------------------------------------------------
                    # EK-8/A - Yağmurlama Sistemi, Yangın Dolabı ve Hidrant
                    # Tasarımı Ön Hesabı İçin Su Deposu En Az Hacmi
                    # ----------------------------------------------------------
                    _ek8a_yapi_h = float(st.session_state.get("yangin_genel_yapi_yuksekligi_m", 0.0) or 0.0)
                    _ek8a_h_araligi = ""
                    if _ek8a_yapi_h <= 15:
                        _ek8a_h_araligi = "h ≤ 15 m"
                    elif _ek8a_yapi_h <= 30:
                        _ek8a_h_araligi = "15 < h ≤ 30 m"
                    elif _ek8a_yapi_h <= 45:
                        _ek8a_h_araligi = "30 < h ≤ 45 m"
                    else:
                        _ek8a_h_araligi = "h > 45 m"

                    # Sınıf adlarını boşluk ve tire farklılıklarından bağımsız eşleştir.
                    # 7.2 ekranındaki "Orta Tehlike -1" biçimi, Ek-8/A'daki
                    # "Orta Tehlike-1" anahtarıyla aynı sınıfı ifade eder.
                    _ek8a_sinif = re.sub(
                        r"\\s+", "", str(_etkin_sinif_75 or "").strip().lower().replace("–", "-")
                    )
                    _ek8a_tip = str(_spr_sistem_tipi_75 or "Islak / Ön etkili").strip()

                    # Ek-8/A tablosunun mevzuattaki gruplanmış satırları aynen korunur.
                    _ek8a_hacimler = {
                        "Düşük Tehlike": [9.0, 10.0, 11.0],
                        "Orta Tehlike-1 ıslak": [55.0, 70.0, 80.0],
                        "Orta Tehlike-2 ıslak": [105.0, 125.0, 140.0],
                        "Orta Tehlike-3 ıslak": [135.0, 160.0, 185.0],
                        "Orta Tehlike-4 ıslak": [160.0, 185.0, 200.0],
                        "Orta Tehlike-1 kuru": [105.0, 125.0, 140.0],
                        "Orta Tehlike-2 kuru": [135.0, 160.0, 185.0],
                        "Orta Tehlike-3 kuru": [160.0, 185.0, 200.0],
                    }

                    # Ek-8/A seçim sütunu: yapı yüksekliğine göre otomatik belirlenir.
                    # 45 m üzerindeki yapılarda da tablodaki 30 < h <= 45 m
                    # sütunu kullanılır; ayrıca bir >45 m sütunu oluşturulmaz.
                    if _ek8a_yapi_h <= 0:
                        _ek8a_idx = -1
                    elif _ek8a_yapi_h <= 15:
                        _ek8a_idx = 0
                    elif _ek8a_yapi_h <= 30:
                        _ek8a_idx = 1
                    else:
                        _ek8a_idx = 2
                    if _ek8a_yapi_h > 45:
                        # 45 m üzerindeki yapılarda tabloda bulunan 30 < h ≤ 45 m
                        # sütunu kullanılır; ayrı bir >45 m sütunu varsayılmaz.
                        _ek8a_h_araligi = "30 < h ≤ 45 m"
                    _ek8a_anahtar = ""
                    if _ek8a_sinif.startswith("düşük tehlike"):
                        _ek8a_anahtar = "Düşük Tehlike"
                    elif _ek8a_sinif.startswith("orta tehlike-1"):
                        _ek8a_anahtar = "Orta Tehlike-1 ıslak" if "Islak" in _ek8a_tip else "Orta Tehlike-1 kuru"
                    elif _ek8a_sinif.startswith("orta tehlike-2"):
                        _ek8a_anahtar = "Orta Tehlike-2 ıslak" if "Islak" in _ek8a_tip else "Orta Tehlike-2 kuru"
                    elif _ek8a_sinif.startswith("orta tehlike-3"):
                        _ek8a_anahtar = "Orta Tehlike-3 ıslak" if "Islak" in _ek8a_tip else "Orta Tehlike-3 kuru"
                    elif _ek8a_sinif.startswith("orta tehlike-4"):
                        _ek8a_anahtar = "Orta Tehlike-4 ıslak" if "Islak" in _ek8a_tip else ""

                    _ek8a_depo_m3 = None
                    if _ek8a_idx >= 0 and _ek8a_anahtar in _ek8a_hacimler:
                        _ek8a_depo_m3 = _ek8a_hacimler[_ek8a_anahtar][_ek8a_idx]

                    # Otomatik hacmi belirledikten sonra kullanıcıya manuel değer girme imkanı verilir.
                    # Widget'ın session_state değerine sonradan doğrudan atama yapılmaz; böylece
                    # StreamlitValueAssignmentNotAllowedError oluşmaz.
                    _ek8a_auto_m3 = float(_ek8a_depo_m3) if _ek8a_depo_m3 is not None else 0.0
                    _ek8a_secim_modu = st.radio(
                        "Yangın suyu depo hacmi seçimi",
                        ["Otomatik (Ek-8/A)", "Manuel gir"],
                        horizontal=True,
                        key="yangin_75_ek8a_depo_secim_modu",
                    )
                    _ek8a_manuel_m3 = 0.0
                    if _ek8a_secim_modu == "Manuel gir":
                        # Widget oluşturulmadan ÖNCE varsayılan değeri hazırlıyoruz.
                        # Böylece number_input oluşturulduktan sonra aynı key'e değer
                        # atamaya çalışıp StreamlitValueAssignmentNotAllowedError üretmeyiz.
                        if "yangin_75_ek8a_depo_manuel_m3" not in st.session_state:
                            st.session_state["yangin_75_ek8a_depo_manuel_m3"] = _ek8a_auto_m3
                        _ek8a_manuel_m3 = st.number_input(
                            "Manuel yangın suyu depo hacmi (m³)",
                            min_value=0.0,
                            step=1.0,
                            format="%.2f",
                            key="yangin_75_ek8a_depo_manuel_m3",
                            help="Ek-8/A otomatik değerini geçersiz kılar. Girilen değer rapora ve 7.12 depo hesabına aktarılır.",
                        )
                        _ek8a_secilen_m3 = float(_ek8a_manuel_m3)
                        _ek8a_hesap_kaynagi = "Manuel kullanıcı girişi"
                    else:
                        _ek8a_secilen_m3 = _ek8a_auto_m3
                        _ek8a_hesap_kaynagi = "BYKHY Ek-8/A"

                    st.session_state["yangin_75_ek8a_yapi_yuksekligi_m"] = _ek8a_yapi_h
                    st.session_state["yangin_75_ek8a_h_araligi"] = _ek8a_h_araligi
                    st.session_state["yangin_75_ek8a_depo_min_hacim_m3"] = _ek8a_secilen_m3
                    st.session_state["yangin_75_ek8a_depo_otomatik_m3"] = _ek8a_auto_m3
                    st.session_state["yangin_75_ek8a_depo_manuel_m3_deger"] = float(_ek8a_manuel_m3)
                    st.session_state["yangin_75_ek8a_hesap_kaynagi"] = _ek8a_hesap_kaynagi

                    st.markdown(
                        '<div style="font-size:18px; font-weight:800; font-style:italic; color:#1F4E79; margin-top:18px;">'
                        "EK-8/A — YAĞMURLAMA SİSTEMİ, YANGIN DOLABI VE HİDRANT TASARIMI ÖN HESABI İÇİN SU DEPOSU EN AZ HACMİ</div>",
                        unsafe_allow_html=True,
                    )
                    st.write(
                        "Su deposu ön hesabında sprinkler debisine süre çarpanı uygulanmaz. "
                        "Ek-8/A tablosundaki su deposu en az hacmi, yapı yüksekliği / h aralığı ve sprinkler sistem tipine göre otomatik seçilir."
                    )

                    # Ek-8/A tablosunu 7.5 altında görünür şekilde gösteriyoruz.
                    # Otomatik seçilen tehlike sınıfı + sistem tipi + h aralığına karşılık gelen
                    # hücre yalnızca sarı renkle vurgulanır; diğer tablo verileri korunur.
                    _ek8a_tablo = [
                        ("Düşük Tehlike", "Islak veya ön uyarılı", "9", "10", "11"),
                        ("Orta Tehlike-1", "Islak veya ön uyarılı", "55", "70", "80"),
                        ("Orta Tehlike-1", "Kuru veya alternatif", "105", "125", "140"),
                        ("Orta Tehlike-2", "Islak veya ön uyarılı", "105", "125", "140"),
                        ("Orta Tehlike-2", "Kuru veya alternatif", "135", "160", "185"),
                        ("Orta Tehlike-3", "Islak veya ön uyarılı", "135", "160", "185"),
                        ("Orta Tehlike-3", "Kuru veya alternatif", "160", "185", "200"),
                        ("Orta Tehlike-4", "Islak veya ön uyarılı", "160", "185", "200"),
                        ("Orta Tehlike-4", "Kuru veya alternatif", "Hidrolik Hesap", "Hidrolik Hesap", "Hidrolik Hesap"),
                        ("Yüksek Tehlike", "Islak veya ön uyarılı", "Hidrolik Hesap", "Hidrolik Hesap", "Hidrolik Hesap"),
                        ("Yüksek Tehlike", "Kuru veya alternatif", "Hidrolik Hesap", "Hidrolik Hesap", "Hidrolik Hesap"),
                    ]
                    _ek8a_sinif_goster = re.sub(
                        r"\\s+", "", str(_etkin_sinif_75 or "").strip().lower().replace("–", "-")
                    )
                    _ek8a_secili_satir = None
                    for _i, _satir in enumerate(_ek8a_tablo):
                        _sinif_satir = re.sub(r"\\s+", "", _satir[0].lower().replace("–", "-"))
                        _tip_satir = _satir[1].lower()
                        if _ek8a_sinif_goster.startswith(_sinif_satir):
                            if (("ıslak" in _ek8a_tip.lower() or "ön" in _ek8a_tip.lower()) and "ıslak" in _tip_satir) or (("kuru" in _ek8a_tip.lower() or "alternatif" in _ek8a_tip.lower()) and "kuru" in _tip_satir):
                                _ek8a_secili_satir = _i
                                break

                    _ek8a_basliklar = [
                        "Yangın Tehlike Sınıfı",
                        "Sistem Tipi",
                        "h ≤ 15 m",
                        "15 < h ≤ 30 m",
                        "30 < h ≤ 45 m",
                    ]
                    _ek8a_cols = st.columns([2.0, 2.0, 1.15, 1.15, 1.15])
                    for _c, _baslik in zip(_ek8a_cols, _ek8a_basliklar):
                        _c.markdown(
                            f'<div style="background-color:#D9EAF7; border:1px solid #9FBAD0; padding:7px 8px; min-height:34px; font-weight:700;">{_baslik}</div>',
                            unsafe_allow_html=True,
                        )

                    for _i, _satir in enumerate(_ek8a_tablo):
                        _satir_secili = (_i == _ek8a_secili_satir)
                        _cols = st.columns([2.0, 2.0, 1.15, 1.15, 1.15])
                        for _j, (_c, _val) in enumerate(zip(_cols, _satir)):
                            _hucre_secili = _satir_secili and _j in (2, 3, 4) and (
                                (_ek8a_h_araligi == "h ≤ 15 m" and _j == 2) or
                                (_ek8a_h_araligi == "15 < h ≤ 30 m" and _j == 3) or
                                (_ek8a_h_araligi == "30 < h ≤ 45 m" and _j == 4) or
                                (_ek8a_yapi_h > 45 and _j == 4)
                            )
                            _bg = "#FFF2CC" if _hucre_secili else ("#FFF9E6" if _satir_secili else "#FFFFFF")
                            _border = "2px solid #D6B656" if _hucre_secili else "1px solid #D9D9D9"
                            _weight = "800" if _hucre_secili else ("700" if _satir_secili else "400")
                            _c.markdown(
                                f'<div style="background-color:{_bg}; border:{_border}; padding:7px 8px; min-height:32px; font-weight:{_weight};">{_val}</div>',
                                unsafe_allow_html=True,
                            )

                    if _ek8a_secili_satir is not None and _ek8a_depo_m3 is not None:
                        st.caption(
                            f"Sarı hücre: 7.2'den otomatik alınan tehlike sınıfı ({_etkin_sinif_75}) + "
                            f"seçilen sistem tipi ({_spr_sistem_tipi_75}) + {_ek8a_h_araligi} kriterine göre otomatik seçilen Ek-8/A hacmidir. "
                            f"45 m üzerindeki yapılarda 30 < h ≤ 45 m sütunundaki (45 m'ye karşılık gelen) değer kullanılır."
                        )

                    _e8c = st.columns(4)
                    _e8c[0].metric("Yapı yüksekliği", f"{_ek8a_yapi_h:g} m")
                    _e8c[1].metric("Ek-8/A h aralığı", _ek8a_h_araligi)
                    _e8c[2].metric("Yangın tehlike sınıfı", _etkin_sinif_75 or "Belirlenemedi")
                    _e8c[3].metric("Seçilen min. depo", f"{_ek8a_secilen_m3:g} m³" if _ek8a_secilen_m3 is not None else "-")
                    if _ek8a_depo_m3 is not None:
                        _ek8a_kaynak_metni = "Ek-8/A otomatik" if _ek8a_secim_modu == "Otomatik (Ek-8/A)" else "Manuel kullanıcı girişi"
                        st.markdown(
                            f'<div style="background-color:#FFF2CC; border:2px solid #D6B656; padding:12px 14px; border-radius:5px; margin-top:8px;">'
                            f'<b>Yangın suyu deposu için seçilen minimum hacim:</b> <b>{_ek8a_secilen_m3:g} m³</b><br>'
                            f'Kriter: {_ek8a_anahtar} → {_ek8a_h_araligi}<br>'
                            f'Kaynak: {_ek8a_kaynak_metni}'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                    else:
                        if _ek8a_yapi_h <= 0:
                            st.warning("Ek-8/A otomatik depo hacmi için 7.2 GENEL BİNA BİLGİLERİ bölümündeki Yapı yüksekliği (m) değeri girilmelidir.")
                        else:
                            st.warning(
                                "Ek-8/A bu yangın tehlike sınıfı / sprinkler sistemi kombinasyonu için otomatik hacim vermiyor. "
                                "Yüksek Tehlike ve Orta Tehlike-4 kuru/alternatif durumlarında hidrolik hesap esas alınmalıdır."
                            )

                    st.caption("Kaynak: Binaların Yangından Korunması Hakkında Yönetmelik — Ek-8/A ve Madde 92/4-6. Ek-8/A'daki h değeri, en alttaki ve en üstteki yağmurlama başlıkları arasındaki yükseklik için kullanılır; burada 7.2'de girilen yapı yüksekliği otomatik ön seçim girdisi olarak kullanılmıştır.")
                    st.caption("Kaynak: Binaların Yangından Korunması Hakkında Yönetmelik Kılavuzu — Ek-8/B: Yağmurlama Sisteminde Tasarım Yoğunlukları.")
                    st.markdown(
                        '<div style="font-size:15px; margin-top:14px; padding:8px 0;"><i>'
                        'Kısaltmalar: BYKHY – Binaların Yangından Korunması Hakkında Yönetmelik; '
                        'L/dak – litre/dakika; kPa – kilopaskal</i></div>',
                        unsafe_allow_html=True,
                    )

            elif _baslik.startswith("7.12 "):
                # --------------------------------------------------------------
                # 7.12 YANGIN SUYU DEPOLAMA SİSTEMİ
                # 7.3.2 + 7.4.2 + 7.5.2 seçilen değerleri otomatik veri kaynağıdır.
                # --------------------------------------------------------------
                with st.expander("7.12 YANGIN SUYU DEPOLAMA SİSTEMİ TASARIM VE HESAPLAMALARI", expanded=True):
                    st.markdown(
                        '<div style="font-size:22px; font-weight:800; font-style:italic; color:#1F4E79;">'
                        "YANGIN SUYU DEPOLAMA SİSTEMİ TASARIM VE HESAPLAMALARI</div>",
                        unsafe_allow_html=True,
                    )
                    _ui_712_rapor_key = "yangin_712_rapor_ui"
                    st.session_state.setdefault(_ui_712_rapor_key, bool(st.session_state.get("rapor_bolum_7120", True)))
                    st.checkbox(
                        "7.12 rapora eklensin",
                        key=_ui_712_rapor_key,
                        on_change=lambda: st.session_state.__setitem__(
                            "rapor_bolum_7120", bool(st.session_state.get("yangin_712_rapor_ui", True))
                        ),
                    )

                    _q73 = float(st.session_state.get("yangin_73_secili_yangin_suyu_debisi_ldak", 0) or 0)
                    _t73 = float(st.session_state.get("yangin_73_ek8c_yangin_dolabi_suresi_dak", 0) or 0)
                    _q74 = float(st.session_state.get("yangin_74_ek8c_hidrant_debisi_ldak", 0) or 0)
                    _t74 = float(st.session_state.get("yangin_74_ek8c_hidrant_suresi_dak", 0) or 0)
                    _q75 = float(st.session_state.get("yangin_75_sprinkler_debisi_ldak", 0) or 0)

                    st.markdown("### 7.3.2 / 7.4.2 / 7.5.2 Otomatik Aktarılan Debiler")
                    _dcols = st.columns(3)
                    with _dcols[0]:
                        st.metric("7.3.2 Yangın Dolabı", f"{_q73:,.2f} L/dak".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.caption(f"Süre: {_t73:g} dk")
                    with _dcols[1]:
                        st.metric("7.4.2 Hidrant", f"{_q74:,.2f} L/dak".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.caption(f"Süre: {_t74:g} dk")
                    with _dcols[2]:
                        st.metric("7.5.2 Sprinkler", f"{_q75:,.2f} L/dak".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.caption("Debi = tasarım yoğunluğu × seçilen koruma alanı")

                    # 7.12'de üç sistemin seçilen DEBİLERİ otomatik olarak birlikte gösterilir
                    # ve toplam yangın söndürme tasarım debisi ayrıca hesaplanır.
                    # ÖNEMLİ: Bu toplam DEBİ, depo hacmi hesabında kullanılmaz.
                    # 7.5.2 sprinkler debisi de kesinlikle Ek-8/A depo hacmine eklenmez.
                    # Sprinkler debisi ileride yangın/sprinkler pompası seçiminde kullanılmak üzere
                    # ayrı bir veri olarak saklanır.
                    _toplam_tasarim_debisi_ldak = _q73 + _q74 + _q75

                    # Su deposu hacmi yalnızca 7.5'te bina yüksekliği + tehlike sınıfı +
                    # sprinkler sistem tipine göre seçilen Ek-8/A minimum hacmidir.
                    _ek8a_depo_712 = float(st.session_state.get("yangin_75_ek8a_depo_min_hacim_m3", 0.0) or 0.0)
                    _toplam_hacim_m3 = _ek8a_depo_712
                    _toplam_hacim_litre = _toplam_hacim_m3 * 1000.0

                    st.session_state["yangin_712_depo_gerekli_hacim_litre"] = float(_toplam_hacim_litre)
                    st.session_state["yangin_712_depo_gerekli_hacim_m3"] = float(_toplam_hacim_m3)
                    st.session_state["yangin_712_q73_ldak"] = _q73
                    st.session_state["yangin_712_t73_dak"] = _t73
                    st.session_state["yangin_712_q74_ldak"] = _q74
                    st.session_state["yangin_712_t74_dak"] = _t74
                    st.session_state["yangin_712_q75_ldak"] = _q75
                    st.session_state["yangin_712_t75_dak"] = 0.0
                    st.session_state["yangin_712_toplam_tasarim_debisi_ldak"] = float(_toplam_tasarim_debisi_ldak)

                    st.markdown("### Otomatik Aktarılan Debiler ve Yangın Suyu Depo Hacmi")
                    st.markdown(
                        f'<div style="background-color:#FFF2CC; border:2px solid #D6B656; padding:12px 14px; border-radius:5px;">'
                        f'<b>7.3.2 Yangın Dolabı Debisi:</b> {_q73:,.2f} L/dak<br>'
                        f'<b>7.4.2 Hidrant Debisi:</b> {_q74:,.2f} L/dak<br>'
                        f'<b>7.5.2 Sprinkler Debisi:</b> {_q75:,.2f} L/dak = tasarım yoğunluğu × koruma alanı<br>'
                        f'<hr>'
                        f'<b>TOPLAM TASARIM DEBİSİ:</b> <b>{_toplam_tasarim_debisi_ldak:,.2f} L/dak</b><br>'
                        f'<span style="font-size:13px;">Bu toplam debi ileride pompa seçim hesabında kullanılacaktır; depo hacmine eklenmez.</span><br><br>'
                        f'<b>EK-8/A SU DEPOSU MİNİMUM HACMİ:</b> <b>{_toplam_hacim_litre:,.0f} L ({_toplam_hacim_m3:,.3f} m³)</b><br>'
                        f'<span style="font-size:13px;">Depo hacmi yalnızca 7.5.2 Ek-8/A otomatik seçimine göre belirlenir.</span><br><hr>'
                        f'<b>GEREKLİ MİNİMUM YANGIN SUYU DEPO HACMİ:</b> <b>{_toplam_hacim_litre:,.0f} L ({_toplam_hacim_m3:,.3f} m³)</b>'
                        f'</div>'.replace(',', 'X').replace('.', ',').replace('X', '.'),
                        unsafe_allow_html=True,
                    )

                    # Sıhhi soğuk su deposundaki seçim mantığı birebir: en yakın standart
                    # kapasite otomatik seçilir, kullanıcı isterse kapasite/poz üzerinde son onay verebilir.
                    _yangin_depo_tipleri = ["Galvaniz Modüler Su Deposu", "Betonarme Su Deposu"]
                    _yangin_depo_tipi = st.selectbox(
                        "Yangın Suyu Deposu Tipi", _yangin_depo_tipleri,
                        index=_yangin_depo_tipleri.index(st.session_state.get("yangin_712_depo_tipi", "Galvaniz Modüler Su Deposu"))
                        if st.session_state.get("yangin_712_depo_tipi", "Galvaniz Modüler Su Deposu") in _yangin_depo_tipleri else 0,
                        key="yangin_712_depo_tipi",
                    )

                    _galvaniz_kapasiteleri = [
                        (1.25, "25.150.1301"), (2.50, "25.150.1302"), (3.75, "25.150.1303"),
                        (5.00, "25.150.1304"), (6.25, "25.150.1305"), (7.50, "25.150.1306"),
                        (10.0, "25.150.1307"), (12.5, "25.150.1308"), (15.0, "25.150.1309"),
                        (20.0, "25.150.1310"), (22.5, "25.150.1311"), (25.0, "25.150.1312"),
                        (30.0, "25.150.1313"), (37.5, "25.150.1314"), (40.0, "25.150.1315"),
                        (45.0, "25.150.1316"), (50.0, "25.150.1317"), (56.0, "25.150.1318"),
                        (59.6, "25.150.1319"), (62.0, "25.150.1320"), (75.0, "25.150.1321"),
                    ]

                    def _712_en_yakin_kapasite(kayitlar, hedef):
                        return min(kayitlar, key=lambda x: abs(x[0] - hedef)) if kayitlar else None

                    if _yangin_depo_tipi == "Galvaniz Modüler Su Deposu":
                        _oto712 = _712_en_yakin_kapasite(_galvaniz_kapasiteleri, _toplam_hacim_m3)
                        if _oto712:
                            _oto_kapasite, _oto_poz = _oto712
                            st.session_state.setdefault("yangin_712_onceki_otomatik_kapasite", float(_oto_kapasite))
                            if st.session_state.get("yangin_712_depo_kapasitesi_m3") is None or st.session_state.get("yangin_712_depo_kapasitesi_m3") == st.session_state.get("yangin_712_onceki_otomatik_kapasite"):
                                st.session_state["yangin_712_depo_kapasitesi_m3"] = float(_oto_kapasite)
                            st.session_state["yangin_712_onceki_otomatik_kapasite"] = float(_oto_kapasite)
                            st.caption(f"Sıhhi su deposu mantığıyla en yakın standart kapasite: {_oto_kapasite:g} m³ (hesaplanan: {_toplam_hacim_m3:g} m³)")
                            _kap712 = st.number_input(
                                "Galvaniz modüler su deposu kapasitesi (m³) — otomatik, elle değiştirilebilir",
                                min_value=0.001, step=0.5, key="yangin_712_depo_kapasitesi_m3",
                            )
                            _uygun712 = _712_en_yakin_kapasite(_galvaniz_kapasiteleri, _kap712)
                            _poz712_oto = _uygun712[1] if _uygun712 else ""
                            _poz712_elle = st.checkbox("Galvaniz su deposu poz numarasını elle düzenle", key="yangin_712_poz_elle")
                            if _poz712_elle:
                                st.session_state.setdefault("yangin_712_poz", _poz712_oto)
                                _poz712 = st.text_input("Galvaniz su deposu Cihaz Poz No", key="yangin_712_poz").strip()
                                _poz_kap712 = next((c for c,p in _galvaniz_kapasiteleri if p == _poz712), None)
                                _goster712 = float(_poz_kap712) if _poz_kap712 is not None else float(_kap712)
                            else:
                                _poz712 = _poz712_oto
                                _goster712 = float(_kap712)
                            st.session_state["yangin_712_depo_poz"] = _poz712
                            st.session_state["yangin_712_depo_kapasitesi_secili_m3"] = _goster712
                            st.markdown(f"**Cihaz Poz No:** {_poz712} **(Kapasite: {_goster712:g} m³)**")
                        else:
                            st.warning("Hesaplanan hacim, galvaniz depo kapasite listesinin üzerindedir.")
                            st.session_state["yangin_712_depo_poz"] = ""
                            st.session_state["yangin_712_depo_kapasitesi_secili_m3"] = _toplam_hacim_m3
                    else:
                        st.info("Betonarme su deposu seçildi. Poz numarası uygulanmaz; gerekli hacim proje hacmi olarak alınır.")
                        _beton_kapasite = st.number_input(
                            "Betonarme su deposu kapasitesi (m³) — otomatik gerekli hacim",
                            min_value=0.001, value=float(_toplam_hacim_m3), step=0.5, key="yangin_712_beton_kapasitesi_m3",
                        )
                        st.session_state["yangin_712_depo_poz"] = ""
                        st.session_state["yangin_712_depo_kapasitesi_secili_m3"] = float(_beton_kapasite)
                        st.markdown(f"**Gerekli betonarme depo hacmi:** {_beton_kapasite:g} m³")

                    st.caption("Not: Galvaniz modüler depo kapasitesi ve poz seçimi, sıhhi kullanma soğuk su deposundaki en yakın standart kapasite mantığıyla yapılır. Galvaniz poz listesi mevcut sıhhi modüldeki listeyle aynıdır.")

            # 7.1 alt maddeleri yukarıda gerçek expander arayüzleriyle oluşturuldu.
            # Tekrar aşağıda statik başlık/placeholder üretmeyelim.
            if not _baslik.startswith("7.1") and not _baslik.startswith("7.2"):
                for _alt_key, _alt_rapor_key, _alt_baslik in _altlar:
                    st.session_state.setdefault(_alt_key, True)
                    st.session_state.setdefault(_alt_rapor_key, True)
                    if st.session_state.get(_alt_key, True):
                        st.markdown(f"**{_alt_baslik}**")
                        st.caption("Bu alt bölümün gerçek hesap/seçim ekranı sonraki geliştirme adımında eklenecektir.")
