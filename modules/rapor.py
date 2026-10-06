# Rapor Oluştur Butonu
_proje_otomatik_kaydet()

_rapor_olustur_sidebar = st.session_state.pop("_rapor_olustur_istegi_v134", False)
if _rapor_olustur_sidebar:
  # Önceki raporu temizle; yeni rapor başarısız olursa eski dosya yanlışlıkla
  # "hazır" görünmesin.
  st.session_state.pop("_rapor_hazir_docx_v134", None)
  st.session_state.pop("_rapor_hazir_adi_v134", None)
  try:
  
    _re_sirk_rapor_kontrol = st.session_state.get("re_sirkulasyon_pompa_sonucu_v99", {})
    gecersiz_var = any(
        str(p.get("poz_durumu", "")).strip().upper() in {"UYGUN DEĞİL", "UYGUN DEGIL", "GEÇERSİZ", "GECERSIZ"}
        for p in (psp_parametreleri or {}).values()
        if isinstance(p, dict)
    ) or (
        bolum_634_aktif
        and isinstance(_re_sirk_rapor_kontrol, dict)
        and str(_re_sirk_rapor_kontrol.get("poz_durumu", "")).strip().upper() in {"UYGUN DEĞİL", "UYGUN DEGIL", "GEÇERSİZ", "GECERSIZ"}
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
            yag_rapor_maddeleri = []
            for i, madde in enumerate(yag_ayirici_maddeleri, start=1):
              if i <= len(yag_ayirici_secimler) and yag_ayirici_secimler[i - 1]:
                yag_rapor_maddeleri.append(madde)
  
            if ek_yag_ayirici_notu.strip():
              for _not in ek_yag_ayirici_notu.split("\n"):
                if _not.strip():
                  yag_rapor_maddeleri.append(_not.strip())
  
            if yag_rapor_maddeleri:
              _genel_esas_p = doc.add_paragraph()
              _genel_esas_r = _genel_esas_p.add_run("Yağ Ayırıcı Genel Esasları")
              _genel_esas_r.bold = True
              _genel_esas_r.italic = True
              _genel_esas_r.font.size = Pt(12)
              _genel_esas_r.font.color.rgb = RGBColor(68, 114, 196)
              for yam in yag_rapor_maddeleri:
                doc.add_paragraph(yam, style="List Bullet")
  
            # Excel'deki her YA sayfası için bağımsız hesap raporu.
            for _ya_index, (_ya, _hesap) in enumerate(yag_ayirici_hesaplari.items(), start=1):
              doc.add_heading(f"6.2.3.{_ya_index} {_ya} YAĞ AYIRICISI KAPASİTE HESAPLARI:", level=3)
              doc.add_paragraph("Hesap yöntemi: EN 1825-2 standardına göre cihaz sayısına bağlı eşzamanlılık yöntemi.")
  
              # Eşzamanlılık faktörlerini Excel şablonundaki gibi ayrı ayrı göster.
              # Kullanılan adet hangi kademeye denk geliyorsa o Zi hücresi sarı vurgulanır.
              _tab = doc.add_table(rows=1, cols=10)
              _tab.style = "Table Grid"
              _hdr = _tab.rows[0].cells
              _basliklar = [
                  "Ekipman", "Adet n", "qi [L/s]", "n × qi",
                  "Zi – 1 adet", "Zi – 2 adet", "Zi – 3 adet",
                  "Zi – 4 adet", "Zi – 5+ adet", "Pis su debisi [L/s]"
              ]
              for _cell, _baslik in zip(_hdr, _basliklar):
                _cell.text = _baslik
  
              for _satir in _hesap["satirlar"]:
                _cells = _tab.add_row().cells
                _cells[0].text = _satir["ekipman"]
                _cells[1].text = str(_satir["adet"])
                _cells[2].text = f"{_satir['qi']:.2f}"
                _cells[3].text = f"{_satir['n_x_qi']:.2f}"
  
                # Excel'deki Zi(n) değerleri: 1/2/3/4/5+ adet.
                _zi_tip = next((e[2] for e in YAG_AYIRICI_EKIPMANLARI if e[0] == _satir["ekipman"]), "normal")
                _zi_degerleri = [
                    _yag_ayirici_zi(1, _zi_tip),
                    _yag_ayirici_zi(2, _zi_tip),
                    _yag_ayirici_zi(3, _zi_tip),
                    _yag_ayirici_zi(4, _zi_tip),
                    _yag_ayirici_zi(5, _zi_tip),
                ]
                for _zi_index, _zi_deger in enumerate(_zi_degerleri):
                  _zi_cell = _cells[4 + _zi_index]
                  _zi_cell.text = f"{_zi_deger:.2f}"
                  # Adet 1,2,3,4 veya 5+ kademesinin aktif olanını sarı boya.
                  _adet = int(_satir["adet"])
                  _aktif_index = min(_adet, 5) - 1
                  if _aktif_index == _zi_index:
                    for _par in _zi_cell.paragraphs:
                      for _run in _par.runs:
                        _run.font.highlight_color = WD_COLOR_INDEX.YELLOW
  
                _cells[9].text = f"{_satir['pis_su_debisi']:.2f}"
  
              _p = doc.add_paragraph()
              _r = _p.add_run(f"TOPLAM Qs = { _hesap['qs']:.2f} L/s")
              _r.bold = True
  
              # Faktörlerin seçilme nedeni raporda açıkça gösterilir.
              # Açıklamalar, YAĞ AYIRICI HESABI.xlsx şablonundaki faktör
              # eşiklerine göre oluşturulur.
              _fd_deger = _hesap["fd"]
              if abs(_fd_deger - 1.0) < 1e-9:
                  _fd_aciklama = "Yağ yoğunluğu ≤ 0.94 g/cm³ olduğu için fd = 1.00 seçilmiştir."
              else:
                  _fd_aciklama = "Yağ yoğunluğu > 0.94 g/cm³ olduğu için fd = 1.30 seçilmiştir."
  
              _ft_deger = _hesap["ft"]
              if abs(_ft_deger - 1.0) < 1e-9:
                  _ft_aciklama = "Su sıcaklığı ≤ 60 °C olduğu için ft = 1.00 seçilmiştir."
              else:
                  _ft_aciklama = "Su sıcaklığı > 60 °C olduğu için ft = 1.30 seçilmiştir."
  
              _fr_deger = _hesap["fr"]
              if abs(_fr_deger - 1.0) < 1e-9:
                  _fr_aciklama = "Temizlik malzemesi kullanılmadığı için fr = 1.00 seçilmiştir."
              elif abs(_fr_deger - 1.3) < 1e-9:
                  _fr_aciklama = "Temizlik malzemesi kullanıldığı için fr = 1.30 seçilmiştir."
              else:
                  _fr_aciklama = "Hastane kullanımı için fr = 1.50 seçilmiştir."
  
              _p = doc.add_paragraph()
              _r = _p.add_run(f"Yoğunluk Faktörü (fd): {_fd_deger:.2f}")
              _r.bold = True
              _p.add_run(f" — {_fd_aciklama}")
  
              _p = doc.add_paragraph()
              _r = _p.add_run(f"Sıcaklık Faktörü (ft): {_ft_deger:.2f}")
              _r.bold = True
              _p.add_run(f" — {_ft_aciklama}")
  
              _p = doc.add_paragraph()
              _r = _p.add_run(f"Deterjan Faktörü (fr): {_fr_deger:.2f}")
              _r.bold = True
              _p.add_run(f" — {_fr_aciklama}")
  
              _p = doc.add_paragraph()
              _r = _p.add_run(
                  "NS (Nominal Kapasite) = Qs × fd × ft × fr = "
                  f"{_hesap['qs']:.2f} × { _hesap['fd']:.2f} × { _hesap['ft']:.2f} × { _hesap['fr']:.2f} "
                  f"= { _hesap['ns']:.2f} L/s"
              )
              _r.bold = True
  
              _p = doc.add_paragraph()
              _r = _p.add_run(
                  f"Seçilen yağ ayırıcı kapasitesi: {_hesap['secilen_kapasite']:.2f} L/s"
              )
              _r.bold = True
  
              if _hesap.get("poz_rapora_aktar") and _hesap.get("secilen_poz"):
                _p = doc.add_paragraph()
                _r = _p.add_run(f"Cihaz Poz No: {_hesap['secilen_poz']['poz']}")
                _r.bold = True
                doc.add_paragraph(f"Yağ Ayırıcı Özelliği: {_hesap['secilen_poz']['tanim']}")
  
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
                        f"{_yr.get('mgm_il', '')} ili için P = {float(_yr.get('mgm_yagis_mm', 0) or 0):.1f} mm "
                        f"({_yr.get('mgm_yagis_tarih', '')})."
                    )
            elif _yr_yontem == "Ortalama Aylık Yağış Miktarı":
                doc.add_paragraph(
                    f"Tasarım yağış verisi: {_yr_yontem}. "
                    f"12 aylık ortalama yağış değerlerinin aritmetik ortalaması ile P = "
                    f"{float(_yr.get('mgm_ortalama_aylik_yagis', 0) or 0):.1f} mm alınmıştır."
                )
            else:
                doc.add_paragraph(
                    f"Tasarım yağış verisi: {_yr_yontem}. "
                    f"{_yr.get('mgm_en_yuksek_ay', '')} ayındaki en yüksek aylık ortalama yağış değeri "
                    f"P = {float(_yr.get('mgm_en_yuksek_ay_yagis', 0) or 0):.1f} mm alınmıştır."
                )
  
            # MGM meteorolojik verileri: RAPORDA YALNIZCA SEÇİLEN HESAP YÖNTEMİ GÖSTERİLİR.
            # Günlük seçildiyse yalnız günlük tablo; aylık yöntemlerden biri seçildiyse
            # yalnız seçilen aylık yöntem tablosu rapora eklenir.
            _secili_il_rapor = str(_yr.get("mgm_il", "")).strip()
            _mgm_gunluk = _yr.get("mgm_yagis_mm")
            _mgm_tarih = _yr.get("mgm_yagis_tarih", "")
            _mgm_aylik_rapor = _yr.get("mgm_aylik_yagis", {}) or {}
            _mgm_periyot = _yr.get("mgm_aylik_periyot", "") or ""
            _mgm_kaynak_url = (
                _yr.get("mgm_url", "")
                if _yr_yontem == "Günlük Toplam En Yüksek Yağış Miktarı"
                else _yr.get("mgm_aylik_url", "")
            )

            doc.add_heading("SEÇİLEN İL METEOROLOJİK VERİLERİ", level=5)

            if _yr_yontem == "Günlük Toplam En Yüksek Yağış Miktarı":
                _met_tbl = doc.add_table(rows=1, cols=3)
                _met_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                _met_tbl.autofit = True
                _mh = _met_tbl.rows[0].cells
                for _i, _baslik in enumerate(["İL", "GÜNLÜK TOPLAM EN YÜKSEK YAĞIŞ (mm)", "GERÇEKLEŞME TARİHİ"]):
                    _mh[_i].text = _baslik
                    _mh[_i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    _tcPr = _mh[_i]._tc.get_or_add_tcPr()
                    _shd = OxmlElement("w:shd")
                    _shd.set(qn("w:fill"), "D9E2F3")
                    _tcPr.append(_shd)
                    for _p in _mh[_i].paragraphs:
                        _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for _r in _p.runs:
                            _r.bold = True
                            _r.font.size = Pt(8.5)
                _mc = _met_tbl.add_row().cells
                _mc[0].text = _secili_il_rapor or "-"
                _mc[1].text = f"{float(_mgm_gunluk):.1f}" if _mgm_gunluk is not None else "Veri alınamadı"
                _mc[2].text = _mgm_tarih or "-"
                for _cell in _mc:
                    _cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for _p in _cell.paragraphs:
                        _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for _r in _p.runs:
                            _r.font.size = Pt(8.5)

            elif _yr_yontem == "Ortalama Aylık Yağış Miktarı":
                _ay_tbl = doc.add_table(rows=1, cols=3)
                _ay_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                _ay_tbl.autofit = True
                _ay_hdr = _ay_tbl.rows[0].cells
                for _i, _baslik in enumerate(["İL", "AY", "AYLIK TOPLAM YAĞIŞ ORTALAMASI (mm)"]):
                    _ay_hdr[_i].text = _baslik
                    _ay_hdr[_i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    _tcPr = _ay_hdr[_i]._tc.get_or_add_tcPr()
                    _shd = OxmlElement("w:shd")
                    _shd.set(qn("w:fill"), "D9E2F3")
                    _tcPr.append(_shd)
                    for _p in _ay_hdr[_i].paragraphs:
                        _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for _r in _p.runs:
                            _r.bold = True
                            _r.font.size = Pt(8.5)
                _aylar_sirali = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
                for _ay in _aylar_sirali:
                    if _ay not in _mgm_aylik_rapor:
                        continue
                    _c = _ay_tbl.add_row().cells
                    _c[0].text = _secili_il_rapor or "-"
                    _c[1].text = _ay
                    _c[2].text = f"{float(_mgm_aylik_rapor[_ay]):.1f}"
                    for _cell in _c:
                        _cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        for _p in _cell.paragraphs:
                            _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for _r in _p.runs:
                                _r.font.size = Pt(8.5)
                doc.add_paragraph(
                    f"Ölçüm periyodu: {_mgm_periyot or 'MGM verisinde belirtilmemiş'}. "
                    f"12 aylık değerlerin aritmetik ortalaması: "
                    f"{float(_yr.get('mgm_ortalama_aylik_yagis', 0) or 0):.1f} mm."
                )

            else:  # En Yüksek Aylık Ortalama Yağış Miktarı
                _ay_tbl = doc.add_table(rows=1, cols=4)
                _ay_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                _ay_tbl.autofit = True
                _ay_hdr = _ay_tbl.rows[0].cells
                for _i, _baslik in enumerate(["İL", "AY", "AYLIK ORTALAMA YAĞIŞ (mm)", "DURUM"]):
                    _ay_hdr[_i].text = _baslik
                    _ay_hdr[_i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    _tcPr = _ay_hdr[_i]._tc.get_or_add_tcPr()
                    _shd = OxmlElement("w:shd")
                    _shd.set(qn("w:fill"), "D9E2F3")
                    _tcPr.append(_shd)
                    for _p in _ay_hdr[_i].paragraphs:
                        _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for _r in _p.runs:
                            _r.bold = True
                            _r.font.size = Pt(8.5)
                _aylar_sirali = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
                _en_ay = _yr.get("mgm_en_yuksek_ay", "")
                for _ay in _aylar_sirali:
                    if _ay not in _mgm_aylik_rapor:
                        continue
                    _c = _ay_tbl.add_row().cells
                    _c[0].text = _secili_il_rapor or "-"
                    _c[1].text = _ay
                    _c[2].text = f"{float(_mgm_aylik_rapor[_ay]):.1f}"
                    _c[3].text = "SEÇİLEN / EN YÜKSEK AY" if _ay == _en_ay else ""
                    for _cell in _c:
                        _cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        for _p in _cell.paragraphs:
                            _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for _r in _p.runs:
                                _r.font.size = Pt(8.5)
                    if _ay == _en_ay:
                        for _cell in _c:
                            _tcPr = _cell._tc.get_or_add_tcPr()
                            _shd = OxmlElement("w:shd")
                            _shd.set(qn("w:fill"), "FFF2CC")
                            _tcPr.append(_shd)
                            for _p in _cell.paragraphs:
                                for _r in _p.runs:
                                    _r.bold = True
                doc.add_paragraph(
                    f"Ölçüm periyodu: {_mgm_periyot or 'MGM verisinde belirtilmemiş'}. "
                    f"Seçilen en yüksek aylık ortalama yağış: {_en_ay or '-'} = "
                    f"{float(_yr.get('mgm_en_yuksek_ay_yagis', 0) or 0):.1f} mm."
                )

            # Kaynak linki: raporu inceleyen kişi doğrudan kullanılan MGM sayfasına gidebilir.
            if _mgm_kaynak_url:
                _p_kaynak = doc.add_paragraph()
                _p_kaynak.add_run("Veri kaynağı: ")
                _part = _p_kaynak.part
                _rid = _part.relate_to(_mgm_kaynak_url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
                _hyperlink = OxmlElement("w:hyperlink")
                _hyperlink.set(qn("r:id"), _rid)
                _run = OxmlElement("w:r")
                _rPr = OxmlElement("w:rPr")
                _color = OxmlElement("w:color")
                _color.set(qn("w:val"), "0563C1")
                _rPr.append(_color)
                _u = OxmlElement("w:u")
                _u.set(qn("w:val"), "single")
                _rPr.append(_u)
                _run.append(_rPr)
                _text = OxmlElement("w:t")
                _text.text = "Meteoroloji Genel Müdürlüğü (MGM) – Resmi İklim İstatistikleri"
                _run.append(_text)
                _hyperlink.append(_run)
                _p_kaynak._p.append(_hyperlink)
            else:
                doc.add_paragraph(
                    "Veri kaynağı: Meteoroloji Genel Müdürlüğü (MGM), Resmi İklim İstatistikleri."
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
                f"Yıllık toplam yağış: {float(_yr.get('mgm_yillik_yagis_mm', 0) or 0):.2f} mm "
                "(MGM aylık ortalama yağışlarının toplamı)"
            )
            doc.add_paragraph(
                f"Yıllık toplanabilir yağış hacmi: {float(_yr.get('yillik_toplam_hacim_m3', 0) or 0):.2f} m³/yıl"
            )
            doc.add_paragraph(
                f"Depolanacak oran: %{float(_yr.get('depolama_orani', 6) or 0):.0f}"
            )
            doc.add_paragraph(
                "V_yıllık = A × P_yıllık × C / 1000"
            )
            doc.add_paragraph(
                f"V_yıllık = {float(_yr.get('cati_alani', 0) or 0):.2f} × "
                f"{float(_yr.get('mgm_yillik_yagis_mm', 0) or 0):.2f} × "
                f"{float(_yr.get('akis_katsayisi', 0) or 0):.2f} / 1000 = "
                f"{float(_yr.get('yillik_toplam_hacim_m3', 0) or 0):.2f} m³/yıl"
            )
            doc.add_paragraph(
                f"V_depo = V_yıllık × %{float(_yr.get('depolama_orani', 6) or 0):.0f} = "
                f"{float(_yr.get('yillik_toplam_hacim_m3', 0) or 0):.2f} × "
                f"%{float(_yr.get('depolama_orani', 6) or 0):.0f} = "
                f"{float(_yr.get('gerekli_depo', 0) or 0):.2f} m³"
            )
            doc.add_paragraph(f"Hesaplanan gerekli depo hacmi: {float(_yr.get('gerekli_depo', 0) or 0):.2f} m³")
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
                f"Yağmur suyu toplama alanı: {float(_yr.get('cati_alani', 0) or 0):.2f} m²"
            )
            doc.add_paragraph(
                f"Filtre adedi: {float(_yr.get('filtre_adet', 1) or 0):.0f} adet"
            )
            doc.add_paragraph(
                f"Filtre başına düşen toplama alanı: {float(_yr.get('filtre_basina_alan_m2', 0) or 0):.2f} m²"
            )
            doc.add_paragraph(
                f"Tek filtre kapasitesi: {float(_yr.get('filtre_kapasite_m2', 0) or 0):.0f} m²/adet"
            )
            doc.add_paragraph(
                f"Toplam filtre kapasitesi: {_yr.get('filtre_toplam_kapasite_m2', _yr.get('filtre_kapasite_m2', 0)):,.0f} m²"
            )
            doc.add_paragraph(
                f"Tek filtre maksimum debisi: {float(_yr.get('filtre_debisi_ls', 0) or 0):.0f} L/s; "
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
            # TAŞMA HATTI HİDROLİK KONTROL TABLOSU
            # Kullanıcı tarafından onaylanan sabit referans değerleri raporda
            # metin/paragraf yerine gerçek Word tablosu olarak gösterilir.
            _tasma_rapor_tablo = [
                ("DN 50", "0.96", "0.96", "0.49", "YETERSİZ"),
                ("DN 65", "1.94", "1.94", "0.58", "YETERSİZ"),
                ("DN 80", "3.37", "3.37", "0.67", "YETERSİZ"),
                ("DN 100", "6.10", "6.10", "0.78", "YETERSİZ"),
                ("DN 125", "11.07", "11.07", "0.90", "YETERSİZ"),
                ("DN 150", "18.00", "18.00", "1.02", "YETERSİZ"),
                ("DN 200", "94.25", "38.76", "1.23", "YETERSİZ"),
                ("DN 250", "147.26", "70.28", "1.43", "YETERSİZ"),
            ]
            doc.add_paragraph("Kontrol edilen çaplar:")
            _tasma_tbl = doc.add_table(rows=1, cols=5)
            _tasma_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            _tasma_tbl.autofit = True
            _hdr = _tasma_tbl.rows[0].cells
            for _i, _baslik in enumerate([
                "BORU ÇAPI", "TASARIM KAPASİTESİ (L/s)",
                "MANNING KAPASİTESİ (L/s)", "MANNING HIZI (m/s)", "SONUÇ"
            ]):
                _hdr[_i].text = _baslik
                _hdr[_i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                _tcPr = _hdr[_i]._tc.get_or_add_tcPr()
                _shd = OxmlElement("w:shd")
                _shd.set(qn("w:fill"), "D9E2F3")
                _tcPr.append(_shd)
                for _p in _hdr[_i].paragraphs:
                    _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for _r in _p.runs:
                        _r.bold = True
                        _r.font.size = Pt(8.5)
            for _satir in _tasma_rapor_tablo:
                _cells = _tasma_tbl.add_row().cells
                for _i, _deger in enumerate(_satir):
                    _cells[_i].text = _deger
                    _cells[_i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for _p in _cells[_i].paragraphs:
                        _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for _r in _p.runs:
                            _r.font.size = Pt(8.5)
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
  
  
      # --- 6.3.5 SU YUMUŞATMA CİHAZI SEÇİMİ ---
      if bolum_635_aktif:
          doc.add_heading(_63_dinamik_baslik("rapor_bolum_635"), level=2)
          _yum_rapor_maddeleri = [m for i, m in enumerate(yumusatma_maddeleri) if i < len(yumusatma_secimler) and yumusatma_secimler[i]]
          if ek_yumusatma_notu.strip():
              _yum_rapor_maddeleri.extend(x.strip() for x in ek_yumusatma_notu.split("\n") if x.strip())
          if _yum_rapor_maddeleri:
              _p_yum = doc.add_paragraph()
              _r_yum = _p_yum.add_run("Yumuşatma Cihazı Seçimi Genel Esasları")
              _r_yum.bold = True
              _r_yum.italic = True
              _r_yum.font.size = Pt(12)
              _r_yum.font.color.rgb = RGBColor(68, 114, 196)
              for _m in _yum_rapor_maddeleri:
                  doc.add_paragraph(_m, style="List Bullet")
  
          _yum_sonuc = st.session_state.get("yumusatma_sonucu", {})
          if _yum_sonuc:
              _vh = float(_yum_sonuc.get("sistem_hacmi_m3", 0.0))
              _vh_gen = float(_yum_sonuc.get("genlesme_hacmi_m3", 0.0))
              _ts = float(_yum_sonuc.get("doldurma_suresi_h", 6.0))
              _qd = float(_yum_sonuc.get("gerekli_debi_m3h", 0.0))
              _kap = float(_yum_sonuc.get("kapasite", 0.0))
              _poz = str(_yum_sonuc.get("poz_no", ""))
              _adet = int(_yum_sonuc.get("adet", 1) or 1)
              _recine = float(_yum_sonuc.get("recine_l", 0.0))
              _baglanti = str(_yum_sonuc.get("baglanti", ""))
              _reg_kap = float(_yum_sonuc.get("rej_kapasitesi_m3_reg", 0.0))
              _sertlik = float(_yum_sonuc.get("toplam_sertlik_fr_m3_reg", 0.0))
  
              _p_yum_kapasite = doc.add_paragraph()
              _r_yum_kapasite = _p_yum_kapasite.add_run("Yumuşatma cihazı kapasite hesabı:")
              _r_yum_kapasite.bold = True
              _r_yum_kapasite.italic = True
              _r_yum_kapasite.font.name = "Times New Roman"
              _r_yum_kapasite.font.size = Pt(12)
              _r_yum_kapasite.font.color.rgb = RGBColor(68, 114, 196)
  
              _vh_litre = _vh * 1000.0
              _vh_gen_litre = _vh_gen * 1000.0
              doc.add_paragraph(
                  f"Sistemdeki su hacmi: V = {_vh_litre:.0f} lt - {_vh:.2f} m³"
                  + (f" (Kapalı genleşme deposu hesabından alınan değer: {_vh_gen_litre:.0f} lt - {_vh_gen:.2f} m³)" if _vh_gen > 0 else " (kullanıcı tarafından girilen değer)")
              )
              doc.add_paragraph(f"Sistemin doldurma süresi: t = {_ts:.2f} saat")
              doc.add_paragraph(f"Gerekli yumuşatma debisi: Q = V / t = {_vh:.2f} / {_ts:.2f} = {_qd:.2f} m³/h")
              _tip_rapor_doc = "ikili tam otomatik tandem" if _yum_sonuc.get("sistem_tipi") == "İkili Tandem" else "tam otomatik"
              doc.add_paragraph(
                  f"Sonuç: {_kap:.2f} m³/h'lik {_tip_rapor_doc} tip su yumuşatma cihazı projelendirilmiştir."
              )
  
              _yum_tbl = doc.add_table(rows=0, cols=2)
              _yum_tbl.style = "Table Grid"
              _yum_satirlar = [
                  ("Sistem Tipi", str(_yum_sonuc.get("sistem_tipi", ""))),
                  ("Cihaz Adedi", f"{_adet} adet"),
                  ("Seçilen Cihaz Kapasitesi", f"{_kap:.2f} m³/h"),
                  ("Reçine Miktarı", f"{_recine:.0f} L"),
                  ("Giriş / Çıkış Bağlantısı", _baglanti),
                  ("Rejenerasyon Kapasitesi", f"{_reg_kap:.2f} m³/reg"),
                  ("Toplam Sertlik Kapasitesi", f"{_sertlik:.0f} °Fr·m³/reg"),
              ]
              if _yum_sonuc.get("poz_rapora") and _poz:
                  _yum_satirlar.append(("Cihaz Poz No", _poz))
              for _etiket, _deger in _yum_satirlar:
                  _cells = _yum_tbl.add_row().cells
                  _cells[0].text = _etiket
                  _cells[1].text = str(_deger)
  
      # -----------------------------------------------------------------------
      # 7. YANGIN TESİSATI - 7.1 RAPORU
      # -----------------------------------------------------------------------
      if st.session_state.get("rapor_bolum_7", True) and st.session_state.get("rapor_bolum_71", True):
          ana_baslik_ekle("7. YANGIN TESİSATI")
          doc.add_heading("7.1 YANGIN TESİSATI GENEL ESASLARI", level=2)
          for _rb, _rm, _rk, _rrk in _yangin_71_gruplari:
              if not st.session_state.get(_rrk, True):
                  continue
              vals = st.session_state.get(_rk, [True] * len(_rm))
              if not any(vals):
                  continue
              doc.add_heading(_rb, level=3)
              for i, madde in enumerate(_rm):
                  if i < len(vals) and vals[i]:
                      p = doc.add_paragraph()
                      r = p.add_run("• " + madde)
                      r.bold = True
  
          if st.session_state.get("rapor_bolum_713", True):
              _ts_vals = [bool(st.session_state.get(f"yangin_713_ts_{_i}", True)) for _i in range(len(_yangin_713_ts_standartlari))]
              _nfpa_vals = [bool(st.session_state.get(f"yangin_713_nfpa_{_i}", True)) for _i in range(len(_yangin_713_nfpa_standartlari))]
              _std_vals = _ts_vals + _nfpa_vals
              if any(_std_vals):
                  doc.add_heading("7.1.3 YANGIN TESİSATI STANDARTLARI", level=3)
                  for _i, _std in enumerate(_yangin_713_standartlari):
                      if _std_vals[_i]:
                          p = doc.add_paragraph()
                          r = p.add_run("• " + _std)
                          r.bold = True
  
      # -----------------------------------------------------------------------
      # 7.2 RAPORU - Ek-1/B tablosu + 7.2.2 otomatik/manuel sonuç
      # 7.2, 7.1 raporundan sonra yazılır; böylece Word raporunda bölüm sırası
      # program ekranındaki 7.1 -> 7.2 akışıyla birebir aynı olur.
      # -----------------------------------------------------------------------
      _r72_parent = bool(st.session_state.get("rapor_bolum_72", True))
      _r721 = bool(st.session_state.get("rapor_bolum_721", True))
      _r722 = bool(st.session_state.get("rapor_bolum_722", True))
  
      if st.session_state.get("rapor_bolum_7", True) and _r72_parent and (_r721 or _r722):
          if not st.session_state.get("rapor_bolum_71", True):
              ana_baslik_ekle("7. YANGIN TESİSATI")
  
          doc.add_heading("7.2 YANGIN TEHLİKE SINIFI VE TASARIM KRİTERLERİ", level=2)

          # GENEL BİNA BİLGİLERİ - Yangın modülünde girilen ortak verileri
          # rapora tablo olarak aktar. Sıhhi Tesisat tarafına dokunulmaz.
          _gbi_rapor_verileri = [
              ("Toplam yapı / kapalı kullanım alanı", f"{float(st.session_state.get('yangin_genel_toplam_alan_m2', 0.0) or 0.0):,.2f} m²"),
              ("Kat sayısı", str(int(st.session_state.get('yangin_genel_kat_sayisi', 0) or 0))),
              ("Bodrum kat sayısı", str(int(st.session_state.get('yangin_genel_bodrum_kat_sayisi', 0) or 0))),
              ("Bina yüksekliği", f"{float(st.session_state.get('yangin_genel_bina_yuksekligi_m', 0.0) or 0.0):,.2f} m"),
              ("Yapı yüksekliği", f"{float(st.session_state.get('yangin_genel_yapi_yuksekligi_m', 0.0) or 0.0):,.2f} m"),
              ("Merdiven kovası yüksekliği", f"{float(st.session_state.get('yangin_genel_merdiven_kovasi_yuksekligi_m', 0.0) or 0.0):,.2f} m"),
              ("Toplam kişi sayısı", str(int(st.session_state.get('yangin_genel_kisi_sayisi', 0) or 0))),
              ("Otopark araç kapasitesi", str(int(st.session_state.get('yangin_genel_otopark_arac_sayisi', 0) or 0))),
              ("Kapalı otopark alanı", f"{float(st.session_state.get('yangin_genel_kapali_otopark_alan_m2', 0.0) or 0.0):,.2f} m²"),
              ("Yatak sayısı", str(int(st.session_state.get('yangin_genel_yatak_sayisi', 0) or 0))),
              ("İmar / yerleşim alanı", f"{float(st.session_state.get('yangin_genel_imar_alani_m2', 0.0) or 0.0):,.2f} m²"),
              ("Acil durum asansörü", "VAR" if bool(st.session_state.get('yangin_genel_acil_durum_asansoru', False)) else "YOK"),
          ]
          _gbi_baslik = doc.add_paragraph()
          _gbi_run = _gbi_baslik.add_run("GENEL BİNA BİLGİLERİ")
          _gbi_run.bold = True
          _gbi_run.italic = True
          _gbi_run.font.size = Pt(13)
          _gbi_tbl = doc.add_table(rows=1, cols=2)
          _gbi_tbl.style = "Table Grid"
          _gbi_tbl.rows[0].cells[0].text = "BİNA BİLGİSİ"
          _gbi_tbl.rows[0].cells[1].text = "DEĞER"
          for _cell in _gbi_tbl.rows[0].cells:
              for _run in _cell.paragraphs[0].runs:
                  _run.bold = True
          for _etiket, _deger in _gbi_rapor_verileri:
              _gc = _gbi_tbl.add_row().cells
              _gc[0].text = _etiket
              _gc[1].text = _deger
  
  
          _secili_kayitlar = _yangin_721_secili_kayitlar()
          _otomatik = _ek1b_otomatik_sinif(_secili_kayitlar)
          _manuel = bool(st.session_state.get("yangin_721_manuel", False))
          _etkin = (
              st.session_state.get("yangin_721_manuel_sinif", _otomatik)
              if _manuel else _otomatik
          )
  
          if _r721:
              doc.add_heading("7.2.1 BİNA KULLANIM AMACI", level=3)
              if _secili_kayitlar:
                  doc.add_paragraph(
                      "Seçilen bina / kullanım alanları: " +
                      ", ".join(x["etiket"] for x in _secili_kayitlar)
                  )
                  doc.add_paragraph(
                      f"Ek-1/B + Ek-1/C'ye göre otomatik yangın tehlike sınıfı (en yüksek seçilen sınıf): {_otomatik}"
                  )
              else:
                  doc.add_paragraph("Ek-1/B kullanım alanı seçilmemiştir.")
  
              # RAPORDA EK-1/B: yalnızca seçilen satırlar gösterilir;
              # ancak seçilen satırın TÜM tehlike sınıfı sütunları korunur.
              # Seçilen hücre(ler) sarı renkle vurgulanır.
              _tbl = doc.add_table(rows=1, cols=len(_ek1b_basliklari) + 1)
              _tbl.style = "Table Grid"
              _hdr = _tbl.rows[0].cells
              _hdr[0].text = "KULLANIM TÜRÜ"
              for _i, _h in enumerate(_ek1b_basliklari, start=1):
                  _hdr[_i].text = str(_h)
              for _cell in _hdr:
                  for _r in _cell.paragraphs[0].runs:
                      _r.bold = True

              # Aynı satırda birden fazla sarı hücre seçilmişse tek satırda birleştir.
              _secili_b_satirlar = {}
              for _kayit in _secili_kayitlar:
                  _tur = _kayit.get("satır", _kayit.get("satir", ""))
                  _kolon = int(_kayit.get("kolon", 0) or 0)
                  _secili_b_satirlar.setdefault(_tur, set()).add(_kolon)

              for _tur, _secili_kolonlar in _secili_b_satirlar.items():
                  _satir_verisi = next((x for x in _ek1b_satirlari if x.get("tur") == _tur), None)
                  if _satir_verisi is None:
                      continue
                  _hucreler = _satir_verisi.get("hücreler", [])
                  _cells = _tbl.add_row().cells
                  _cells[0].text = str(_tur)
                  for _kolon in range(len(_ek1b_basliklari)):
                      _metin = _hucreler[_kolon] if _kolon < len(_hucreler) else ""
                      _cells[_kolon + 1].text = str(_metin)
                      if _kolon in _secili_kolonlar:
                          _tcPr = _cells[_kolon + 1]._tc.get_or_add_tcPr()
                          _shd = OxmlElement("w:shd")
                          _shd.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill", "FFF2CC")
                          _tcPr.append(_shd)
                          for _pr in _cells[_kolon + 1].paragraphs:
                              for _run in _pr.runs:
                                  _run.bold = True

              _pnot = doc.add_paragraph()
              _rn = _pnot.add_run(
                  "Raporda yalnızca seçilen Ek-1/B kullanım alanları gösterilmiştir."
              )
              _rn.italic = True

              _pkaynak = doc.add_paragraph()
              _rs = _pkaynak.add_run(
                  "Kaynak / Tablo: Binaların Yangından Korunması Hakkında Yönetmelik Kılavuzu — "
                  "Ek-1/B: Orta Tehlike Kullanım Alanları, s. 243."
              )
              _rs.bold = True

              # Ek-1/C: yalnızca seçilen satırlar gösterilir; satırın TÜM
              # tehlike sınıfı sütunları korunur, seçilen hücre(ler) sarı vurgulanır.
              _secili_c = _ek1c_secili_kayitlar()
              if _secili_c:
                  doc.add_paragraph("Ek-1/C — Yüksek Tehlike Kullanım Alanları")
                  _tc = doc.add_table(rows=1, cols=len(_ek1c_basliklari) + 1)
                  _tc.style = "Table Grid"
                  _tc.rows[0].cells[0].text = "KULLANIM TÜRÜ"
                  for _i, _h in enumerate(_ek1c_basliklari, start=1):
                      _tc.rows[0].cells[_i].text = str(_h)
                  for _cell in _tc.rows[0].cells:
                      for _r in _cell.paragraphs[0].runs:
                          _r.bold = True

                  _secili_c_satirlar = {}
                  for _kayit in _secili_c:
                      _satir_no = int(_kayit.get("satır", _kayit.get("satir", 0)) or 0)
                      _kolon = int(_kayit.get("kolon", 0) or 0)
                      _secili_c_satirlar.setdefault(_satir_no, set()).add(_kolon)

                  for _satir_no, _secili_kolonlar in _secili_c_satirlar.items():
                      _satir_verisi = _ek1c_satirlari[_satir_no] if 0 <= _satir_no < len(_ek1c_satirlari) else None
                      if _satir_verisi is None:
                          continue
                      _hucreler = _satir_verisi.get("hücreler", [])
                      _cells = _tc.add_row().cells
                      _cells[0].text = str(_satir_verisi.get("tur", ""))
                      for _kolon in range(len(_ek1c_basliklari)):
                          _metin = _hucreler[_kolon] if _kolon < len(_hucreler) else ""
                          _cells[_kolon + 1].text = str(_metin)
                          if _kolon in _secili_kolonlar:
                              _tcPr = _cells[_kolon + 1]._tc.get_or_add_tcPr()
                              _shd = OxmlElement("w:shd")
                              _shd.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill", "FFF2CC")
                              _tcPr.append(_shd)
                              for _pr in _cells[_kolon + 1].paragraphs:
                                  for _run in _pr.runs:
                                      _run.bold = True
                  _pc = doc.add_paragraph()
                  _rc = _pc.add_run("Raporda yalnızca seçilen Ek-1/C kullanım alanları gösterilmiştir.")
                  _rc.italic = True
                  _src_c = doc.add_paragraph()
                  _src_c.add_run(
                      "Kaynak / Tablo: Binaların Yangından Korunması Hakkında Yönetmelik Kılavuzu — "
                      "Ek-1/C: Yüksek Tehlike Kullanım Alanları, s. 245."
                  ).bold = True

          if _r722:
              doc.add_heading("7.2.2 YANGIN TEHLİKE SINIFI", level=3)
  
              doc.add_paragraph("Çoklu seçilen kullanım alanları ve yangın tehlike sınıfları:")
              _t722_secimler = doc.add_table(rows=1, cols=2)
              _t722_secimler.style = "Table Grid"
              _t722_secimler.rows[0].cells[0].text = "SEÇİLEN KULLANIM ALANI"
              _t722_secimler.rows[0].cells[1].text = "YANGIN TEHLİKE SINIFI"
              for _c in _t722_secimler.rows[0].cells:
                  for _pr in _c.paragraphs:
                      for _run in _pr.runs:
                          _run.bold = True
              for _kayit in _secili_kayitlar:
                  _cc = _t722_secimler.add_row().cells
                  _cc[0].text = str(_kayit.get("etiket", ""))
                  _cc[1].text = str(_kayit.get("sinif", ""))
  
              _t722 = doc.add_table(rows=0, cols=2)
              _t722.style = "Table Grid"
              _rws = [
                  ("Seçilen kullanım alanı sayısı", len(_secili_kayitlar)),
                  ("Ek-1/B + Ek-1/C otomatik sonucu (en yüksek)", _otomatik),
                  ("Seçilen yangın tehlike sınıfı", _etkin),
              ]
              for _etiket, _deger in _rws:
                  _cc = _t722.add_row().cells
                  _cc[0].text = _etiket
                  _cc[1].text = str(_deger)
                  if _etiket == "Seçilen yangın tehlike sınıfı":
                      for _cell in _cc:
                          _tcPr = _cell._tc.get_or_add_tcPr()
                          _shd = OxmlElement("w:shd")
                          _shd.set(qn("w:fill"), "FFF2CC")
                          _tcPr.append(_shd)
                          for _pr in _cell.paragraphs:
                              for _run in _pr.runs:
                                  _run.bold = True
              _src722 = doc.add_paragraph()
              _src722.add_run(
                  "Kaynak: Binaların Yangından Korunması Hakkında Yönetmelik Kılavuzu, "
                  "Ek-1/B ve Ek-1/C; tehlike sınıfının farklı bölümlerdeki kullanım alanlarına göre "
                  "en yüksek sınıfa göre belirlenmesi için Madde 19."
              ).italic = True
  
      # -----------------------------------------------------------------------
      # 7.3 - 7.15 YANGIN TESİSATI ALT BÖLÜMLERİ
      # Her başlık ayrı bir bölüm olarak tutulur; hesap içerikleri sonraki
      # aşamalarda bölüm bölüm geliştirilecektir.
      # -----------------------------------------------------------------------
      _yangin_73_715 = [
          ("bolum_73", "rapor_bolum_73", "7.3 YANGIN DOLABI SİSTEMİ TASARIMI VE HESAPLAMALARI"),
          ("bolum_74", "rapor_bolum_74", "7.4 HİDRANT SİSTEMİ TASARIMI VE HESAPLAMALARI"),
          ("bolum_75", "rapor_bolum_75", "7.5 SPRİNKLER (YAĞMURLAMA) SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_76", "rapor_bolum_76", "7.6 GAZLI SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_77", "rapor_bolum_77", "7.7 KÖPÜKLÜ SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_78", "rapor_bolum_78", "7.8 DAVLUMBAZ SÖNDÜRME SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_79", "rapor_bolum_79", "7.9 MERDİVEN BASINÇLANDIRMA SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_710", "rapor_bolum_710", "7.10 ASANSÖR BASINÇLANDIRMA SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_7110", "rapor_bolum_7110", "7.11 DUMAN TAHLİYE SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_7120", "rapor_bolum_7120", "7.12 YANGIN SUYU DEPOLAMA SİSTEMİ TASARIM VE HESAPLAMALARI"),
          ("bolum_7130", "rapor_bolum_7130", "7.13 YANGIN POMPA GRUBU TASARIM VE HESAPLAMALARI"),
          ("bolum_7140", "rapor_bolum_7140", "7.14 YANGIN TESİSATI HİDROLİK HESAPLARI"),
          ("bolum_7150", "rapor_bolum_7150", "7.15 YANGIN TESİSATI SONUÇ TABLOSU"),
      ]
      if st.session_state.get("rapor_bolum_7", True):
          _r73_715 = [x for x in _yangin_73_715 if st.session_state.get(x[1], True)]
          if _r73_715:
              _has_71_report = bool(st.session_state.get("rapor_bolum_71", True)) and (
                  any(st.session_state.get(_rrk, True) for _rb, _rm, _rk, _rrk in _yangin_71_gruplari)
                  or bool(st.session_state.get("rapor_bolum_713", True))
              )
              _has_72_report = bool(st.session_state.get("rapor_bolum_72", True)) and (
                  bool(st.session_state.get("rapor_bolum_721", True))
                  or bool(st.session_state.get("rapor_bolum_722", True))
              )
              if not (_has_71_report or _has_72_report):
                  ana_baslik_ekle("7. YANGIN TESİSATI")
              for _bk, _rk, _bt in _r73_715:
                  _p7 = doc.add_paragraph()
                  _p7.paragraph_format.space_before = Pt(12)
                  _p7.paragraph_format.space_after = Pt(6)
                  _run7 = _p7.add_run(_bt.upper())
                  _run7.bold = True
                  _run7.italic = True
                  _run7.font.size = Pt(15)
                  _run7.font.color.rgb = RGBColor(31, 78, 121)
                  _body7 = doc.add_paragraph(
                      "Bu bölümün tasarım ve hesaplama içeriği sonraki aşamada ayrı olarak geliştirilecektir."
                  )
                  _body7.paragraph_format.space_after = Pt(6)
  
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
  
      # Rapor aynı çalıştırmada oluşturuldu. Aşağıdaki blok, oluşturulan
      # dosyayı bu çalıştırmanın sonunda doğrudan indirme düğmesine bağlar.
  except Exception as _rapor_hata:
    st.error("❌ Rapor oluşturulurken hata oluştu. Aşağıdaki gerçek hata rapor oluşturma işlemini durdurdu:")
    st.exception(_rapor_hata)

# ---------------------------------------------------------------------------
# RAPOR ÇIKTISI — AYNI ÇALIŞTIRMADA İNDİRME
# ---------------------------------------------------------------------------
if st.session_state.get("_rapor_hazir_docx_v134"):
    _rapor_docx_veri_son = st.session_state["_rapor_hazir_docx_v134"]
    _rapor_ad_son = st.session_state.get("_rapor_hazir_adi_v134", "Mekanik_Uygulama_Raporu")
    _rapor_format_son = st.session_state.get("rapor_cikti_format_v134", "Word (.docx)")
    st.sidebar.markdown("### 📥 RAPOR HAZIR")
    if _rapor_format_son == "Word (.docx)":
        st.sidebar.download_button(
            "📥 Word'u İndir",
            data=_rapor_docx_veri_son,
            file_name=f"{_rapor_ad_son}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="rapor_word_indir_son_v2",
            use_container_width=True,
        )
    elif _rapor_format_son == "PDF (.pdf)":
        _pdf_son = _rapor_docx_pdf_donustur(_rapor_docx_veri_son)
        if _pdf_son:
            st.sidebar.download_button(
                "📥 PDF'yi İndir", data=_pdf_son,
                file_name=f"{_rapor_ad_son}.pdf", mime="application/pdf",
                key="rapor_pdf_indir_son_v2", use_container_width=True,
            )
        else:
            st.sidebar.warning("PDF dönüştürme motoru bu ortamda çalışmadı. Word çıktısını seçebilirsiniz.")
    elif _rapor_format_son == "HTML (.html)":
        _html_son = _rapor_docx_html(_rapor_docx_veri_son)
        st.sidebar.download_button(
            "📥 HTML'yi İndir", data=_html_son,
            file_name=f"{_rapor_ad_son}.html", mime="text/html",
            key="rapor_html_indir_son_v2", use_container_width=True,
        )
    elif _rapor_format_son == "Metin (.txt)":
        _txt_son = _rapor_docx_txt(_rapor_docx_veri_son)
        st.sidebar.download_button(
            "📥 Metni İndir", data=_txt_son,
            file_name=f"{_rapor_ad_son}.txt", mime="text/plain",
            key="rapor_txt_indir_son_v2", use_container_width=True,
        )



# SAYFA SONU ANKORU
st.markdown('<div id="sayfa_sonu"></div>', unsafe_allow_html=True)
