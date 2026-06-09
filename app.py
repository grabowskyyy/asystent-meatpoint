# ==============================================================================
# 🚀 ZAKŁADKA 1: INTELIGENTNY GENERATOR MULTIMODALNY (TRANSSKRYPCJA + GLOBALNE ZAŁĄCZNIKI)
# ==============================================================================
with tab1:
    st.title("🐾 Generator opisów wizyt")
    st.markdown("Wklej transkrypcję i dołącz dowolne załączniki (notatki, stare plany, wyniki badań w PDF lub zdjęciach). AI inteligentnie przyporządkuje informacje do odpowiednich sekcji.")
    
    col1, col2 = st.columns([1, 1], gap="large")
    
    with col1:
        transcript = st.text_area("🔊 Wklej tutaj kompletną transkrypcję z rozmowy:", height=380, key="surowy_wklejony_tekst")
        
        # Uniwersalny uploader na wszelkiego rodzaju dokumenty od opiekuna
        zalaczniki = st.file_uploader(
            "📂 Dołącz załączniki (PDF, notatki, zdjęcia dokumentacji, wyniki badań itp.):", 
            type=["pdf", "png", "jpg", "jpeg"], 
            accept_multiple_files=True,
            key="pliki_kliniczne"
        )
        if zalaczniki:
            st.success(f"📎 Pomyślnie załadowano załączniki kontekstowe: {len(zalaczniki)} szt.")
        
    with col2:
        st.subheader("📋 Opis wizyty")
        LINK_DO_ARKUSZA = "https://docs.google.com/spreadsheets/d/1qgSX_t4_fb36CqtFUluPDKDQILpR9_SLOlYBPTXSTes/edit?usp=sharing"
        
        if st.button("🚀 Wygeneruj i uzupełnij opisy wizyty w Word", type="primary", use_container_width=True):
            if not api_key or not transcript: 
                st.error("❌ Uzupełnij klucz API oraz upewnij się, że okno transkrypcji nie jest puste!")
            else:
                with st.spinner("Globalna analiza transkrypcji i załączników oraz dopasowywanie nagłówków..."):
                    try:
                        csv_url = LINK_DO_ARKUSZA.replace('/edit?usp=sharing', '/export?format=csv')
                        df = pd.read_csv(csv_url); l_p = ""
                        for _, r in df.iterrows(): l_p += f"- Link: {r['URL']} | Tytuł: {r['Nazwa']} | Kiedy dołączyć (Wskazanie): {r['Opis dla AI']}\n"
                        
                        genai.configure(api_key=api_key)
                        
                        m = genai.GenerativeModel(
                            model_name=model_choice, 
                            system_instruction=(
                                "Jesteś doświadczonym, pedantycznym asystentem klinicznym dla dietetyk Anny Michalskiej. "
                                "Twoim zadaniem jest stworzenie jednego, spójnego protokołu na podstawie dwóch źródeł: ustnej transkrypcji oraz przesłanych dokumentów/zdjęć (załączników).\n\n"
                                "ZASADA INTELIGENTNEGO DOPASOWANIA (CROSS-ANALYSIS):\n"
                                "1. Przeanalizuj treść każdego załącznika. Informacje w nich zawarte mogą dotyczyć DOWOLNEJ sekcji protokołu (notatki o wodzie, uwagi o smaczkach, dawki leków, opisy samopoczucia, wyniki badań).\n"
                                "2. NIE wrzucaj wszystkiego z załączników do sekcji 'Aktualne badania'. Przyporządkuj fakty tematycznie: informacje o diecie komercyjnej do 'Karmy komercyjne', informacje o dawkowaniu wody do 'Piciu/Jakiej wody używać', wyniki krwi do 'Aktualne badania', a opisy dolegliwości do 'Powód konsultacji' lub 'Kał/Biegunka/Wymioty'.\n"
                                "3. Zintegruj wiedzę z transkrypcji i załączników. Jeśli dokumenty i transkrypcja mówią o tym samym, połącz te fakty w spójny, medyczny opis.\n\n"
                                "ZASADY OGÓLNE:\n"
                                "- Pisz WYŁĄCZNIE absolutną prawdę na podstawie dostarczonych materiałów. ZAKAZ zmyślania faktów czy dawek.\n"
                                "- Jeśli chcesz coś wyróżnić medycznie, używaj podwójnych gwiazdek **tekst**.\n"
                                "- Jeśli w obu źródłach brakuje danych dla danej sekcji, wstaw fragment [BRAK INFORMACJI]."
                            )
                        )
                        
                        instrukcja_szablonu = ""
                        for naglowek in STRUKTURA_PROTOKOLU:
                            if naglowek == "Załączniki:":
                                instrukcja_szablonu += f"## {naglowek}\n- Dołącz wyłącznie pasujące linki z bazy, jeśli ich warunki kliniczne zostały spełnione.\n- Pod nimi dodaj dokładnie te słowa:\nW razie pytań dotyczących tego opisu, jestem do Państwa dyspozycji.\nZachęcamy również do poszerzenia wiedzy o diecie na naszej stronie meatpoint.io lub Facebooku https://www.facebook.com/meatpoint.io\n\nPozdrawiam serdecznie,\nAnna Michalska"
                            elif naglowek == "Tyndalizacja:":
                                instrukcja_szablonu += f"## {naglowek}\n{TEKST_TYNDALIZACJA_STALY}\n\n"
                            elif naglowek == "Inne smaczki:":
                                instrukcja_szablonu += f"## {naglowek}\n{TEKST_INNE_SMACZKI_STALY}\n\n"
                            elif naglowek == "Wprowadzanie suplementów:":
                                instrukcja_szablonu += f"## {naglowek}\n{TEKST_WPROWADZANIE_SUPLEMENTOW_STALY}\n\n"
                            elif naglowek == "Aktualne badania:":
                                instrukcja_szablonu += f"## {naglowek}\n- Wypisz parametry, wyniki i opisy badań laboratoryjnych/obrazowych (krwi, moczu, USG) znalezione w transkrypcji lub bezpośrednio w plikach załączników. Przedstaw je w postaci czytelnych punktów. Jeśli brak typowych badań laboratoryjnych w obu źródłach, wstaw [BRAK INFORMACJI].\n"
                            elif naglowek == "Badania kontrolne:":
                                instrukcja_szablonu += f"## {naglowek}\n- Przedstaw zalecane przez Annę badania kontrolne w formie czystej listy punktów wraz z przypisanymi im w transkrypcji lub dokumentach terminami. Jeśli brak, wstaw sztywno [BRAK INFORMACJI].\n"
                            else:
                                prefix = "" if naglowek.startswith("###") else "## "
                                instrukcja_szablonu += f"{prefix}{naglowek}\n- Analizuj pod kątem tego nagłówka zarówno tekst transkrypcji, jak i dołączone pliki załączników. Wyciągnij precyzyjne fakty.\n"

                        pakiety_danych_dla_ai = []
                        
                        # Mapowanie plików binarnego bloba dla Gemini
                        if zalaczniki:
                            for plik in zalaczniki:
                                bytes_data = plik.read()
                                pakiety_danych_dla_ai.append({
                                    "mime_type": plik.type,
                                    "data": bytes_data
                                })
                        
                        prompt_glowny = f"Przeanalizuj podaną transkrypcję wizyty oraz wszystkie dołączone pliki kontekstowe.\n\nWygeneruj dokument według tej rygorystycznej kolejności:\n\nKROK 1: Na samej górze stwórz wyrównaną DO LEWEJ linię: 'Data wizyty: DD.MM.YYYY' (wyciągnij datę z rozmowy/plików lub wstaw [BRAK INFORMACJI])\n\nKROK 2: Bezpośrednio POD DATĄ wypisz linie metryczki podstawowej (ZAKAZ używania znaków '##' na ich początku, po dwukropku ma być dokładnie jedna spacja. Dane wyciągaj z transkrypcji oraz załączników):\nDane Opiekuna: \nPacjent: \nGatunek: \nRasa: \nWiek: \nWaga: \nBCS: \nMCS: \nIlość zwierząt w domu: \nSterylizacja/kastracja: \n\nKROK 3: Pod metryczką umieść poniższe nagłówki i uzupełnij je danymi z transkrypcji oraz plików, zachowując ich identyczną wielkość liter i pisownię:\n{instrukcja_szablonu}\n\n🚨 DEDYKOWANE DOPASOWANIE LINKÓW Z ARKUSZA:\nOto dostępna baza załączników zewnętrznych:\n{l_p}\n\nPrzeanalizuj pole 'Kiedy dołączyć (Wskazanie)'. Dołącz dany adres URL do dokumentu TYLKO wtedy, gdy z transkrypcji lub przesłanych załączników wynika, że pacjent cierpi na opisaną dolegliwość. Jeśli brak dopasowania, pomiń link.\n\nTranskrypcja rozmowy:\n{transcript}"
                        
                        pakiety_danych_dla_ai.append(prompt_glowny)
                        
                        res = m.generate_content(pakiety_danych_dla_ai)
                        
                        st.text_area("Podgląd tekstu wynikowego:", value=res.text, height=350, key="podglad_gen")
                        st.download_button("📥 POBIERZ GOTOWY PLIK WORD (.DOCX)", konwertuj_do_docx(res.text), "Protokol_MeatPoint.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    except Exception as e: 
                        st.error(f"🚨 Błąd generatora: {e}")
