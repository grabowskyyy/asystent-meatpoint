import streamlit as st
import google.generativeai as genai
import os
import re
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from streamlit_mic_recorder import mic_recorder

# --- DANE SZABLONU (Skompresowane strukturalnie dla bezpieczeństwa schowka) ---
SZABLON_PRODUKCYJNY = """Data wizyty: [Wpisz datę wizyty lub BRAK INFORMACJI]
## DANE FORMALNE OPIEKUNA
- **Imię i Nazwisko:** [Imię i nazwisko opiekuna]
- **Adres zamieszkania:** [Adres zamieszkania opiekuna (ulica, kod, miasto)]
- **Numer PESEL:** [Numer PESEL opiekuna]
- **Numer telefonu:** [Numer telefonu kontaktowego]
- **Adres e-mail:** [Adres e-mail opiekuna]
## DANE PACJENTA
- **Pacjent:** [Imię zwierzęcia]
- **Gatunek:** [Kot/Pies]
- **Rasa:** [Rasa pacjenta]
- **Wiek:** [Wiek pacjenta]
- **Waga:** [Aktualna waga, tendencje wagowe i waga docelowa]
- **BCS:** [Ocena kondycji w skali 1-9/9 oraz krótki opis fizyczny sylwetki]
- **Ilość zwierząt w domu:** [Liczba zwierząt w stadzie, status pacjenta i relacje]
- **Sterylizacja/kastracja:** [Tak/Nie + rok i miejsce zabiegu]
## WYWIAD KLINICZNY
- **Powód konsultacji:** [Historia schorzenia, zdiagnozowane jednostki chorobowe np. PNN, IBD, zapalenie trzustki oraz główne oczekiwania i cele opiekuna wobec diety]
- **Aktualne samopoczucie:** [Zachowanie pacjenta, przebyte niedawno zabiegi np. sanacja jamy ustnej, stan po zabiegu]
- **Aktywność:** [Umiarkowana/duża, adekwatność do aktualnego stanu zdrowia]
- **Apetyt:** [Stan apetytu, częstotliwość podawania karmy w ciągu doby, historia wybredności]
- **Pragnienie:** [Ilość samodzielnego picia, częstotliwość, stosowane kroplówki - objętość dobowa i rodzaj płynów, plany redukcji płynoterapii]
## WYPRÓŻNIENIA I OBJAWY GASTRYCZNE
- **Kał:** [Częstotliwość na dobę, uformowanie, zapach, konsystencja, opis jelit z USG pod kątem zmian typowych dla IBD]
- **Wymioty:** [Częstotliwość występowania, po jakich pokarmach lub lekach]
- **Mocz:** [Barwa, klarowność, ciężar właściwy, proteinuria/białko, obecność erytrocytów, infekcje, dobowy schemat mikcji]
- **Odrobaczanie:** [Ostatnia data odrobaczania, powód, forma podania i preparat]
## AKTUALNE BADANIA LABORATORYJNE
- **Kreatynina:** [Wartość + jednostka + trend]
- **Mocznik:** [Wartość + jednostka + trend]
- **Fosfor:** [Wartość + jednostka + trend]
- **T4 całkowita:** [Wartość + jednostka + trend]
- **Morfologia (HGB / Anemia):** [Wartość HGB, stan układu czerwonokrwinkowego, diagnoza niedokrwistości]
- **Albuminy:** [Wartość + jednostka + trend]
- **&alpha;-amylaza:** [Wartość + jednostka + trend]
- **Cholesterol:** [Wartość + jednostka + trend]
- **WBC (Leukocyty):** [Wartość + stan zapalny/infekcja]
- **Gospodarka cukrowa (Fruktozamina):** [Wartość fruktozaminy, glukoza w moczu, wykluczenie/potwierzenie cukrzycy]
## AKTUALNE LEKI I SUPLEMENTY MEDYCZNE
[Pełna lista przyjmowanych preparatów, dawkowanie, częstotliwość i od kiedy są stosowane]
## KOMENTARZ DO WYWIADU I GŁÓWNE ZAŁOŻENIA DIETY
- **Komentarz:** [Podsumowanie stopnia trudności pacjenta, tolerancji składników i wymaganych kompromisów klinicznych między nerkami a przewodem pokarmowym]
- **Główne założenia diety:** [Kluczowe cele makroskładnikowe: poziom fosforu, jakość i strawność białka, poziomy tłuszczów i węglowodanów pod kątem trzustki, zasada stopniowego wdrażania]
## EDUKACJA OPIEKUNA: CO SIĘ ZMIENI NA DIECIE BARF/BACF
- **Częstotliwość kału:** Zwierzę może oddawać mniejszy kał i może go oddawać co 2–3 dni. Na wysokomięsnej diecie to normalne. Ważne, żeby był dobrego kształtu i konsystencji (wdł skali bristolskiej).
- **UWAGA NA ZAPARCIA:** Należy odróżnić rzadkie oddawanie kału od zaparć. Jeśli pacjent na diecie BARF będzie miał: suchą kupę, twardą, bobki / rodzynki / kamyczki, z dużą ilością włosa… to może być zaparcie lub do niego prowadzić. Nie chodzi o samą częstotliwość oddawania stolca, ale o jego wygląd i o zachowanie w kuwecie.
- **Parametry krwi:** Parametry nerkowe krwi na wysoko mięsnej diecie mogą się różnić od zdrowych zwierząt (nie tylko z powodu choroby nerek), zwłaszcza mocznik i kreatynina. W zależności od pozostałych parametrów i samopoczucia - nie oznacza od razu pogorszenia choroby nerek. Ważna jest stała kontrola u nefrologa: badanie USG, SDMA, badania moczu i stanu ogólnego, być może FGF-23 - zgodnie z zaleceniami lekarza.
- **Objętość posiłku:** Początkowo może się wydawać, że diety jest mało. Dieta BARF/BACF nerkowa jest bardziej kaloryczna i treściwa w mniejszej objętości niż puszki i saszetki. Przyzwyczajanie się do tej zmniejszonej ilości może zająć ok. 2–3 miesiące i to jest normalne.
## HISTORIA ŻYWIENIOWA I PREFERENCJE SMAKOWE
- **Dotychczasowe żywienie:** [Opis dotychczasowych modeli żywienia, stosowane wcześniej przepisy, źródła białka, używane marki, stopień akceptacji i przyczyny rezygnacji/modyfikacji]
- **KATEGORYCZNIE TAK (Ulubione smaki):** [Lista akceptowanych rodzajów mięs, części tuszy, podrobów, warzyw i forma podania. UWAGA: Podkreśl czy je potrawy mrożone czy tylko świeże/z lodówki]
- **KATEGORYCZNIE NIE (Odrzucone składniki):** [Lista absolutnie odrzucanych przez zwierzę składników, mięs, form wapnia lub suplementów wywołujących wymioty, niechęć lub całkowity bunt]
## SPECYFIKACJA NOWEGO PLANU DIETETYCEDNEGO
- **Model diety:** [Model diety np. BACF domowa gotowana przygotowywana na świeżo, logistyka, częstotliwość rotacji przepisów w miesiącach]
- **Białka bazowe i dodatki:** [Wybrane gatunki mięs, podrobów oraz dozwolonych warzyw]
- **Kaloryczność próbna:** [Wartość] kcal/dzień (ustawiona w odniesieniu do dotychczasowej karmy). Warunki jednorazowej zmiany kaloryczności lub przeliczenia składnika w ramach wizyty.
## GOSPODARKA WODNA (PICIU)
- **Docelowa podaż płynów:** Wyliczona łączna dobowa objętość płynów in ml na masę ciała. Instrukcja szacowania spożycia wody metodą stałej dolewki referencyjnej (np. nalewanie 100 ml i mierzenie ubytku).
- **Zalecana woda:** Niskozmineralizowana (zwłaszcza z niskim wapń i sód) np. Żywiecki kryształ, Primavera źródlana, Mama i ja, przegotowana i odstana.
- **Wody Niezalecane:** Nie używać komercyjnych „wód dla kotów” (nieznana mineralizacja, plastik) oraz wód ze studni głębinowej (za wysoka mineralizacja).
## SUPLEMENTACJA DODATKOWA (CELOWANA)
[Precyzyjne dawkowanie, sugerowane preparaty komercyjne, wpływ na smakowitość i cel wdrożenia dla substancji wymienionych w rozmowie, m.in. Ubichinol, L-karnityna, Kwasy Omega 3, Cordyceps, Astaksantyna]
## WIĄZANIE FOSFORU I GOSPODARKA ŻELAZEM
- **Wiązanie fosforu:** [Zalecenia dotyczące wyłapywaczy fosforu np. sewelamer, wymagane odstępy godzinowe od innych leków, status PorusOne]
- **Gospodarka żelazem:** [Decyzje dotyczące niedokrwistości, diagnostyki laboratoryjnej ferrytyny/TIBC vs stosowanie form iniekcyjnych żelaza]
- **Smaczki funkcjonalne (do 5% kcal / maks 10 kcal dziennie):** Precyzyjne gramatury dobowe dla dopuszczonych bezpiecznych przysmaków (np. łopatka, polędwiczka, indyk).
## AWARYJNE KARMY KOMERCYJNE
W stanach awaryjnych stosować karmy o najniższej zawartości węglowodanów i fosforu w suchej masie (s.m.) (np. Cat's Plate Venison sarna, Cat's Plate Lamb jagnięcina, Cat's Plate Gastro indyk).
## HARMONOGRAM TRANZYCJI (WPROWADZANIE KROK PO KROKU)
- **Tydzień 1:** Woda + Mięso + Podroby + Tłuszcz + Tauryna
- **Tydzień 2:** Składniki z Tygodnia 1 + Wapń/Sól + Dodatkowo: L-karnityna
- **Tydzień 3:** Składniki z Tygodnia 2 + Kwasy Omega 3
- **Tydzień 4:** Składniki z Tygodnia 3 + Witamina E + Dodatkowo: koenzym Q10, olej z kryla, cordyceps
- **Tydzień 5:** Składniki z Tygodnia 4 + Witaminy z grupy B
- **Tydzień 6:** Składniki z Tygodnia 5 + Jod
- **Tydzień 7:** Pełna, kompletna dieta zbilansowana (To będzie już kompletna dieta).
## HARMONOGRAM BADAŃ KONTROLNYCH
- **Parametry nerkowe, wątrobowe, pełny jonogram, żelazo:** Za ok. 1-2 miesiące od wprowadzenia nowej diety lub wcześniej, jeśli stan tego wymaga.
- **Rozmaz manualny krwi, ferrytyna, TIBC, retikulocyty:** Do rozważenia w celu dokładniejszej oceny anemii (morfologia automatyczna bywa niedokładna).
- **T4 całkowita:** Minimum raz na 4 miesiące.
- **USG jamy brzusznej oraz Echo serca / pomiar ciśnienia:** Zgodnie z harmonogramem lekarza weterynarii.
## ZAŁOŻONE ZAŁĄCZNIKI
W pakiecie dokumentów Opiekun otrzymuje: „BARF podstawy“, eBook „BARFNA KUCHNIA“, „Przygotowywanie diety gotowanej w domu”, przepis BACF (2 warianty).
---
Pozdrawiam serdecznie, Anna Michalska | miesnepsokotki@gmail.com | https://www.facebook.com/meatpoint.io"""

# --- SEGMENTATOR PLIKÓW DOCX ---
def segmentuj_docx(file_bytes):
    doc = Document(BytesIO(file_bytes))
    sekcje = {}
    biezaca_sekcja = "Nagłówek i Data wizyty"
    sekcje[biezaca_sekcja] = []
    
    znane_naglowki = [
        "DANE FORMALNE OPIEKUNA", "DANE PACJENTA", "WYWIAD KLINICZNY",
        "WYPRÓŻNIENIA I OBJAWY GASTRYCZNE", "AKTUALNE BADANIA LABORATORYJNE",
        "AKTUALNE LEKI I SUPLEMENTY MEDYCZNE", "KOMENTARZ DO WYWIADU I GŁÓWNE ZAŁOŻENIA DIETY",
        "EDUKACJA OPIEKUNA", "HISTORIA ŻYWIENIOWA I PREFERENCJE SMAKOWE",
        "SPECYFIKACJA NOWEGO PLANU DIETETYCEDNEGO", "GOSPODARKA WODNA (PICIU)",
        "SUPLEMENTACJA DODATKOWA (CELOWANA)", "WIĄZANIE FOSFORU I GOSPODARKA ŻELAZEM",
        "AWARYJNE KARMY KOMERCYJNE", "HARMONOGRAM TRANZYCJI",
        "HARMONOGRAM BADAŃ KONTROLNYCH", "ZAŁOŻONE ZAŁĄCZNIKI"
    ]
    
    for p in doc.paragraphs:
        tekst = p.text.strip()
        if not tekst:
            continue
        znaleziono_naglowek = False
        for naglowek in znane_naglowki:
            if naglowek in tekst.upper() and len(tekst) < 65:
                biezaca_sekcja = naglowek
                if biezaca_sekcja not in sekcje:
                    sekcje[biezaca_sekcja] = []
                znaleziono_naglowek = True
                break
        if not znaleziono_naglowek:
            sekcje[biezaca_sekcja].append(tekst)
            
    for k in sekcje:
        sekcje[k] = "\n".join(sekcje[k])
    return sekcje

def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0563C1')  
    rPr.append(c)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')  
    rPr.append(u)
    new_run.append(rPr)
    text_node = OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def parsuj_i_formatuj_tekst(paragraph, tekst):
    czesci_brak = tekst.split('[BRAK INFORMACJI]')
    for index_brak, czesc_brak in enumerate(czesci_brak):
        if czesc_brak:
            segmenty_url = re.split(r'(https?://[^\s]+)', czesc_brak)
            for idx_seg, seg in enumerate(segmenty_url):
                if idx_seg % 2 == 1:
                    add_hyperlink(paragraph, seg, seg)
                else:
                    if seg:
                        run = paragraph.add_run(seg)
                        run.bold = False
        if index_brak < len(czesci_brak) - 1:
            run_alert = paragraph.add_run('[BRAK INFORMACJI]')
            run_alert.bold = True
            run_alert.font.color.rgb = RGBColor(220, 38, 38)

def konwertuj_do_docx(tekst_markdown):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1.3)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        section.header_distance = Inches(0.4)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10.5)
    style.paragraph_format.line_spacing = 1.25
    style.paragraph_format.space_after = Pt(4)

    section = doc.sections[0]
    section.different_first_page_header_footer = True  

    tabela_naglowka = section.first_page_header.add_table(1, 2, Inches(6.7))
    tabela_naglowka.autofit = False
    tblPr = tabela_naglowka._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

    kol_lewa = tabela_naglowka.rows[0].cells[0]
    kol_prawa = tabela_naglowka.rows[0].cells[1]
    kol_lewa.width = Inches(4.9)
    kol_prawa.width = Inches(1.8)

    p_kontakt = kol_lewa.paragraphs[0]
    p_kontakt.paragraph_format.space_after = Pt(0)
    run_name = p_kontakt.add_run("Anna Michalska\n")
    run_name.bold = True
    run_name.font.size = Pt(11)
    run_sub = p_kontakt.add_run("Dietetyka Psów i Kotów\n")
    run_sub.font.size = Pt(9)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)
    run_det = p_kontakt.add_run("miesnepsokotki@gmail.com  |  https://www.facebook.com/meatpoint.io")
    run_det.font.size = Pt(8.5)
    run_det.font.color.rgb = RGBColor(100, 116, 139)

    if os.path.exists("logo.png"):
        kol_prawa.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        kol_prawa.paragraphs[0].paragraph_format.space_after = Pt(0)
        kol_prawa.paragraphs[0].add_run().add_picture("logo.png", width=Inches(1.0))

    subsequent_header = section.header
    if os.path.exists("logo.png"):
        subsequent_header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        subsequent_header.paragraphs[0].paragraph_format.space_after = Pt(0)
        subsequent_header.paragraphs[0].add_run().add_picture("logo.png", width=Inches(1.0))

    for linia in tekst_markdown.split('\n'):
        linia_strip = linia.strip()
        if not linia_strip:
            continue
        linia_strip = linia_strip.replace('**', '')
            
        if linia_strip.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(linia_strip.replace('## ', ''))
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(194, 65, 12)
        elif linia_strip.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linia_strip.replace('### ', ''))
            run.bold = True
            run.font.size = Pt(10.5)
        elif linia_strip.startswith('- ') or linia_strip.startswith('* '):
            czysty_tekst = linia_strip.lstrip('-* ').strip()
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(3)
            if ':' in czysty_tekst and not czysty_tekst.strip().startswith('http'):
                przed_kolonem, za_kolonem = czysty_tekst.split(':', 1)
                if len(przed_kolonem) < 45:
                    run_bold = p.add_run(przed_kolonem.strip() + ':')
                    run_bold.bold = True
                    parsuj_i_formatuj_tekst(p, za_kolonem)
                    continue
            parsuj_i_formatuj_tekst(p, czysty_tekst)
        else:
            if ':' in linia_strip and not linia_strip.strip().startswith('http'):
                przed_kolonem, za_kolonem = linia_strip.split(':', 1)
                if len(przed_kolonem) < 45:
                    p = doc.add_paragraph()
                    run_bold = p.add_run(przed_kolonem.strip() + ':')
                    run_bold.bold = True
                    parsuj_i_formatuj_tekst(p, za_kolonem)
                    continue
            p = doc.add_paragraph()
            parsuj_i_formatuj_tekst(p, linia_strip)
            
    bufor = BytesIO()
    doc.save(bufor)
    return bufor.getvalue()

with st.sidebar:
    st.header("🔑 Autoryzacja i Model")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-3.5-flash", "gemini-3.1-pro"])

tab1, tab2 = st.tabs(["🚀 Generator Protokołów", "🎙️ Głosowy Edytor (Voice Editor)"])

# ==============================================================================
# 🚀 ZAKŁADKA 1: GENERATOR PROTOKOŁÓW
# ==============================================================================
with tab1:
    st.title("🐾 MeatPoint.io - Asystent Dietetyczny")
    st.write("Wklej surową transkrypcję, aby uzupełnić kliniczny protokół wizyty.")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("🔊 Surowa transkrypcja")
        transcript = st.text_area("Wklej tutaj tekst transkrypcji z Google Meet:", height=550, key="transkrypcja_gen")

    with col2:
        st.subheader("📋 Wynikowy Protokół Wizyty")
        LINK_DO_ARKUSZA = "https://docs.google.com/spreadsheets/d/1qgSX_t4_fb36CqtFUluPDKDQILpR9_SLOlYBPTXSTes/edit?usp=sharing"
        
        if st.button("🚀 Generuj i wypełnij szablon", type="primary", key="btn_gen"):
            if not api_key or not transcript:
                st.error("❌ Uzupełnij klucz API oraz transkrypcję przed uruchomieniem!")
            else:
                with st.spinner("Pobieranie zewnętrznej bazy linków i analiza transkrypcji..."):
                    try:
                        csv_url = LINK_DO_ARKUSZA.replace('/edit?usp=sharing', '/export?format=csv')
                        if '/edit' in LINK_DO_ARKUSZA and '?usp=sharing' not in LINK_DO_ARKUSZA:
                            csv_url = LINK_DO_ARKUSZA.split('/edit')[0] + '/export?format=csv'
                            
                        df_linki = pd.read_csv(csv_url)
                        lista_linkow_prompt = ""
                        for _, row in df_linki.iterrows():
                            lista_linkow_prompt += f"- Link: {row['URL']} | Nazwa: {row['Nazwa']} | Kiedy użyć: {row['Opis dla AI']}\n"
                        
                        genai.configure(api_key=api_key)
                        system_instruction = "Jesteś elitarnym asystentem medycznym dla marki MeatPoint.io. Twoim jedynym zadaniem jest precyzyjne uzupełnianie struktury protokołu wizyty na podstawie transkrypcji. Pisz TYLKO fakty. Brak informacji oznacz jako [BRAK INFORMACJI]. Nie używaj żadnych tagów HTML ani kodów CSS."
                        model = genai.GenerativeModel(model_name=model_choice, system_instruction=system_instruction)
                        
                        prompt = f"Przeanalizuj transkrypcję i uzupełnij dokładnie poniższy szablon kliniczny. Zewnętrzna baza linków:\n{lista_linkow_prompt}\n\nSZABLON:\n{SZABLON_PRODUKCYJNY}\n\nTRANSKRYPCJA ROZMOWY:\n{transcript}"
                        
                        response = model.generate_content(prompt)
                        st.text_area("Podgląd tekstu wygenerowanego przez AI:", value=response.text, height=350, key="podglad_gen")
                        
                        plik_docx = konwertuj_do_docx(response.text)
                        st.markdown("---")
                        st.download_button(
                            label="📥 POBIERZ PROFESJONALNY PLIK WORD (.DOCX)",
                            data=plik_docx,
                            file_name="Protokol_Konsultacji_MeatPoint.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="dl_gen"
                        )
                    except Exception as e:
                        st.error(f"🚨 Błąd generowania dokumentu DOCX: {e}")

# ==============================================================================
# 🎙️ ZAKŁADKA 2: GŁOSOWY EDYTOR PROTOKOŁÓW (Voice Editor)
# ==============================================================================
with tab2:
    st.title("🎙️ Inteligentny Edytor Głosowy Protokółów")
    st.write("Wgraj wcześniej wygenerowany plik Word (.docx). Aplikacja rozbije go na sekcje, a Ty będziesz mógł dokonać poprawek za pomocą głosu.")
    
    if 'sekcje_dokumentu' not in st.session_state:
        st.session_state.sekcje_dokumentu = None
    if 'reset_uploader' not in st.session_state:
        st.session_state.reset_uploader = 0

    col_top1, col_top2 = st.columns([3, 1])
    with col_top1:
        uploaded_file = st.file_uploader(
            "📂 Wgraj plik protokołu MeatPoint (.docx):", 
            type=["docx"], 
            key=f"uploader_voice_{st.session_state.reset_uploader}"
        )
    with col_top2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Rozpocznij nową edycję (Reset)", type="secondary", use_container_width=True):
            st.session_state.sekcje_dokumentu = None
            st.session_state.reset_uploader += 1
            st.rerun()

    if uploaded_file and st.session_state.sekcje_dokumentu is None:
        if st.button("⚙️ Przeanalizuj i załaduj strukturę pliku", type="secondary"):
            st.session_state.sekcje_dokumentu = segmentuj_docx(uploaded_file.read())
            st.success("✅ Plik został pomyślnie rozbity na sekcje kliniczne!")
            st.rerun()

    if st.session_state.sekcje_dokumentu:
        st.markdown("---")
        col_ed1, col_ed2 = st.columns([1, 1])
        
        with col_ed1:
            st.markdown("### 1️⃣ Wybór obszaru do korekty")
            opcje_sekcji = list(st.session_state.sekcje_dokumentu.keys())
            wybrana_sekcja = st.selectbox("Wybierz nagłówek, który chcesz zmodyfikować:", opcje_sekcji, key="wybor_sekcji_voice")
            st.text_area("📄 Aktualna treść tej sekcji w pliku:", value=st.session_state.sekcje_dokumentu[wybrana_sekcja], height=250, disabled=True, key=f"text_obszar_{wybrana_sekcja}")
            
        with col_ed2:
            st.markdown("### 2️⃣ Dyktowanie instrukcji dla AI")
            st.info("Kliknij start, wypowiedz instrukcję i zatrzymaj nagrywanie.")
            
            audio_instrukcja = mic_recorder(
                start_prompt="🎙️ Rozpocznij nagrywanie głosu",
                stop_prompt="🛑 Zatrzymaj i zapisz instrukcję",
                key=f"audio_recorder_{wybrana_sekcja}_{st.session_state.reset_uploader}"
            )
            
            if audio_instrukcja:
                st.audio(audio_instrukcja['bytes'], format="audio/wav")
                st.warning("⚠️ UWAGA: Kliknij poniższy przycisk, aby AI najpierw zaktualizowało tekst sekcji po lewej stronie!")
                
                if st.button("🚀 Wprowadź poprawki głosowe do tekstu", type="primary", key=f"btn_apply_{wybrana_sekcja}"):
                    if not api_key:
                        st.error("❌ Musisz podać klucz API Gemini w panelu bocznym!")
                    else:
                        with st.spinner("Gemini słucha Twojego nagrania i redaguje tekst medyczny..."):
                            try:
                                genai.configure(api_key=api_key)
                                model_edytor = genai.GenerativeModel(model_name=model_choice)
                                
                                prompt_edycji = f"""Zmodyfikuj oryginalny tekst sekcji medycznej na podstawie dołączonych instrukcji głosowych.
                                Oryginalna treść sekcji '{wybrana_sekcja}':
                                \"\"\"
                                {st.session_state.sekcje_dokumentu[wybrana_sekcja]}
                                \"\"\"
                                ZWROT WYŁĄCZNIE zaktualizowany, czysty i profesjonalny tekst tej sekcji. Nie dodawaj żadnych komentarzy od siebie, wyjaśnień ani wstępów typu 'Oto poprawiony tekst:'. Zachowaj strukturę listy i dwukropki."""
                                
                                audio_part = {
                                    "data": audio_instrukcja['bytes'],
                                    "mime_type": "audio/wav"
                                }
                                response_edycja = model_edytor.generate_content([prompt_edycji, audio_part])
                                st.session_state.sekcje_dokumentu[wybrana_sekcja] = response_edycja.text.strip()
                                st.success("🎉 Sekcja zaktualizowana pomyślnie! Zmiany są widoczne po lewej stronie.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"🚨 Błąd edytora głosowego: {e}")
                                
        st.markdown("---")
        st.markdown("### 3️⃣ Zapis gotowego dokumentu")
        st.info("Gdy poprawisz już głosowo wszystkie interesujące Cię sekcje, kliknij poniższy przycisk, aby skompilować finalny plik Word.")
        
        if st.button("📦 Generuj finalny, zaktualizowany plik Word", type="primary", key="btn_build_final"):
            tekst_md_final = ""
            for sekcja, tresc in st.session_state.sekcje_dokumentu.items():
                if sekcja == "Nagłówek i Data wizyty":
                    tekst_md_final += f"{tresc}\n\n"
                else:
                    tekst_md_final += f"## {sekcja}\n{tresc}\n\n"
                    
            plik_docx_final = konwertuj_do_docx(tekst_md_final)
            st.download_button(
                label="📥 POBIERZ POPRAWIONY PROTOKÓŁ (.DOCX)",
                data=plik_docx_final,
                file_name="Protokol_Consultacji_MeatPoint_Poprawiony.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_final_fixed"
            )
