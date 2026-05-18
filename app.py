import streamlit as st
import google.generativeai as genai
import os
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Funkcja do natywnego formatowania tekstu w Wordzie (obsługa pogrubień i czerwonego alertu)
def parsuj_i_formatuj_tekst(paragraph, tekst):
    czesci_bold = tekst.split('**')
    for index_bold, czesc_bold in enumerate(czesci_bold):
        is_bold = (index_bold % 2 == 1)
        
        # Wyszukiwanie znacznika [BRAK INFORMACJI]
        czesci_brak = czesc_bold.split('[BRAK INFORMACJI]')
        for index_brak, czesc_brak in enumerate(czesci_brak):
            if czesc_brak:
                run = paragraph.add_run(czesc_brak)
                if is_bold:
                    run.bold = True
            
            # Kolorowanie alertu o braku danych na czerwono
            if index_brak < len(czesci_brak) - 1:
                run_alert = paragraph.add_run('[BRAK INFORMACJI]')
                run_alert.bold = True
                run_alert.font.color.rgb = RGBColor(220, 38, 38) # Czerwony

# Główna funkcja generująca zaawansowany plik .docx ze stylami MeatPoint
def konwertuj_do_docx(tekst_markdown):
    doc = Document()
    
    # Ustawienia marginesów dokumentu
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Ustawienie globalnej czcionki Arial i interlinii 1.25
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10.5)
    style.paragraph_format.line_spacing = 1.25
    style.paragraph_format.space_after = Pt(4)

    # --- NAGŁÓWEK BRANDINGOWY (Tabela dwukolumnowa) ---
    tabela_naglowka = doc.add_table(rows=1, cols=2)
    tabela_naglowka.autofit = False
    
    # Usunięcie obramowania tabeli nagłówkowej
    tblPr = tabela_naglowka._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

    kol_lewa = tabela_naglowka.rows[0].cells[0]
    kol_prawa = tabela_naglowka.rows[0].cells[1]
    kol_lewa.width = Inches(4.7)
    kol_prawa.width = Inches(1.8)

    # Dane po lewej stronie
    p_kontakt = kol_lewa.paragraphs[0]
    p_kontakt.paragraph_format.space_after = Pt(2)
    run_name = p_kontakt.add_run("Anna Michalska\n")
    run_name.bold = True
    run_name.font.size = Pt(12)
    
    run_sub = p_kontakt.add_run("Dietetyka Psów i Kotów\n")
    run_sub.font.size = Pt(10)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)
    
    run_det = p_kontakt.add_run("miesnepsokotki@gmail.com\nhttps://www.facebook.com/meatpoint.io")
    run_det.font.size = Pt(9.5)

    # NAPRAWIONY BŁĄD LOGO: Bezpośrednie wstrzyknięcie obrazka do akapitu kolumny
    if os.path.exists("logo.png"):
        p_logo = kol_prawa.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_logo.add_run().add_picture("logo.png", width=Inches(1.2))
    
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- PARSER TEKSTU ---
    for linia in tekst_markdown.split('\n'):
        linia_strip = linia.strip()
        
        if not linia_strip:
            continue
            
        # Nagłówki Główne Sekcji (##) -> Kolor Ciemnopomarańczowy
        if linia_strip.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(linia_strip.replace('## ', ''))
            run.bold = True
            run.font.size = Pt(12.5)
            run.font.color.rgb = RGBColor(194, 65, 12) # Kolor MeatPoint
            
        # Podnagłówki (###)
        elif linia_strip.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linia_strip.replace('### ', ''))
            run.bold = True
            run.font.size = Pt(11)
            
        # Listy punktowane
        elif linia_strip.startswith('- ') or linia_strip.startswith('* '):
            czysty_tekst = linia_strip.lstrip('-* ').strip()
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(3)
            parsuj_i_formatuj_tekst(p, czysty_tekst)
            
        # Zwykły tekst
        else:
            p = doc.add_paragraph()
            parsuj_i_formatuj_tekst(p, linia_strip)
            
    bufor = BytesIO()
    doc.save(bufor)
    return bufor.getvalue()

# Interfejs Streamlit
st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

st.title("🐾 MeatPoint.io - Generator Protokołów Konsultacji")
st.write("Wklej surową transkrypcję, aby wygenerować profesjonalny dokument Word zgodny z identyfikacją marki.")

with st.sidebar:
    st.header("🔑 Autoryzacja")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-2.5-flash", "gemini-1.5-flash"])

system_instruction = """
Jesteś elitarnym asystentem medycznym MeatPoint.io. Twoim jedynym zadaniem jest precyzyjne uzupełnianie struktury protokołu wizyty.
REGUŁY:
1. Pisz TYLKO fakty podane bezpośrednio w transkrypcji.
2. Jeśli brakuje informacji dla jakiegoś punktu, wstaw tekst: [BRAK INFORMACJI]
3. NIE używaj żadnych tagów HTML ani kodów kolorów. Wstawiaj czysty tekst tekstowy.
"""

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("🔊 Surowa transkrypcja")
    transcript = st.text_area("Wklej tutaj tekst z Google Meet:", height=500)

with col2:
    st.subheader("📋 Wynikowy Protokół Wizyty")
    
    if st.button("🚀 Generuj i wypełnij szablon", type="primary"):
        if not api_key or not transcript:
            st.error("❌ Uzupełnij klucz API oraz transkrypcję!")
        else:
            with st.spinner("Przetwarzanie dokumentu..."):
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(model_name=model_choice, system_instruction=system_instruction)
                    
                    # Wstrzyknięcie pełnego, rygorystycznego szablonu firmy MeatPoint
                    prompt = f"""
                    Uzupełnij dokładnie poniższy szablon na podstawie załączonej transkrypcji rozmowy. Jeśli jakiegoś parametru brakuje, wpisz [BRAK INFORMACJI].
                    
                    ### SZABLON DO WYPEŁNIENIA:
                    Data wizyty: [Data wizyty]

                    ## DANE PACJENTA I OPIEKUNA
                    - **Dane Opiekuna:** [Imię i nazwisko opiekuna]
                    - **Pacjent:** [Imię zwierzęcia]
                    - **Gatunek:** [Kot/Pies]
                    - **Rasa:** [Rasa]
                    - **Wiek:** [Wiek zwierzęcia]
                    - **Waga:** [Aktualna waga i cele wagowe]
                    - **BCS:** [Ocena kondycji w skali 1-9/9]
                    - **Ilość zwierząt w domu:** [Liczba zwierząt w stadzie i relacje]
                    - **Sterylizacja/kastracja:** [Tak/Nie + rok]

                    ## WYWIAD KLINICZNY
                    - **Powód konsultacji:** [Historia schorzenia, zdiagnozowane jednostki chorobowe, główne oczekiwania opiekuna wobec diety]
                    - **Aktualne samopoczucie:** [Zachowanie, przebyte niedawno zabiegi, np. sanacja jamy ustnej, stopień aktywności]
                    - **Apetyt:** [Stan apetytu, częstotliwość karmienia, wybredność, przyjmowane dotychczas karmy]
                    - **Pragnienie i nawodnienie:** [Ilość samodzielnego picia, stosowane kroplówki - objętość i rodzaj płynów]

                    ## WYPRÓŻNIENIA I OBJAWY GASTRYCZNE
                    - **Kał:** [Częstotliwość na dobę, uformowanie, opis jelit z USG pod kątem IBD]
                    - **Wymioty:** [Częstotliwość, po czym występują]
                    - **Mocz:** [Barwa, ciężar właściwy, proteinuria/białko, obecność erytrocytów, infekcje dróg moczowych, dobowy schemat mikcji]
                    - **Odrobaczanie:** [Ostatnia data, zastosowany preparat]

                    ## AKTUALNE BADANIA LABORATORYJNE
                    [Zestawienie najnowszych wyników wraz z interpretacją trendu klinicznego]:
                    - **Kreatynina:** [Wartość]
                    - **Mocznik:** [Wartość]
                    - **Fosfor:** [Wartość]
                    - **T4 całkowita:** [Wartość]
                    - **Morfologia (HGB / Anemia):** [Wartość]
                    - **Albuminy:** [Wartość]
                    - **&alpha;-amylaza:** [Wartość]
                    - **Cholesterol:** [Wartość]
                    - **WBC (Leukocyty):** [Wartość]
                    - **Gospodarka cukrowa (Fruktozamina):** [Wartość]

                    ## AKTUALNE LEKI I SUPLEMENTY MEDYCZNE
                    [Wypisz listę stosowanych leków i suplementów wraz z dawkowaniem podanym przez opiekuna]

                    ## KOMENTARZ DO WYWIADU I GŁÓWNE ZAŁOŻENIA DIETY
                    - **Komentarz:** [Podsumowanie stopnia trudności pacjenta, tolerancji składników i wymaganych kompromisów]
                    - **Główne założenia diety:** [Cele makroskładnikowe: poziom fosforu, jakość białka, tłuszcze, węglowodany, wsparcie celowane narządów]

                    ## EDUKACJA OPIEKUNA: CO SIĘ ZMIENI NA DIECIE BARF/BACF
                    - **Częstotliwość kału:** Zwierzę może oddawać mniejszy kał i rzadziej (co 2-3 dni). Ważny jest kształt i konsystencja zgodna ze skalą bristolską. Filmy edukacyjne: https://www.facebook.com/reel/1860436634490613 oraz https://www.facebook.com/reel/1701233670818761
                    - **UWAGA NA ZAPARCIA:** Należy monitorować wygląd stolca (sucha, twarda kupa, bobki).
                    - **Parametry krwi:** Mocznik i kreatynina na diecie wysokomięsnej mogą się różnić od norm referencyjnych dla zwierząt komercyjnych. Wymagana stała kontrola nefrologiczna.
                    - **Objętość posiłku:** Dieta domowa jest bardziej kaloryczna w mniejszej objętości. Przyzwyczajanie zwierzęcia może zająć 2-3 miesiące.

                    ## HISTORIA ŻYWIENIOWA I PREFERENCJE SMAKOWE
                    - **Dotychczasowe żywienie:** [Opis stosowanych wcześniej diet, przepisy, powody rezygnacji]
                    - **KATEGORYCZNIE TAK (Ulubione smaki):** [Zaakceptowane białka, warzywa, preferowana forma podania i temperatura posiłków. UWAGA: Podkreśl, czy je mrożone czy tylko świeże]
                    - **KATEGORYCZNIE NIE (Odrzucone składniki):** [Składniki, mięsa, formy wapnia lub suplementy wywołujące wymioty lub całkowity bunt pacjenta]

                    ## SPECYFIKACJA NOWEGO PLANU DIETETYCZNEGO
                    - **Model diety:** [Model żywienia, logistyka przygotowania, rotacja przepisów]
                    - **Białka bazowe i dodatki:** [Wybrane gatunki mięs, podrobów i dozwolonych warzyw]
                    - **Kaloryczność próbna:** [Wartość] kcal/dzień. Warunki jednorazowej korekty składników lub kaloryczności.

                    ## GOSPODARKA WODNA (PICIU)
                    - **Docelowa podaż płynów:** Wyliczona dobowa objętość płynów w ml na masę ciała.
                    - **Zalecana woda:** Niskozmineralizowana (Żywiecki Kryształ, Primavera źródlana, Mama i ja), przegotowana i odstana. Unikać wód komercyjnych dla kotów oraz wód ze studni głębinowych.

                    ## SUPLEMENTACJA DODATKOWA (CELOWANA)
                    [Precyzyjne dawkowanie i schemat wprowadzania dla substancji wymienionych w rozmowie, m.in. Ubichinol, L-karnityna, Omega 3, Cordyceps, Astaksantyna]

                    ## WIĄZANIE FOSFORU I GOSPODARKA ŻELAZEM
                    - **Wiązanie fosforu:** [Rekomendacje dotyczące wyłapywaczy fosforu, np. sewelamer, odstępy godzinowe od leków]
                    - **Gospodarka żelazem:** [Decyzje dotyczące anemii, diagnostyki laboratoryjnej lub form iniekcyjnych żelaza]
                    - **Smaczki funkcjonalne (do 5% kcal / maks 10 kcal dziennie):** Precyzyjne gramatury dobowe dla dopuszczonych przysmaków (np. łopatka, polędwiczka, indyk). Link do kalkulatora: https://meatpoint.io/pl/barf-wiedza/smaczki-i-dodatkowe-kalorie-obliczanie-kalorycznosci-komercyjnych-produktow

                    ## AWARYJNE KARMY KOMERCYJNE
                    W stanach awaryjnych stosować karmy o niskiej zawartości węglowodanów i fosforu w suchej masie (np. Cat's Plate Venison, Lamb, Gastro).
                    Edukacja o tyndalizacji posiłków: https://meatpoint.io/pl/barf-wiedza/tyndalizacja-czyli-jak-przechowywac-posilki-jesli-nie-chcemy-ich-mrozic oraz film: https://youtu.be/tyfT3kmq3ME

                    ## HARMONOGRAM TRANZYCJI (WPROWADZANIE KROK PO KROKU)
                    - **Tydzień 1:** Woda + Mięso + Podroby + Tłuszcz + Tauryna
                    - **Tydzień 2:** Składniki z Tygodnia 1 + Wapń/Sól + Dodatkowo: L-karnityna
                    - **Tydzień 3:** Składniki z Tygodnia 2 + Kwasy Omega 3
                    - **Tydzień 4:** Składniki z Tygodnia 3 + Witamina E + Dodatkowo: koenzym Q10, olej z kryla, cordyceps
                    - **Tydzień 5:** Składniki z Tygodnia 4 + Witaminy z grupy B
                    - **Tydzień 6:** Składniki z Tygodnia 5 + Jod
                    - **Tydzień 7:** Pełna, kompletna dieta zbilansowana

                    ## HARMONOGRAM BADAŃ KONTROLNYCH
                    [Lista zalecanych badań kontrolnych krwi, moczu, USG, echo serca wraz z terminami]

                    ## ZAŁOŻONE ZAŁĄCZNIKI
                    W pakiecie dokumentów Opiekun otrzymuje: „BARF podstawy“, eBook „BARFNA KUCHNIA“, „Przygotowywanie diety gotowanej w domu”, Przepis BACF (2 warianty).

                    ---
                    TU WKLEJ TRANSKRYPCJĘ ROZMOWY:
                    {transcript}
                    """
                    
                    response = model.generate_content(prompt)
                    
                    # Czysty, czytelny podgląd tekstu w aplikacji
                    st.text_area("Podgląd wygenerowanego tekstu:", value=response.text, height=400)
                    
                    # Generowanie i formatowanie pliku DOCX
                    plik_docx = konwertuj_do_docx(response.text)
                    
                    st.markdown("---")
                    st.download_button(
                        label="📥 POBIERZ PROFESJONALNY PLIK WORD (.DOCX)",
                        data=plik_docx,
                        file_name="Protokol_Konsultacji_MeatPoint.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"🚨 Błąd generowania pliku DOCX: {e}")
