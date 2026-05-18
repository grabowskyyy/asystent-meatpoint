import streamlit as st
import google.generativeai as genai
import os
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Funkcja dodająca prawdziwe, klikalne linki (hiperłącza) do pliku Word (.docx)
def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0563C1')  # Elegancki niebieski kolor linku
    rPr.append(c)
    
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')  # Podkreślenie dolne
    rPr.append(u)
    
    new_run.append(rPr)
    
    text_node = OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)
    
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

# Zaawansowana funkcja formatująca tekst (obsługa linków i czerwonych alertów)
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

# Główna funkcja generująca plik .docx ze zróżnicowanymi nagłówkami
def konwertuj_do_docx(tekst_markdown):
    doc = Document()
    
    # Ustawienia marginesów dokumentu
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

    # --- LOGIKA ZRÓŻNICOWANIA NAGŁÓWKÓW ---
    section = doc.sections[0]
    section.different_first_page_header_footer = True  # Podział: Strona 1 ma inny nagłówek niż reszta

    # 1. NAGŁÓWEK DLA PIERWSZEJ STRONY (Pełne dane firmy + Logo)
    first_page_header = section.first_page_header
    tabela_naglowka = first_page_header.add_table(1, 2, Inches(6.7))
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
        p_logo = kol_prawa.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_logo.paragraph_format.space_after = Pt(0)
        p_logo.add_run().add_picture("logo.png", width=Inches(1.0))

    # 2. NAGŁÓWEK DLA KOLEJNYCH STRON (Wyłącznie samo czyste Logo w rogu)
    subsequent_header = section.header
    if os.path.exists("logo.png"):
        p_sub_logo = subsequent_header.paragraphs[0]
        p_sub_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_sub_logo.paragraph_format.space_after = Pt(0)
        p_sub_logo.add_run().add_picture("logo.png", width=Inches(1.0))

    # --- PARSER LINII TEKSTU ---
    for linia in tekst_markdown.split('\n'):
        linia_strip = linia.strip()
        
        if not linia_strip:
            continue
            
        linia_strip = linia_strip.replace('**', '')
            
        # Nagłówki główne (##)
        if linia_strip.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(linia_strip.replace('## ', ''))
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(194, 65, 12)  # Pomarańczowy MeatPoint
            
        # Podnagłówki (###)
        elif linia_strip.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linia_strip.replace('### ', ''))
            run.bold = True
            run.font.size = Pt(10.5)
            
        # Punkty z listy (- / *) -> Podział do dwukropka
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
            
        # Zwykłe linie tekstu
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

# Interfejs Streamlit
st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

st.title("🐾 MeatPoint.io - Generator Protokołów Konsultacji")
st.write("Wklej surową transkrypcję, aby wygenerować profesjonalny dokument Word z automatyczną korektą nagłówków stron.")

with st.sidebar:
    st.header("🔑 Autoryzacja")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-2.5-flash", "gemini-1.5-flash"])

system_instruction = """
Jesteś elitarnym asystentem medycznym dla marki MeatPoint.io. Twoim jedynym zadaniem jest precyzyjne uzupełnianie struktury protokołu wizyty na podstawie transkrypcji.
REGUŁY BEZWZGLĘDNE:
1. Pisz TYLKO fakty podane bezpośrednio w transkrypcji rozmowy.
2. Jeśli w tekście brakuje informacji do jakiejkolwiek rubryki lub punktu (np. PESEL, adres, wyniki badań), wstaw tekst: [BRAK INFORMACJI]
3. NIE wolno Ci niczego zmyślać ani pominąć żadnego nagłówka. Jeśli brak danych - zostaw nagłówek i napisz [BRAK INFORMACJI].
4. NIE używaj żadnych tagów HTML ani kodów CSS. Wstawiaj wyłącznie czysty tekst.
"""

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("🔊 Surowa transkrypcja")
    transcript = st.text_area("Wklej tutaj tekst transkrypcji z Google Meet:", height=550)

with col2:
    st.subheader("📋 Wynikowy Protokół Wizyty")
    
    if st.button("🚀 Generuj i wypełnij szablon", type="primary"):
        if not api_key or not transcript:
            st.error("❌ Uzupełnij klucz API oraz transkrypcję przed uruchomieniem!")
        else:
            with st.spinner("AI analizuje transkrypcję i formatuje strukturę stron..."):
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(model_name=model_choice, system_instruction=system_instruction)
                    
                    prompt = f"""
                    Przeanalizuj poniższą transkrypcję i uzupełnij dokładnie ten szablon. Zachowaj wszystkie teksty edukacyjne i linki bez zmian. Jeśli brakuje danych, wstaw [BRAK INFORMACJI].
                    
                    ### SZABLON DO WYPEŁNIENIA:
                    Data wizyty: [Wpisz datę wizyty lub BRAK INFORMACJI]

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
                    [Zestawienie najnowszych wyników wraz z datami i interpretacją trendu klinicznego]:
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
                    - **Częstotliwość kału:** Zwierzę może oddawać mniejszy kał i może go oddawać co 2–3 dni. Na wysokomięsnej diecie to normalne. Ważne, żeby był dobrego kształtu i konsystencji (wdł skali bristolskiej). Dokładne informacje w tych filmach: https://www.facebook.com/reel/1860436634490613 oraz https://www.facebook.com/reel/1701233670818761
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
                    - **Smaczki funkcjonalne (do 5% kcal / maks 10 kcal dziennie):** Precyzyjne gramatury dobowe dla dopuszczonych bezpiecznych przysmaków (np. łopatka, polędwiczka, indyk). Link do kalkulatora: https://meatpoint.io/pl/barf-wiedza/smaczki-i-dodatkowe-kalorie-obliczanie-kalorycznosci-komercyjnych-produktow

                    ## AWARYJNE KARMY KOMERCYJNE
                    W stanach awaryjnych stosować karmy o najniższej zawartości węglowodanów i fosforu w suchej masie (s.m.) (np. Cat's Plate Venison sarna, Cat's Plate Lamb jagnięcina, Cat's Plate Gastro indyk).
                    Edukacja o tyndalizacji posiłków jako metodzie przechowywania: https://meatpoint.io/pl/barf-wiedza/tyndalizacja-czyli-jak-przechowywac-posilki-jesli-nie-chcemy-ich-mrozic oraz film instruktażowy: https://youtu.be/tyfT3kmq3ME

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
                    Pozdrawiam serdecznie,
                    Anna Michalska
                    miesnepsokotki@gmail.com
                    https://www.facebook.com/meatpoint.io
                    
                    ---
                    TU WKLEJ TRANSKRYPCJĘ ROZMOWY:
                    {transcript}
                    """
                    
                    response = model.generate_content(prompt)
                    
                    st.text_area("Podgląd tekstu wygenerowanego przez AI:", value=response.text, height=350)
                    
                    plik_docx = konwertuj_do_docx(response.text)
                    
                    st.markdown("---")
                    st.download_button(
                        label="📥 POBIERZ PROFESJONALNY PLIK WORD (.DOCX)",
                        data=plik_docx,
                        file_name="Protokol_Konsultacji_MeatPoint.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"🚨 Błąd generowania dokumentu DOCX: {e}")
