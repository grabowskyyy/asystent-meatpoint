import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document

# Funkcja konwertująca tekst Markdown na sformatowany plik Word (.docx)
def konwertuj_do_docx(tekst_markdown):
    doc = Document()
    
    # Podstawowy parser linii
    for linia in tekst_markdown.split('\n'):
        linia_strip = linia.strip()
        
        if linia_strip.startswith('### '):
            doc.add_heading(linia_strip.replace('### ', ''), level=3)
        elif linia_strip.startswith('## '):
            doc.add_heading(linia_strip.replace('## ', ''), level=2)
        elif linia_strip.startswith('# '):
            doc.add_heading(linia_strip.replace('# ', ''), level=1)
        elif linia_strip.startswith('- ') or linia_strip.startswith('* '):
            czysty_tekst = linia_strip.lstrip('-* ').strip()
            p = doc.add_paragraph(style='List Bullet')
            _parsuj_pogrubienia(p, czysty_tekst)
        elif linia_strip:
            p = doc.add_paragraph()
            _parsuj_pogrubienia(p, linia_strip)
            
    # Zapis do pamięci podręcznej (bufora) zamiast na dysk
    bufor = BytesIO()
    doc.save(bufor)
    return bufor.getvalue()

def _parsuj_pogrubienia(paragraph, tekst):
    # Dzieli tekst według znaczników ** żeby zachować pogrubienia w Wordzie
    czesci = tekst.split('**')
    for i, czesc in enumerate(czesci):
        if i % 2 == 1:
            paragraph.add_run(czesc).bold = True
        else:
            paragraph.add_run(czesc)

# Konfiguracja strony Streamlit
st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

st.title("🐾 MeatPoint.io - Generator Opisów Wizyt (.DOCX)")
st.write("Wklej surową transkrypcję z Google Meet, aby automatycznie wygenerować sformatowany dokument Word.")

# Panel boczny (Sidebar)
with st.sidebar:
    st.header("🔑 Autoryzacja")
    api_key = st.text_input("Klucz API Gemini (skopiuj z Google AI Studio)", type="password")
    model_choice = st.selectbox("Wybierz model AI", ["gemini-2.5-flash", "gemini-1.5-flash"])
    st.markdown("---")
    st.info("Dane są przesyłane bezpośrednio do API Google i nie są wykorzystywane do trenowania modeli.")

system_instruction = """
Jesteś elitarnym, klinicznym asystentem medycznym dla marki MeatPoint.io. Twoim jedynym zadaniem jest precyzyjne przekształcanie surowych transkrypcji w opisy wizyt.
BEZWZGLĘDNE REGUŁY:
1. Działasz WYŁĄCZNIE na bazie faktów z transkrypcji.
2. NIE WOLNO Ci niczego zmyślać (wag, wyników, dawek).
3. Jeśli w transkrypcji brakuje informacji dla jakiegokolwiek punktu w szablonie, MASZ BEZWZGLĘDNY NAKAZ wstawienia w to miejsce kodu: <span style="color: red;">[BRAK INFORMACJI]</span>.
4. Zachowaj wszystkie nagłówki i stałe linki edukacyjne.
"""

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("🔊 Surowa transkrypcja")
    transcript = st.text_area("Wklej tutaj cały tekst wyciągnięty z nagrania Google Meet:", height=450)

with col2:
    st.subheader("📋 Wynikowy Protokół Wizyty")
    
    if st.button("🚀 Generuj i wypełnij szablon", type="primary"):
        if not api_key:
            st.error("❌ Musisz podać klucz API Gemini w panelu bocznym!")
        elif not transcript:
            st.error("❌ Wklej tekst transkrypcji przed uruchomieniem!")
        else:
            with st.spinner("AI analizuje transkrypcję i generuje dokument..."):
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(
                        model_name=model_choice,
                        system_instruction=system_instruction
                    )
                    
                    prompt = f"""
                    Uzupełnij poniższy szablon na podstawie transkrypcji. Pamiętaj o zasadzie czerwonego tagu w przypadku braku danych.
                    
                    ### SZABLON DO WYPEŁNIENIA:
                    ## DANE PACJENTA I OPIEKUNA
                    - **Dietetyk prowadzący:** Anna Michalska
                    - **Data wizyty:** [Wpisz datę lub BRAK INFORMACJI]
                    - **Dane Opiekuna:** [Wpisz imię opiekuna]
                    - **Pacjent:** [Imię zwierzęcia]
                    - **Gatunek/Rasa:** [Kot/Pies / Rasa]
                    - **Wiek:** [Wiek]
                    - **Waga:** [Aktualna waga, tendencje]
                    - **BCS:** [Ocena kondycji w skali 1-9]
                    - **Ilość zwierząt w domu:** [Liczba zwierząt i relacje]
                    - **Sterylizacja/kastracja:** [Tak/Nie + rok]

                    ## WYWIAD KLINICZNY
                    - **Powód konsultacji:** [Historia, zdiagnozowane schorzenia, oczekiwania opiekuna]
                    - **Aktualne samopoczucie:** [Zachowanie, przebyte zabiegi, aktywność]
                    - **Apetyt:** [Stan apetytu, częstotliwość podawania karmy, historia wybredności]
                    - **Pragnienie i nawodnienie:** [Spontaniczne picie, stosowane kroplówki - objętość dobowa]

                    ## WYPRÓŻNIENIA I OBJAWY GASTRYCZNE
                    - **Kał:** [Częstotliwość, uformowanie, stan jelit z USG]
                    - **Wymioty:** [Częstotliwość, po czym występują]
                    - **Mocz:** [Ciężar właściwy, pH, proteinuria, erytrocyty/infekcja, częstotliwość mikcji]
                    - **Odrobaczanie:** [Ostatnia data, preparat]

                    ## AKTUALNE BADANIA LABORATORYJNE
                    - **Kreatynina:** [Wartość + trend]
                    - **Mocznik:** [Wartość + trend]
                    - **Fosfor:** [Wartość + trend]
                    - **T4 całkowita:** [Wartość + trend]
                    - **Morfologia (HGB / Anemia):** [Wartość, diagnoza]
                    - **Albuminy:** [Wartość + trend]
                    - **&alpha;-amylaza:** [Wartość + trend]
                    - **Cholesterol:** [Wartość + trend]
                    - **WBC (Leukocyty):** [Wartość, stan zapalny]
                    - **Gospodarka cukrowa:** [Fruktozamina/Glukoza]

                    ## AKTUALNE LEKI I SUPLEMENTY MEDYCZNE
                    [Wypisz listę leków z dawkowaniem podaną w rozmowie]

                    ## KOMENTARZ DO WYWIADU I GŁÓWNE ZAŁOŻENIA DIETY
                    - **Komentarz:** [Podsumowanie stopnia trudności pacjenta]
                    - **Główne założenia diety:** [Poziom fosforu, jakość białka, tłuszcze, węglowodany]

                    ## EDUKACJA OPIEKUNA: CO SIĘ ZMIENI NA DIECIE BARF/BACF
                    - **Częstotliwość kału:** Kot może oddawać mniejszy kał i rzadziej. Filmy edukacyjne:
                      * https://www.facebook.com/reel/1860436634490613
                      * https://www.facebook.com/reel/1701233670818761
                    - **UWAGA NA ZAPARCIA:** Opis struktury stolca (sucha, twarda, bobki).
                    - **Parametry krwi:** Mocznik i kreatynina mogą się różnić od norm referencyjnych dla kotów komercyjnych. Kontrola u nefrologa.

                    ## HISTORIA ŻYWIENIOWA I PREFERENCJE SMAKOWE
                    - **Dotychczasowe żywienie:** [Opis diet, marki, akceptacja]
                    - **KATEGORYCZNIE TAK (Ulubione smaki):** [Zaakceptowane białka, warzywa. WAŻNE: preferencje temperaturowe/mrożenie]
                    - **KATEGORYCZNIE NIE (Odrzucone składniki):** [Czego zwierzę nie zje, co wywołuje wymioty]

                    ## SPECYFIKACJA NOWEGO PLANU DIETETYCZNEGO
                    - **Model diety:** [Model, np. BACF świeży/lodówka]
                    - **Białka bazowe i dodatki:** [Wybrane gatunki mięs, podrobów, warzyw]
                    - **Kaloryczność próbna:** ok. [Wartość] kcal/dzień.

                    ## GOSPODARKA WODNA (PICIU)
                    - **Docelowa podaż płynów:** Łącznie ok. [Wartość] ml wody na dobę.
                    - **Zalecana woda:** Niskozmineralizowana (Żywiecki Kryształ, Primavera, Mama i ja).

                    ## SUPLEMENTACJA DODATKOWA (CELOWANA)
                    [Wypisz dawkowanie z rozmowy dla: Ubichinol, L-karnityna, Omega 3, Cordyceps, Astaksantyna]

                    ## WIĄZANIE FOSFORU I GOSPODARKA ŻELAZEM
                    - **Wiązanie fosforu:** [Zalecenia, sewelamer, status PorusOne]
                    - **Smaczki funkcjonalne (do 5% kcal / maks 10 kcal dziennie):** Dopuszczalne gramatury dla łopatki, polędwiczki, indyka. Kalkulator: https://meatpoint.io/pl/barf-wiedza/smaczki-i-dodatkowe-kalorie-obliczanie-kalorycznosci-komercyjnych-produktow

                    ## AWARYJNE KARMY KOMERCYJNE
                    W stanach awaryjnych stosować karmy o niskiej zawartości węglowodanów i fosforu (np. Cat's Plate Venison, Lamb, Gastro).
                    Edukacja o tyndalizacji: https://meatpoint.io/pl/barf-wiedza/tyndalizacja-czyli-jak-przechowywac-posilki-jesli-nie-chcemy-ich-mrozic

                    ## HARMONOGRAM TRANZYCJI (WPROWADZANIE KROK PO KROKU)
                    - Tydzień 1: Woda + Mięso + Podroby + Tłuszcz + Tauryna
                    - Tydzień 2: Baza + Wapń/Sól + L-karnityna
                    - Tydzień 3: Skład z Tygodnia 2 + Kwasy Omega 3
                    - Tydzień 4: Skład z Tygodnia 3 + Witamina E + koenzym Q10, olej z kryla, cordyceps
                    - Tydzień 5: Skład z Tygodnia 4 + Witaminy z grupy B
                    - Tydzień 6: Skład z Tygodnia 5 + Jod

                    ## HARMONOGRAM BADAŃ KONTROLNECH
                    [Lista zalecanych badań krew/mocz/USG i terminy]

                    ---
                    Pozdrawiam serdecznie,
                    Anna Michalska
                    https://www.facebook.com/meatpoint.io
                    
                    ---
                    TU WKLEJ TRANSKRYPCJĘ ROZMOWY:
                    {transcript}
                    """
                    
                    response = model.generate_content(prompt)
                    
                    # Pokazujemy podgląd na ekranie
                    st.markdown(response.text, unsafe_allow_html=True)
                    
                    # Budowanie pliku Word w pamięci
                    plik_docx = konwertuj_do_docx(response.text)
                    
                    st.markdown("---")
                    # Przycisk pobierania pliku Word
                    st.download_button(
                        label="📥 Pobierz opis jako plik Word (.docx)",
                        data=plik_docx,
                        file_name="Opis_Wizyty_Uzupelniony.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"🚨 Wystąpił błąd: {e}")