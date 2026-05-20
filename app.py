import streamlit as st, google.generativeai as genai, os, re, pandas as pd, uuid, time
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from streamlit_mic_recorder import mic_recorder

# --- KOMPLETNA STRUKTURA I UNIFIKACJA NAZEWNICTWA 1:1 Z DOKUMENTEM ANI ---
STRUKTURA_PROTOKOLU = [
    "Powód konsultacji:", "Aktualne samopoczucie:", "Aktywność:", "Apetyt:", "Pragnienie:",
    "Dotychczasowe żywienie:", "Smaczki i przysmaki:", "Ulubione smaki:",
    "### Kategorycznie tak:", "### Kategorycznie nie:", "### Kluczowa uwaga dot. przechowywania:",
    "Kał / Biegunka / Wymioty:", "Mocz:", "Odrobaczanie:", "Aktualne badania:", "Aktualne leki:",
    "Komentarz do wywiadu:", "Główne założenia diety:", "Co się zmieni na diecie BARF/BACF:",
    "Plan dietetyczny:", "Tranzycja i przechowywanie:", "Kaloryczność:", "Piciu:",
    "### Jakiej wody używać?", "Suplementy dodatkowe:", "Wiązanie fosforu:", "Smaczki:",
    "Inne smaczki:", "Karmy komercyjne:", "Tyndalizacja:", "Wprowadzanie suplementów:",
    "Badania kontrolne:", "Załączniki:"
]

TEKST_TYNDALIZACJA_STALY = (
    "Jeżeli robią Państwo dietę na dłużej niż 5-6 dni (mowa o diecie surowej) i chcą Państwo "
    "ją bezpiecznie przechowywać w słoiczkach w lodówce (bez zamrażania) LUB przygotowują Państwo "
    "dietę gotowaną (BACF) na zapas, konieczne jest przeprowadzenie procesu tyndalizacji (potrójnej pasteryzacji).\n\n"
    "Proces ten skutecznie eliminuje formy przetrwalnikowe bakterii (m.in. Clostridium botulinum - jadu kiełbasianego), "
    "które mogłyby namnażać się w warunkach beztlenowych zamkniętego słoika.\n\n"
    "Pełną instrukcję krok po kroku, jak prawidłowo i bezpiecznie przeprowadzić ten proces w domowych warunkach, "
    "znajdą Państwo w naszym artykule na blogu: https://meatpoint.io/pl/barf-wiedza/tyndalizacja-czyli-jak-przechowywac-posilki-jesli-nie-chcemy-ich-mrozic\n\n"
    "Dodatkowo przygotowaliśmy dla Państwa praktyczny poradnik w formie wideo na platformie YouTube, "
    "gdzie pokazujemy cały proces krok po kroku: https://www.youtube.com/watch?v=tyfT3kmq3ME"
)

TEKST_INNE_SMACZKI_STALY = (
    "Wprowadzając do codziennej rutyny jakiekolwiek inne smaczki komercyjne, należy bezwzględnie "
    "pamiętać o kontrolowaniu ich kaloryczności, aby nie zaburzyć bilansu nowej diety pacjenta.\n\n"
    "Szczegółowy poradnik oraz instrukcję, jak samodzielnie wyliczyć kaloryczność dowolnego produktu komercyjnego "
    "na podstawie danych z etykiety, znajdawą Państwo w naszym artykule: "
    "https://meatpoint.io/pl/barf-wiedza/smaczki-i-dodatkowe-kalorie-obliczanie-kalorycznosci-komercyjnych-produktow"
)

TEKST_WPROWADZANIE_SUPLEMENTOW_STALY = (
    "Proszę zacząć od:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• tłuszczu\n• Tauryny\n"
    "Proszę przygotować dietę tylko z ich zawartością i na razie pominąć pozostałe suplementy.\n\n"
    "Jak Kicia będzie się dobrze czuła, na następny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n"
    "• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n"
    "• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n"
    "• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n"
    "• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n• Jodu\n"
    "• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n• Jodu\n"
    "• Dodatkowo: \n\n"
    "To będzie już kompletna dieta."
)

def segmentuj_docx(file_bytes):
    doc = Document(BytesIO(file_bytes))
    sekcje = {}; biezaca = "Nagłówek i Metryczka"; sekcje[biezaca] = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t: continue
        z = False
        for n in STRUKTURA_PROTOKOLU:
            czysty_n = n.replace("### ", "").strip().upper()
            if czysty_n in t.upper() and len(t) < 65: biezaca = n; sekcje[biezaca] = []; z = True; break
        if not z: sekcje[biezaca].append(t)
    return {k: "\n".join(v) for k, v in sekcje.items()}

def add_hyperlink(p, url, text):
    part = p.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hl = OxmlElement('w:hyperlink'); hl.set(qn('r:id'), r_id)
    nr = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    c = OxmlElement('w:color'); c.set(qn('w:val'), '0563C1'); rPr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    nr.append(rPr); tn = OxmlElement('w:t'); tn.text = text; nr.append(tn); hl.append(nr); p._p.append(hl)
    return hl

def parsuj_i_formatuj_tekst(p, tekst):
    parts = tekst.split('[BRAK INFORMACJI]')
    for i, part in enumerate(parts):
        if part:
            segs = re.split(r'(https?://[^\s]+)', part)
            for idx, seg in enumerate(segs):
                if idx % 2 == 1: add_hyperlink(p, seg, seg)
                else: p.add_run(seg).bold = False
        if i < len(parts) - 1:
            ra = p.add_run('[BRAK INFORMACJI]'); ra.bold = True; ra.font.color.rgb = RGBColor(220, 38, 38)

def konwertuj_do_docx(tekst_md):
    doc = Document()
    for s in doc.sections: s.top_margin, s.bottom_margin, s.left_margin, s.right_margin, s.header_distance = Inches(1.3), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.4)
    style = doc.styles['Normal']; font = style.font; font.name, font.size, style.paragraph_format.line_spacing, style.paragraph_format.space_after = 'Arial', Pt(10.5), 1.25, Pt(4)
    sec = doc.sections[0]; sec.different_first_page_header_footer = True
    
    t_h = sec.first_page_header.add_table(1, 2, Inches(6.7)); t_h.autofit = False
    t_h._tbl.tblPr.append(OxmlElement('w:tblBorders'))
    kl, kp = t_h.rows[0].cells[0], t_h.rows[0].cells[1]; kl.width, kp.width = Inches(4.9), Inches(1.8)
    
    pk = kl.paragraphs[0]; pk.paragraph_format.space_after = Pt(0)
    pk.add_run("Anna Michalska\n").bold = True; pk.runs[-1].font.size = Pt(11)
    ps = pk.add_run("Dietetyka Psów i Kotów\n"); ps.font.size, ps.font.color.rgb = Pt(9), RGBColor(100, 116, 139)
    pd_t = pk.add_run("miesnepsokotki@gmail.com  |  https://www.facebook.com/meatpoint.io"); pd_t.font.size, pd_t.font.color.rgb = Pt(8.5), RGBColor(100, 116, 139)
    
    if os.path.exists("logo.png"):
        kp.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        kp.paragraphs[0].paragraph_format.space_after = Pt(0)
        kp.paragraphs[0].add_run().add_picture("logo.png", width=Inches(1.0))
        sec.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sec.header.paragraphs[0].paragraph_format.space_after = Pt(0)
        sec.header.paragraphs[0].add_run().add_picture("logo.png", width=Inches(1.0))

    for line in tekst_md.split('\n'):
        l_s = line.strip()
        if not l_s: continue
        l_s = l_s.replace('**', '')
        
        if "DATA WIZYTY:" in l_s.upper():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(12), Pt(6)
            poczatek_daty, *koniec_daty = l_s.split(':', 1)
            p.add_run(poczatek_daty + ': ').bold = True
            if koniec_daty: parsuj_i_formatuj_tekst(p, koniec_daty[0].strip())
            continue

        if l_s.startswith('## '):
            p = doc.add_paragraph(); p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(14), Pt(4)
            r = p.add_run(l_s.replace('## ', '')); r.bold = True; r.font.size, r.font.color.rgb = Pt(12), RGBColor(194, 65, 12)
        elif l_s.startswith('### '):
            p = doc.add_paragraph(); p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(8), Pt(2)
            r = p.add_run(l_s.replace('### ', '')); r.bold, r.font.size = True, Pt(10.5)
        elif l_s.startswith('- ') or l_s.startswith('* ') or l_s.startswith('• '):
            c_t = l_s.lstrip('-*• ').strip()
            p = doc.add_paragraph(style='List Bullet'); p.paragraph_format.space_after = Pt(3)
            if ':' in c_t and not c_t.strip().startswith('http'):
                pk_s, zk_s = c_t.split(':', 1)
                if len(pk_s) < 45: p.add_run(pk_s.strip() + ': ').bold = True; parsuj_i_formatuj_tekst(p, zk_s); continue
            parsuj_i_formatuj_tekst(p, c_t)
        else:
            if ':' in l_s and not l_s.strip().startswith('http'):
                pk_s, zk_s = l_s.split(':', 1)
                if len(pk_s) < 45: 
                    p = doc.add_paragraph()
                    p.add_run(pk_s.strip() + ': ').bold = True
                    parsuj_i_formatuj_tekst(p, zk_s)
                    continue
            p = doc.add_paragraph(); parsuj_i_formatuj_tekst(p, l_s)
            
    b = BytesIO(); doc.save(b); return b.getvalue()

st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

with st.sidebar:
    st.header("🔑 Autoryzacja")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-3.5-flash", "gemini-3.1-pro"])

tab1, tab2 = st.tabs(["🚀 Automatyczny Transkrybent i Generator", "🎙️ Głosowy Edytor (Voice Editor)"])

with tab1:
    st.title("🐾 Multimedialny Hub MeatPoint.io")
    st.markdown("Wgraj plik wideo z Google Meet lub nagranie głosowe z konsultacji. Model asynchronicznie stworzy transkrypcję bez limitów wielkości i wygeneruje finalny protokół.")
    
    col1, col2 = st.columns([1, 1], gap="large")
    
    with col1:
        st.markdown("### 1️⃣ Wgranie pliku i Transkrypcja")
        media_file = st.file_uploader("📂 Wybierz plik audio lub wideo (.mp4, .mp3, .wav, .m4a):", type=["mp4", "mp3", "wav", "m4a"])
        
        if 'aktywna_transkrypcja' not in st.session_state: st.session_state.aktywna_transkrypcja = ""
        
        if st.button("🎙️ Uruchom inteligentną transkrypcję AI", type="secondary", use_container_width=True):
            if not api_key or not media_file: st.error("❌ Podaj klucz API oraz wgraj plik!")
            else:
                status_placeholder = st.empty()
                try:
                    genai.configure(api_key=api_key)
                    
                    with status_placeholder.container():
                        st.info("⏳ Krok 1/3: Przesyłanie dużego pliku do bezpiecznej chmury Google AI...")
                    
                    temp_filename = f"temp_{uuid.uuid4()}_{media_file.name}"
                    with open(temp_filename, "wb") as f:
                        f.write(media_file.getbuffer())
                    
                    uploaded_file_ref = genai.upload_file(path=temp_filename)
                    
                    # POPRAWKA: Algorytm Wykładniczego Wydłużania Czasu (Exponential Backoff) chroniący przed błędem 503
                    sleep_time = 5
                    total_waited = 0
                    while uploaded_file_ref.state.name == "PROCESSING":
                        with status_placeholder.container():
                            st.info(f"⏳ Krok 2/3: Trwa zaawansowana analiza audio przez Gemini... (Czekam już {total_waited}s, ponowne sprawdzenie za {sleep_time}s)")
                        time.sleep(sleep_time)
                        total_waited += sleep_time
                        # Stopniowo wydłużamy interwał zapytań do serwera
                        sleep_time = min(sleep_time * 1.5, 30) 
                        uploaded_file_ref = genai.get_file(uploaded_file_ref.name)
                    
                    if uploaded_file_ref.state.name == "FAILED":
                        raise Exception("Plik nie mógł zostać przetworzony przez serwery Google.")

                    with status_placeholder.container():
                        st.info("⏳ Krok 3/3: Generowanie pełnego tekstu transkrypcji z podziałem na role...")

                    model_transkrybent = genai.GenerativeModel(model_name="gemini-3.5-flash")
                    prompt_tr = (
                        "Przeanalizuj to nagranie audio/wideo z wizyty dietetycznej zwierzęcia. "
                        "Stwórz bardzo dokładną transkrypcję ortograficzną SŁOWO W SŁOWO. "
                        "Zastosuj wyraźny podział na role (Diarization), oznaczając kiedy mówi Ania (Dietetyk), "
                        "a kiedy właściciel zwierzęcia (Opiekun). Nie pomijaj żadnych nazw leków, dawek ani wyników badań."
                    )
                    
                    # Obsługa potencjalnego przeciążenia modelu przy samym generowaniu tekstu
                    try:
                        response_tr = model_transkrybent.generate_content([prompt_tr, uploaded_file_ref])
                        st.session_state.aktywna_transkrypcja = response_tr.text
                    except Exception as gemini_err:
                        if "503" in str(gemini_err) or "high demand" in str(gemini_err).lower():
                            with status_placeholder.container():
                                st.warning("⚠️ Serwery Google zgłaszają duże obciążenie. Automatyczna próba ponowienia za 15 sekund...")
                            time.sleep(15)
                            response_tr = model_transkrybent.generate_content([prompt_tr, uploaded_file_ref])
                            st.session_state.aktywna_transkrypcja = response_tr.text
                        else:
                            raise gemini_err
                    
                    # Czyszczenie śladów
                    try:
                        genai.delete_file(uploaded_file_ref.name)
                    except: pass
                    if os.path.exists(temp_filename): os.remove(temp_filename)
                        
                    status_placeholder.empty()
                    st.success("✅ Pełna transkrypcja wygenerowana pomyślnie!")
                    st.rerun()
                except Exception as e:
                    if os.path.exists(temp_filename): os.remove(temp_filename)
                    st.error(f"🚨 Status: Nie udało się dokończyć transkrypcji z powodu przeciążenia sieci Google. Spróbuj kliknąć przycisk ponownie za chwilę. Szczegóły: {e}")
        
        transcript = st.text_area("📝 Podgląd / Edycja tekstu transkrypcji:", value=st.session_state.aktywna_transkrypcja, height=380, key="transkrypcja_obszar")

    with col2:
        st.markdown("### 2️⃣ Budowanie gotowego dokumentu Word")
        LINK_DO_ARKUSZA = "https://docs.google.com/spreadsheets/d/1qgSX_t4_fb36CqtFUluPDKDQILpR9_SLOlYBPTXSTes/edit?usp=sharing"
        
        if st.button("🚀 Wygeneruj i uzupełnij Protokół Word", type="primary", use_container_width=True):
            if not api_key or not transcript: st.error("❌ Uzupełnij klucz API oraz upewnij się, że transkrypcja nie jest pusta!")
            else:
                with st.spinner("Analiza kliniczna i dopasowywanie załączników..."):
                    try:
                        csv_url = LINK_DO_ARKUSZA.replace('/edit?usp=sharing', '/export?format=csv')
                        df = pd.read_csv(csv_url); l_p = ""
                        for _, r in df.iterrows(): l_p += f"- Link: {r['URL']} | Tytuł: {r['Nazwa']} | Kiedy dołączyć (Wskazanie): {r['Opis dla AI']}\n"
                        
                        genai.configure(api_key=api_key)
                        m = genai.GenerativeModel(model_name=model_choice, system_instruction="Jesteś doświadczonym, pedantycznym asystentem klinicznym dla dietetyk Anny Michalskiej. Pisz wyłącznie prawdę na podstawie dostarczonego pliku tekstowego. Jeśli brakuje danych, bezwzględnie wstaw fragment [BRAK INFORMACJI]. Zakaz zmyślania preparatów celowanych.")
                        
                        instrukcja_szablonu = ""
                        for naglowek in STRUKTURA_PROTOKOLU:
                            if naglowek == "Załączniki:":
                                instrukcja_szablonu += f"## {naglowek}\n- Dołącz wyłącznie pasujące linki z bazy, jeśli ich warunki kliniczne zostały spełnione.\n- Pod nimi dodaj dokładnie te słowa:\nW razie pytań dotyczących tego opisu, jestem do Państwa dyspozycji.\nZachęcamy również do poszerzenia wiedzy o diecie na naszej stronie meatpoint.io lub Facebooku https://www.facebook.com/meatpoint.io\n\nPozdrawiam serdecznie,\nAnna Michalska"
                            elif naglowek == "Tyndalizacja:":
                                instrukcja_szablonu += f"## {naglowek}\n{TEKST_TYNDALIZACJA_STALY}\n\n"
                            elif naglowek == "Inne smaczki:":
                                instrukcja_szablonu += f"## {naglowek}\n{TEKST_INNE_SMACZKI_STALY}\n\n"
                            elif naglowek == "Wprowadzanie suplementów:":
                                instrukcja_szablonu += (
                                    f"## {naglowek}\n"
                                    "Proszę zacząć od:\n• Wody\n• Mięsa\n• Podrobów\n• tłuszczu\n• Tauryny\n"
                                    "Proszę przygotować dietę tylko z ich zawartością i na razie pominąć pozostałe suplementy.\n\n"
                                    "Jak Kicia będzie się dobrze czuła, na następny tydzień proszę przygotować dietę z zawartością:\n"
                                    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n"
                                    "• Dodatkowo: [Przeanalizuj transkrypcję. Wypisz po przecinku zalecane dodatki medyczne, które Anna wymieniła na ten krok diety. Jeśli w transkrypcji nie ma mowy o dodatkach celowanych dla tego tygodnia, wstaw sztywno tekst: [BRAK INFORMACJI]]\n\n"
                                    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
                                    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n"
                                    "• Dodatkowo: [Przeanalizuj transkrypcję. Wypisz po przecinku zalecane dodatki medyczne, które Anna wymieniła na ten krok diety. Jeśli w transkrypcji nie ma mowy o dodatkach celowanych dla tego tygodnia, wstaw sztywno tekst: [BRAK INFORMACJI]]\n\n"
                                    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
                                    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n"
                                    "• Dodatkowo: [Przeanalizuj transkrypcję. Wypisz po przecinku zalecane dodatki medyczne, które Anna wymieniła na ten krok diety. Jeśli w transkrypcji nie ma mowy o dodatkach celowanych dla tego tygodnia, wstaw sztywno tekst: [BRAK INFORMACJI]]\n\n"
                                    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
                                    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n"
                                    "• Dodatkowo: [Przeanalizuj transkrypcję. Wypisz po przecinku zalecane dodatki medyczne, które Anna wymieniła na ten krok diety. Jeśli w transkrypcji nie ma mowy o dodatkach celowanych dla tego tygodnia, wstaw sztywno tekst: [BRAK INFORMACJI]]\n\n"
                                    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
                                    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n• Jodu\n"
                                    "• Dodatkowo: [Przeanalizuj transkrypcję. Wypisz po przecinku zalecane dodatki medyczne, które Anna wymieniła na ten krok diety. Jeśli w transkrypcji nie ma mowy o dodatkach celowanych dla tego tygodnia, wstaw sztywno tekst: [BRAK INFORMACJI]]\n\n"
                                    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
                                    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n• Jodu\n"
                                    "• Dodatkowo: [Przeanalizuj transkrypcję. Wypisz po przecinku zalecane dodatki medyczne, które Anna wymieniła na ten krok diety. Jeśli w transkrypcji nie ma mowy o dodatkach celowanych dla tego tygodnia, wstaw sztywno tekst: [BRAK INFORMACJI]]\n\n"
                                    "To będzie już kompletna dieta."
                                )
                            else:
                                prefix = "" if naglowek.startswith("###") else "## "
                                instrukcja_szablonu += f"{prefix}{naglowek}\n- Uzupełnij precyzyjnymi faktami medycznymi z transkrypcji.\n"

                        p = f"Przeanalizuj podaną transkrypcję wizyty.\n\nWygeneruj dokument według tej rygorystycznej kolejności:\n\nKROK 1: Na samej górze stwórz wyrównaną DO LEWEJ linię: 'Data wizyty: DD.MM.YYYY' (wyciągnij datę lub wstaw [BRAK INFORMACJI])\n\nKROK 2: Bezpośrednio POD DATĄ wypisz linie metryczki podstawowej (ZAKAZ używania znaków '##' na ich początku):\nDane Opiekuna: \nPacjent: \nGatunek: \nRasa: \nWiek: \nWaga: \nBCS: \nIlość zwierząt w domu: \nSterylizacja/kastracja: \n\nKROK 3: Pod metryczką umieść poniższe nagłówki zachowując ich identyczną wielkość liter i pisownię:\n{instrukcja_szablonu}\n\n🚨 DEDYKOWANE DOPASOWANIE LINKÓW Z ARKUSZA:\nOto dostępna baza załączników zewnętrznych:\n{l_p}\n\nZAKAZ bezwarunkowego umieszczania linków. Przeanalizuj pole 'Kiedy dołączyć (Wskazanie)'. Dołącz dany adres URL do dokumentu TYLKO wtedy, gdy pacjent w transkrypcji cierpi na opisaną dolegliwość. Jeśli brak dopasowania, pomiń link.\n\nTranskrypcja:\n{transcript}"
                        
                        res = m.generate_content(p)
                        st.text_area("Podgląd tekstu wynikowego:", value=res.text, height=350, key="podglad_gen")
                        st.download_button("📥 POBIERZ GOTOWY PLIK WORD (.DOCX)", konwertuj_do_docx(res.text), "Protokol_MeatPoint.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    except Exception as e: st.error(f"🚨 Błąd generatora: {e}")

# ==============================================================================
# 🎙️ ZAKŁADKA 2: GŁOSOWY EDYTOR PROTOKOŁÓW
# ==============================================================================
with tab2:
    st.title("🎙️ Inteligentny Edytor Głosowy Protokółów")
    
    if 'sekcje_dokumentu' not in st.session_state: st.session_state.sekcje_dokumentu = None
    if 'koszyk_nagran' not in st.session_state: st.session_state.koszyk_nagran = {}
    if 'v_key' not in st.session_state: st.session_state.v_key = str(uuid.uuid4())
    if 'klucze_mikrofonow' not in st.session_state: st.session_state.klucze_mikrofonow = {}

    col_top1, col_top2 = st.columns([3, 1])
    with col_top1:
        u_file = st.file_uploader("📂 Wgraj plik protokołu (.docx):", type=["docx"], key=f"u_{st.session_state.v_key}")
    with col_top2:
        st.write("</br>", unsafe_allow_html=True)
        if st.button("🔄 Nowy protokół / Reset", type="secondary", use_container_width=True):
            st.session_state.sekcje_dokumentu = None
            st.session_state.koszyk_nagran = {}
            st.session_state.klucze_mikrofonow = {}
            st.session_state.v_key = str(uuid.uuid4())
            st.rerun()
            
    if u_file and st.session_state.sekcje_dokumentu is None:
        if st.button("⚙️ Załaduj strukturę pliku"):
            st.session_state.sekcje_dokumentu = segmentuj_docx(u_file.read()); st.rerun()

    if st.session_state.sekcje_dokumentu:
        st.markdown("---")
        col_ed1, col_ed2 = st.columns([1, 1], gap="large")
        
        with col_ed1:
            st.markdown("### 1️⃣ Wybór obszaru do korekty")
            wybrana_sekcja = st.selectbox("Wybierz nagłówek, do którego chcesz dodać nagranie:", list(st.session_state.sekcje_dokumentu.keys()), key="sel_voice")
            st.text_area("📄 Aktualna treść sekcji:", value=st.session_state.sekcje_dokumentu[wybrana_sekcja], height=220, disabled=True, key=f"t_{wybrana_sekcja}")
            
            if wybrana_sekcja not in st.session_state.klucze_mikrofonow:
                st.session_state.klucze_mikrofonow[wybrana_sekcja] = str(uuid.uuid4())
            
            mic_id = f"mic_{wybrana_sekcja}_{st.session_state.klucze_mikrofonow[wybrana_sekcja]}"
            audio_instrukcja = mic_recorder(start_prompt="🎙️ Nagraj uwagę dla tej sekcji", stop_prompt="🛑 Zatrzymaj i zapisz w pamięci", key=mic_id)
            
            if audio_instrukcja:
                if wybrana_sekcja not in st.session_state.koszyk_nagran or st.session_state.koszyk_nagran[wybrana_sekcja] != audio_instrukcja['bytes']:
                    st.session_state.koszyk_nagran[wybrana_sekcja] = audio_instrukcja['bytes']
                    st.rerun()

        with col_ed2:
            st.markdown("### 2️⃣ Lista zarejestrowanych uwag głosowych")
            if not st.session_state.koszyk_nagran:
                st.info("Brak oczekujących nagrań. Wybierz sekcję po lewej stronie i nagraj głos.")
            else:
                for s_nazwa in list(st.session_state.koszyk_nagran.keys()):
                    if s_nazwa in st.session_state.koszyk_nagran:
                        a_bytes = st.session_state.koszyk_nagran[s_nazwa]
                        c_box1, c_box2 = st.columns([5, 1])
                        c_box1.markdown(f"**📌 {s_nazwa}**")
                        c_box1.audio(a_bytes, format="audio/wav")
                        
                        if c_box2.button("❌", key=f"del_{s_nazwa}", help="Usuń to nagranie"):
                            if s_nazwa in st.session_state.koszyk_nagran:
                                del st.session_state.koszyk_nagran[s_nazwa]
                            st.session_state.klucze_mikrofonow[s_nazwa] = str(uuid.uuid4())
                            st.toast(f"🗑️ Usunięto nagranie z sekcji: {s_nazwa}")
                            st.rerun()
                
                st.markdown("---")
                if st.button("🚀 WPROWADŹ WSZYSTKIE POPRAWKI GŁOSOWE (HURTOWO)", type="primary", use_container_width=True):
                    if not api_key: st.error("❌ Podaj klucz API Gemini!")
                    else:
                        with st.spinner("Gemini edytuje wybrane fragmenty..."):
                            try:
                                genai.configure(api_key=api_key)
                                model_edytor = genai.GenerativeModel(model_name=model_choice)
                                
                                for s_nazwa, a_bytes in list(st.session_state.koszyk_nagran.items()):
                                    p_ed = f"Zmodyfikuj oryginalny tekst sekcji '{s_nazwa}' na podstawie instrukcji głosowych.\nTekst:\n{st.session_state.sekcje_dokumentu[s_nazwa]}\n\nZWROT WYŁĄCZNIE zaktualizowany tekst medyczny bez żadnych komentarzy ani wstępów."
                                    a_part = {"data": a_bytes, "mime_type": "audio/wav"}
                                    response_edycja = model_edytor.generate_content([p_ed, a_part])
                                    st.session_state.sekcje_dokumentu[s_nazwa] = response_edycja.text.strip()
                                    st.session_state.klucze_mikrofonow[s_nazwa] = str(uuid.uuid4())
                                
                                st.success("🎉 Wszystkie sekcje zostały pomyślnie zaktualizowane!")
                                st.rerun()
                            except Exception as e: st.error(f"🚨 Błąd edytora: {e}")

        st.markdown("---")
        st.markdown("### 3️⃣ Pobieranie gotowego dokumentu")
        if st.button("📦 Generuj finalny plik Word z poprawkami", type="primary", key="btn_build_final"):
            t_md = ""
            for sk, ts in st.session_state.sekcje_dokumentu.items():
                if sk in ["Nagłówek i Metryczka", "Nagłówek i Data wizyty"]: t_md += f"{ts}\n\n"
                else:
                    prefix = "" if sk.startswith("###") else "## "
                    t_md += f"{prefix}{sk}\n{ts}\n\n"
            st.download_button("📥 POBIERZ PROTOKÓŁ (.DOCX)", konwertuj_do_docx(t_md), "Protokol_MeatPoint_Poprawiony.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
