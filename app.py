import streamlit as st, google.generativeai as genai, os, re, pandas as pd, uuid
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

# Stałe bloki edukacyjne z właściwymi, działającymi linkami MeatPoint
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
    "na podstawie danych z etykiety, znają Państwo w naszym artykule: "
    "https://meatpoint.io/pl/barf-wiedza/smaczki-i-dodatkowe-kalorie-obliczanie-kalorycznosci-komercyjnych-produktow"
)

TEKST_WPROWADZANIE_SUPLEMENTOW_STALY = (
    "Proszę zacząć od:\n• Wody\n• Mięsa\n• Podrobów\n• tłuszczu\n• Tauryny\n"
    "Proszę przygotować dietę tylko z ich zawartością i na razie pominąć pozostałe suplementy.\n\n"
    "Jak Kicia będzie się dobrze czuła, na następny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n• Jodu\n• Dodatkowo: \n\n"
    "Jak wszystko będzie w porządku za kolejny tydzień proszę przygotować dietę z zawartością:\n"
    "• Wody\n• Mięsa\n• Podrobów\n• Tłuszczu / żółtka\n• Tauryny\n• Wapnia/soli\n• Kwasów omega\n• Witaminy E\n• Witamin B\n• Jodu\n• Dodatkowo: \n\n"
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
            sub_segs = part.split('**')
            for idx, sub_seg in enumerate(sub_segs):
                if not sub_seg: continue
                czy_pogrubiony = (idx % 2 == 1)
                url_segs = re.split(r'(https?://[^\s]+)', sub_seg)
                for u_idx, u_seg in enumerate(url_segs):
                    if u_idx % 2 == 1:
                        add_hyperlink(p, u_seg, u_seg)
                    else:
                        run = p.add_run(u_seg)
                        if czy_pogrubiony:
                            run.bold = True
                            
        if i < len(parts) - 1:
            ra = p.add_run('[BRAK INFORMACJI]')
            ra.bold = True
            ra.font.color.rgb = RGBColor(220, 38, 38)

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
        
        if "DATA WIZYTY:" in l_s.upper():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(12), Pt(6)
            poczatek_daty, *koniec_daty = l_s.split(':', 1)
            poczatek_czysty = poczatek_daty.replace('**', '').strip()
            p.add_run(poczatek_czysty + ': ').bold = True
            if koniec_daty: parsuj_i_formatuj_tekst(p, koniec_daty[0].strip())
            continue

        if l_s.startswith('## '):
            p = doc.add_paragraph(); p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(14), Pt(4)
            czysty_h2 = l_s.replace('## ', '').replace('**', '')
            r = p.add_run(czysty_h2); r.bold = True; r.font.size, r.font.color.rgb = Pt(12), RGBColor(194, 65, 12)
        elif l_s.startswith('### '):
            p = doc.add_paragraph(); p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(8), Pt(2)
            czysty_h3 = l_s.replace('### ', '').replace('**', '')
            r = p.add_run(czysty_h3); r.bold, r.font.size = True, Pt(10.5)
        elif l_s.startswith('- ') or l_s.startswith('* ') or l_s.startswith('• '):
            c_t = l_s.lstrip('-*• ').strip()
            p = doc.add_paragraph(style='List Bullet'); p.paragraph_format.space_after = Pt(3)
            if ':' in c_t and not c_t.strip().startswith('http'):
                pk_s, zk_s = c_t.split(':', 1)
                if len(pk_s) < 45: 
                    pk_czysty = pk_s.replace('**', '').strip()
                    p.add_run(pk_czysty + ': ').bold = True
                    parsuj_i_formatuj_tekst(p, zk_s)
                    continue
            parsuj_i_formatuj_tekst(p, c_t)
        else:
            if ':' in l_s and not l_s.strip().startswith('http'):
                pk_s, zk_s = l_s.split(':', 1)
                if len(pk_s) < 45: 
                    p = doc.add_paragraph()
                    pk_czysty = pk_s.replace('**', '').strip()
                    p.add_run(pk_czysty + ': ').bold = True
                    parsuj_i_formatuj_tekst(p, zk_s)
                    continue
            p = doc.add_paragraph(); parsuj_i_formatuj_tekst(p, l_s)
            
    b = BytesIO(); doc.save(b); return b.getvalue()

st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

with st.sidebar:
    st.header("🔑 Autoryzacja")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-1.5-flash", "gemini-1.5-pro"])

tab1, tab2 = st.tabs(["🚀 Generator opisów wizyt (Wklej Tekst)", "🎙️ Edytor głosowy opisów wizyt"])

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
                                "1. Przeanalizuj treść każdego załącznika. Informacje w nich zawarte mogą dotyczyć DOWOLNEY sekcji protokołu (notatki o wodzie, uwagi o smaczkach, dawki leków, opisy samopoczucia, wyniki badań).\n"
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

# ==============================================================================
# 🎙️ ZAKŁADKA 2: GŁOSOWY EDYTOR PROTOKOŁÓW
# ==============================================================================
with tab2:
    st.title("🎙️ Edytor głosowy opisów wizyt")
    
    if 'sekcje_dokumentu' not in st.session_state: st.session_state.sekcje_dokumentu = None
    if 'koszyk_nagran' not in st.session_state: st.session_state.koszyk_nagran = {}
    if 'v_key' not in st.session_state: st.session_state.v_key = str(uuid.uuid4())
    if 'klucze_mikrofonow' not in st.session_state: st.session_state.klucze_mikrofonow = {}

    col_top1, col_top2 = st.columns([3, 1])
    with col_top1:
        u_file = st.file_uploader("📂 Wgraj plik opisu (.docx):", type=["docx"], key=f"u_{st.session_state.v_key}")
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
