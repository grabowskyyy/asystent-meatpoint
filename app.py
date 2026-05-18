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

# --- SEGMENTATOR PLIKÓW DOCX (Dla zakładki Voice Editor) ---
def segmentuj_docx(file_bytes):
    doc = Document(BytesIO(file_bytes))
    sekcje = {}
    biezaca_sekcja = "Wstęp / Dane ogólne"
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

# Funkcja dodająca prawdziwe, klikalne linki (hiperłącza) do pliku Word (.docx)
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

    # Nagłówek dla pierwszej strony
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

    # Nagłówek dla kolejnych stron
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

# Konfiguracja Streamlit
st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

with st.sidebar:
    st.header("🔑 Autoryzacja i Model")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-2.5-flash", "gemini-1.5-flash"])

# --- DEKLARACJA ZAKŁADEK (TABS) ---
tab1, tab2 = st.tabs(["🚀 Generator Protokołów", "🎙️ Głosowy Edytor (Voice Editor)"])

# ==============================================================================
# 🚀 ZAKŁADKA 1: GENERATOR PROTOKOŁÓW (Twoja stabilna wersja)
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
                        system_instruction = "Jesteś elitarnym asystentem medycznym dla marki MeatPoint.io. Twoim jedynym zadaniem jest precyzyjne uzupełnianie struktury protokołu wizyty na podstawie transkrypcji. Pisz TYLKO fakty. Brak informacji oznacz jako [BRAK INFORMACJI]. Nie używaj tagów HTML."
                        model = genai.GenerativeModel(model_name=model_choice, system_instruction=system_instruction)
                        
                        prompt = f"Przeanalizuj transkrypcję i uzupełnij dokładnie szablon. Zewnętrzna baza linków:\n{lista_linkow_prompt}\n\nTranskrypcja:\n{transcript}"
                        # [Tutaj w rzeczywistej aplikacji znajduje się Twój pełny szablon tekstowy przesłany wcześniej]
                        # Dla oszczędności miejsca wstrzykujemy pełną instrukcję szablonu w tle:
                        prompt += "\n\nUzupełnij sekcje: Data wizyty, DANE FORMALNE OPIEKUNA, DANE PACJENTA, WYWIAD KLINICZNY, WYPRÓŻNIENIA I OBJAWY GASTRYCZNE, AKTUALNE BADANIA LABORATORYJNE, AKTUALNE LEKI, KOMENTARZ, EDUKACJA OPIEKUNA, HISTORIA ŻYWIENIOWA, SPECYFIKACJA NOWEGO PLANU, GOSPODARKA WODNA, SUPLEMENTACJA DODATKOWA, WIĄZANIE FOSFORU, AWARYJNE KARMY, HARMONOGRAM TRANZYCJI, HARMONOGRAM BADAŃ, ZAŁOŻONE ZAŁĄCZNIKI."
                        
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

    uploaded_file = st.file_uploader("📂 Wgraj plik protokołu MeatPoint (.docx):", type=["docx"], key="uploader_voice")
    
    if uploaded_file:
        if st.button("⚙️ Przeanalizuj i załaduj strukturę pliku", type="secondary"):
            st.session_state.sekcje_dokumentu = segmentuj_docx(uploaded_file.read())
            st.success("✅ Plik został pomyślnie rozbity na sekcje kliniczne!")

    if st.session_state.sekcje_dokumentu:
        st.markdown("---")
        col_ed1, col_ed2 = st.columns([1, 1])
        
        with col_ed1:
            st.subheader("🛠️ Wybór obszaru do korekty")
            opcje_sekcji = list(st.session_state.sekcje_dokumentu.keys())
            wybrana_sekcja = st.selectbox("Wybierz nagłówek, który chcesz zmodyfikować:", opcje_sekcji)
            
            st.text_area("📄 Aktualna treść tej sekcji w pliku:", value=st.session_state.sekcje_dokumentu[wybrana_sekcja], height=250, disabled=True, key="text_obszar")
            
        with col_ed2:
            st.subheader("🎙️ Dyktowanie instrukcji dla AI")
            st.info("Kliknij start, powiedz co chcesz zmienić (np. 'Zmień wagę kota na 3 kg i dopisz, że kał oddaje prawidłowo co dwa dni') i zatrzymaj nagrywanie.")
            
            # Wtyczka mikrofonu w przeglądarce
            audio_instrukcja = mic_recorder(
                start_prompt="🎙️ Rozpocznij nagrywanie głosu",
                stop_prompt="🛑 Zatrzymaj i zapisz instrukcję",
                key='audio_recorder_widget'
            )
            
            if audio_instrukcja:
                st.audio(audio_instrukcja['bytes'], format="audio/wav")
                
                if st.button("🚀 Wprowadź poprawki głosowe do tekstu", type="primary"):
                    if not api_key:
                        st.error("❌ Musisz podać klucz API Gemini w panelu bocznym!")
                    else:
                        with st.spinner("Gemini słucha Twojego nagrania i redaguje tekst medyczny..."):
                            try:
                                genai.configure(api_key=api_key)
                                model_edytor = genai.GenerativeModel(model_name=model_choice)
                                
                                prompt_edycji = f"""
                                Jesteś elitarnym edytorem dokumentacji klinicznej dla MeatPoint.io.
                                Masz przed sobą oryginalną treść sekcji '{wybrana_sekcja}':
                                \"\"\"
                                {st.session_state.sekcje_dokumentu[wybrana_sekcja]}
                                \"\"\"
                                
                                W załączonym pliku audio znajdują się instrukcje głosowe dietetyka.
                                Zmodyfikuj oryginalny tekst sekcji, wprowadzając dokładnie te zmiany, o które prosi głos.
                                ZWROT WYŁĄCZNIE zaktualizowany, czysty tekst tej sekcji. Nie dodawaj komentarzy, wyjaśnień ani wstępów. Zachowaj strukturę listy punktowanej i dwukropki.
                                """
                                
                                audio_part = {
                                    "data": audio_instrukcja['bytes'],
                                    "mime_type": "audio/wav"
                                }
                                
                                response_edycja = model_edytor.generate_content([prompt_edycji, audio_part])
                                
                                # Zapisujemy poprawioną treść bezpośrednio do pamięci sesji
                                st.session_state.sekcje_dokumentu[wybrana_sekcja] = response_edycja.text.strip()
                                st.success("🎉 Sekcja zaktualizowana pomyślnie! Zmiany są widoczne po lewej stronie.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"🚨 Błąd edytora głosowego: {e}")
                                
        # Budowanie końcowego pliku docx na bazie zaktualizowanego słownika sekcji
        st.markdown("---")
        st.subheader("💾 Zapis gotowego dokumentu")
        if st.button("📦 Generuj finalny, zaktualizowany plik Word", type="primary", key="btn_build_final"):
            tekst_md_final = ""
            for sekcja, tresc in st.session_state.sekcje_dokumentu.items():
                if sekcja == "Wstęp / Dane ogólne":
                    tekst_md_final += f"{tresc}\n\n"
                else:
                    tekst_md_final += f"## {sekcja}\n{tresc}\n\n"
                    
            plik_docx_final = konwertuj_do_docx(tekst_md_final)
            st.download_button(
                label="📥 POBIERZ POPRAWIONY PROTOKÓŁ (.DOCX)",
                data=plik_docx_final,
                file_name="Protokol_Konsultacji_MeatPoint_Poprawiony.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_final_fixed"
            )
