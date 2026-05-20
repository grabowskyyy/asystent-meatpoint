import streamlit as st, google.generativeai as genai, os, re, pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from streamlit_mic_recorder import mic_recorder

# --- SZABLON (Zoptymalizowany dla GitHuba) ---
SZABLON = "Data wizyty: [BRAK] \n## DANE FORMALNE OPIEKUNA\n- **Imię i Nazwisko:** [BRAK]\n## DANE PACJENTA\n- **Pacjent:** [BRAK]\n## WYWIAD KLINICZNY\n- **Powód:** [BRAK]\n## WYPRÓŻNIENIA I OBJAWY GASTRYCZNE\n- **Kał:** [BRAK]\n## AKTUALNE BADANIA LABORATORYJNE\n- **Kreatynina:** [BRAK]\n## AKTUALNE LEKI I SUPLEMENTY\n[BRAK]\n## KOMENTARZ I GŁÓWNE ZAŁOŻENIA DIETY\n- **Komentarz:** [BRAK]\n## EDUKACJA OPIEKUNA: CO SIĘ ZMIENI\n- **Częstotliwość kału:** [BRAK]\n## HISTORIA ŻYWIENIOWA\n- **Dotychczasowe żywienie:** [BRAK]\n## SPECYFIKACJA NOWEGO PLANU\n- **Model:** [BRAK]\n## GOSPODARKA WODNA\n- **Podaż płynów:** [BRAK]\n## SUPLEMENTACJA DODATKOWA\n[BRAK]\n## WIĄZANIE FOSFORU I ŻELAZO\n- **Wiązanie:** [BRAK]\n## AWARYJNE KARMY\n[BRAK]\n## HARMONOGRAM TRANZYCJI\n- **Tydzień 1-7:** [BRAK]\n## HARMONOGRAM BADAŃ\n[BRAK]\n## ZAŁĄCZNIKI\n[BRAK]"

def segmentuj_docx(file_bytes):
    doc = Document(BytesIO(file_bytes))
    sekcje = {}; biezaca = "Nagłówek"; sekcje[biezaca] = []
    for p in doc.paragraphs:
        if p.text.strip(): sekcje[biezaca].append(p.text.strip())
    return {k: "\n".join(v) for k, v in sekcje.items()}

def konwertuj_do_docx(tekst):
    doc = Document()
    for linia in tekst.split('\n'):
        if linia.startswith("## "): doc.add_paragraph(linia.replace("## ", "")).bold = True
        else: doc.add_paragraph(linia)
    bio = BytesIO(); doc.save(bio); return bio.getvalue()

st.set_page_config(layout="wide")
api_key = st.sidebar.text_input("Klucz API", type="password")
model_choice = st.sidebar.selectbox("Model", ["gemini-3.5-flash", "gemini-3.1-pro"])
tab1, tab2 = st.tabs(["🚀 Generator Protokołów", "🎙️ Głosowy Edytor"])

with tab1:
    col1, col2 = st.columns([1, 1], gap="large")
    trans = col1.text_area("Transkrypcja", height=400)
    if col2.button("🚀 Generuj"):
        if api_key and trans:
            genai.configure(api_key=api_key)
            m = genai.GenerativeModel(model_choice)
            resp = m.generate_content(f"Wypełnij:\n{SZABLON}\n\nTranskrypcja:\n{trans}")
            col2.download_button("📥 Pobierz", konwertuj_do_docx(resp.text), "protokol.docx")

with tab2:
    if 'sekcje' not in st.session_state: st.session_state.sekcje = None
    file = st.file_uploader("Wgraj .docx", type=["docx"])
    if file and st.button("Analizuj plik"): st.session_state.sekcje = segmentuj_docx(file.read())
    
    if st.session_state.sekcje:
        sekcja = st.selectbox("Sekcja", list(st.session_state.sekcje.keys()))
        st.text_area("Treść:", value=st.session_state.sekcje[sekcja], disabled=True)
        audio = mic_recorder(start_prompt="🎙️ Nagraj poprawkę", stop_prompt="🛑 Stop")
        if audio and st.button("Wprowadź zmiany"):
            m = genai.GenerativeModel(model_choice)
            res = m.generate_content(["Popraw tekst zgodnie z audio", audio])
            st.session_state.sekcje[sekcja] = res.text
            st.rerun()
