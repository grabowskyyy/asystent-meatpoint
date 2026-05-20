import streamlit as st, google.generativeai as genai, re, pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from streamlit_mic_recorder import mic_recorder

# --- SEGMENTATOR DOCX ---
def segmentuj_docx(file_bytes):
    doc = Document(BytesIO(file_bytes))
    sekcje = {}; biezaca = "Nagłówek"; sekcje[biezaca] = []
    naglowki = ["DANE FORMALNE OPIEKUNA", "DANE PACJENTA", "WYWIAD KLINICZNY", "WYPRÓŻNIENIA", "AKTUALNE BADANIA", "AKTUALNE LEKI", "KOMENTARZ", "EDUKACJA OPIEKUNA", "HISTORIA ŻYWIENIOWA", "SPECYFIKACJA PLANU", "GOSPODARKA WODNA", "SUPLEMENTACJA", "WIĄZANIE FOSFORU", "AWARYJNE KARMY", "HARMONOGRAM TRANZYCJI", "HARMONOGRAM BADAŃ", "ZAŁĄCZNIKI"]
    for p in doc.paragraphs:
        tekst = p.text.strip()
        if not tekst: continue
        znaleziono = False
        for n in naglowki:
            if n in tekst.upper() and len(tekst) < 65:
                biezaca = n; sekcje[biezaca] = []; znaleziono = True; break
        if not znaleziono: sekcje[biezaca].append(tekst)
    return {k: "\n".join(v) for k, v in sekcje.items()}

# --- KONWERTOR DOCX ---
def konwertuj_do_docx(tekst):
    doc = Document()
    for linia in tekst.split('\n'):
        if linia.startswith("## "): doc.add_paragraph(linia.replace("## ", "")).bold = True
        else: doc.add_paragraph(linia)
    bio = BytesIO(); doc.save(bio); return bio.getvalue()

st.set_page_config(layout="wide")
api = st.sidebar.text_input("Klucz API", type="password")
model = st.sidebar.selectbox("Model", ["gemini-3.5-flash", "gemini-3.1-pro"])

tab1, tab2 = st.tabs(["🚀 Generator", "🎙️ Edytor Głosowy"])

with tab1:
    col1, col2 = st.columns(2)
    trans = col1.text_area("Transkrypcja", height=400)
    if col2.button("Generuj Protokół"):
        genai.configure(api_key=api)
        m = genai.GenerativeModel(model)
        resp = m.generate_content(f"Uzupełnij protokół MeatPoint:\n{trans}")
        col2.download_button("📥 Pobierz .docx", konwertuj_do_docx(resp.text), "protokol.docx")

with tab2:
    # --- PRZYCISK RESETUJĄCY ---
    if st.button("🔄 Nowy protokół / Resetuj edytor"):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

    file = st.file_uploader("Wgraj .docx do edycji", type=["docx"])
    if file and 'sekcje' not in st.session_state:
        st.session_state.sekcje = segmentuj_docx(file.read())
    
    if st.session_state.sekcje:
        sekcja = st.selectbox("Wybierz sekcję", list(st.session_state.sekcje.keys()))
        st.text_area("Treść sekcji:", value=st.session_state.sekcje[sekcja], height=200, disabled=True)
        
        audio = mic_recorder(start_prompt="🎙️ Nagraj poprawkę", stop_prompt="🛑 Stop")
        if audio and st.button("Wprowadź głosową poprawkę"):
            m = genai.GenerativeModel(model)
            res = m.generate_content(["Popraw ten fragment tekstu zgodnie z instrukcją głosową:", st.session_state.sekcje[sekcja], audio])
            st.session_state.sekcje[sekcja] = res.text
            st.success("Zaktualizowano!"); st.rerun()
            
        if st.button("📦 Pobierz poprawiony plik"):
            tekst_finalny = "\n\n".join([f"## {k}\n{v}" for k, v in st.session_state.sekcje.items()])
            st.download_button("Pobierz .docx", konwertuj_do_docx(tekst_finalny), "protokol_poprawiony.docx")
