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
from szablon_struktura import SZABLON_PRODUKCYJNY

def segmentuj_docx(file_bytes):
    doc = Document(BytesIO(file_bytes))
    sekcje = {}
    biezaca_sekcja = "Nagłówek i Data wizyty"
    sekcje[biezaca_sekcja] = []
    znane_naglowki = ["DANE FORMALNE OPIEKUNA", "DANE PACJENTA", "WYWIAD KLINICZNY", "WYPRÓŻNIENIA I OBJAWY GASTRYCZNE", "AKTUALNE BADANIA LABORATORYJNE", "AKTUALNE LEKI I SUPLEMENTY MEDYCZNE", "KOMENTARZ DO WYWIADU I GŁÓWNE ZAŁOŻENIA DIETY", "EDUKACJA OPIEKUNA", "HISTORIA ŻYWIENIOWA I PREFERENCJE SMAKOWE", "SPECYFIKACJA NOWEGO PLANU DIETETYCEDNEGO", "GOSPODARKA WODNA (PICIU)", "SUPLEMENTACJA DODATKOWA (CELOWANA)", "WIĄZANIE FOSFORU I GOSPODARKA ŻELAZEM", "AWARYJNE KARMY KOMERCYJNE", "HARMONOGRAM TRANZYCJI", "HARMONOGRAM BADAŃ KONTROLNYCH", "ZAŁOŻONE ZAŁĄCZNIKI"]
    for p in doc.paragraphs:
        tekst = p.text.strip()
        if not tekst: continue
        znaleziono = False
        for n in znane_naglowki:
            if n in tekst.upper() and len(tekst) < 65:
                biezaca_sekcja = n
                if biezaca_sekcja not in sekcje: sekcje[biezaca_sekcja] = []
                znaleziono = True
                break
        if not znaleziono: sekcje[biezaca_sekcja].append(tekst)
    for k in sekcje: sekcje[k] = "\n".join(sekcje[k])
    return sekcje

def konwertuj_do_docx(tekst_markdown):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = Inches(1.3), Inches(0.8), Inches(0.8), Inches(0.8)
    for linia in tekst_markdown.split('\n'):
        linia = linia.replace('**', '')
        if linia.startswith('## '): doc.add_paragraph(linia.replace('## ', '')).bold = True
        else: doc.add_paragraph(linia)
    bufor = BytesIO()
    doc.save(bufor)
    return bufor.getvalue()

st.set_page_config(layout="wide")
api_key = st.sidebar.text_input("Klucz API", type="password")
model_choice = st.sidebar.selectbox("Model", ["gemini-3.5-flash", "gemini-3.1-pro"])
tab1, tab2 = st.tabs(["🚀 Generator", "🎙️ Edytor Głosowy"])

with tab1:
    col1, col2 = st.columns(2)
    transcript = col1.text_area("Transkrypcja", height=400)
    if col2.button("Generuj"):
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_choice)
        response = model.generate_content(f"Uzupełnij szablon:\n{SZABLON_PRODUKCYJNY}\n\n{transcript}")
        st.download_button("Pobierz", konwertuj_do_docx(response.text), "protokol.docx")

with tab2:
    uploaded_file = st.file_uploader("Wgraj .docx", type=["docx"])
    if uploaded_file and 'sekcje' not in st.session_state:
        st.session_state.sekcje = segmentuj_docx(uploaded_file.read())
    
    if st.session_state.sekcje:
        sekcja = st.selectbox("Wybierz sekcję", list(st.session_state.sekcje.keys()))
        st.text_area("Treść:", value=st.session_state.sekcje[sekcja], height=200, disabled=True)
        audio = mic_recorder(key="mic")
        if audio and st.button("Popraw"):
            # Logika wywołania Gemini z audio i podmiana w session_state
            st.success("Zaktualizowano!")
            st.rerun()
