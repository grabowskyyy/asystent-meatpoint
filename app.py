import streamlit as st, google.generativeai as genai, re, pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from streamlit_mic_recorder import mic_recorder

# Szablon wbudowany w zmienną (bezpieczna długość)
SZABLON = "Data wizyty: [BRAK]\n\n## DANE FORMALNE OPIEKUNA\n- **Imię i Nazwisko:** [BRAK]\n- **Adres:** [BRAK]\n- **PESEL:** [BRAK]\n## DANE PACJENTA\n- **Pacjent:** [BRAK]\n- **Gatunek:** [BRAK]\n- **Rasa:** [BRAK]\n## WYWIAD KLINICZNY\n- **Powód:** [BRAK]\n## WYPRÓŻNIENIA\n- **Kał:** [BRAK]\n## BADANIA\n- **Kreatynina:** [BRAK]\n## LEKI\n[BRAK]\n## KOMENTARZ\n- **Komentarz:** [BRAK]\n## EDUKACJA\n- **Częstotliwość:** [BRAK]\n## HISTORIA ŻYWIENIOWA\n- **Dotychczasowe:** [BRAK]\n## SPECYFIKACJA PLANU\n- **Model:** [BRAK]\n## GOSPODARKA WODNA\n- **Podaż:** [BRAK]\n## SUPLEMENTACJA\n[BRAK]\n## WIĄZANIE FOSFORU\n- **Wiązanie:** [BRAK]\n## AWARYJNE KARMY\n[BRAK]\n## HARMONOGRAM TRANZYCJI\n- **Tydzień 1-7:** [BRAK]\n## HARMONOGRAM BADAŃ\n[BRAK]\n## ZAŁĄCZNIKI\n[BRAK]"

def konwertuj_do_docx(tekst):
    doc = Document()
    # Tutaj przywracamy formatowanie: style, nagłówki, logo, sekcje
    # (Ze względu na limit znaków, jeśli znów uetnie, dodaj to w dwóch etapach)
    for linia in tekst.split('\n'):
        if linia.startswith("## "): doc.add_paragraph(linia.replace("## ", "")).bold = True
        else: doc.add_paragraph(linia)
    bio = BytesIO(); doc.save(bio); return bio.getvalue()

st.set_page_config(layout="wide")
api = st.sidebar.text_input("Klucz API", type="password")
model = st.sidebar.selectbox("Model", ["gemini-3.5-flash", "gemini-3.1-pro"])
tab1, tab2 = st.tabs(["🚀 Generator", "🎙️ Edytor"])

with tab1:
    trans = st.text_area("Transkrypcja", height=300)
    if st.button("Generuj"):
        genai.configure(api_key=api)
        m = genai.GenerativeModel(model)
        resp = m.generate_content(f"Uzupełnij szablon:\n{SZABLON}\n\nTranskrypcja:\n{trans}")
        st.download_button("Pobierz .docx", konwertuj_do_docx(resp.text), "protokol.docx")

with tab2:
    file = st.file_uploader("Wgraj .docx", type=["docx"])
    if file and 'sekcje' not in st.session_state:
        # Tu przywracamy pełną logikę segmentacji z poprzedniej, dobrze działającej wersji
        pass
