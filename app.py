import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from streamlit_mic_recorder import mic_recorder

# Funkcja generująca docx
def konwertuj_do_docx(tekst):
    doc = Document()
    for linia in tekst.split('\n'):
        if linia.startswith("## "): doc.add_paragraph(linia.replace("## ", "")).bold = True
        else: doc.add_paragraph(linia)
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

st.set_page_config(layout="wide")
api_key = st.sidebar.text_input("Klucz API", type="password")
model = st.sidebar.selectbox("Model", ["gemini-3.5-flash", "gemini-3.1-pro"])

tab1, tab2 = st.tabs(["🚀 Generator", "🎙️ Edytor"])

with tab1:
    trans = st.text_area("Transkrypcja", height=300)
    if st.button("Generuj"):
        genai.configure(api_key=api_key)
        m = genai.GenerativeModel(model)
        resp = m.generate_content(f"Uzupełnij protokół:\n{trans}")
        st.download_button("Pobierz", konwertuj_do_docx(resp.text), "protokol.docx")

with tab2:
    file = st.file_uploader("Wgraj .docx", type=["docx"])
    if file:
        if 'sekcje' not in st.session_state: st.session_state.sekcje = {"Treść": file.getvalue().decode('latin-1', errors='ignore')}
        sekcja = st.selectbox("Sekcja", list(st.session_state.sekcje.keys()))
        audio = mic_recorder(key="mic")
        if audio and st.button("Popraw"):
            m = genai.GenerativeModel(model)
            res = m.generate_content(["Popraw tekst zgodnie z nagraniem", audio])
            st.session_state.sekcje[sekcja] = res.text
            st.rerun()
