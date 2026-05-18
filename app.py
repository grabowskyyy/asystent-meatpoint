import streamlit as st
import google.generativeai as genai
import os
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Funkcja do natywnego formatowania tekstu w Wordzie (obsługa pogrubień i czerwonego alertu)
def parsuj_i_formatuj_tekst(paragraph, tekst):
    # Dzielimy tekst według znaczników pogrubienia **
    czesci_bold = tekst.split('**')
    for index_bold, czesc_bold in enumerate(czesci_bold):
        is_bold = (index_bold % 2 == 1)
        
        # Wewnątrz każdej części sprawdzamy obecność znacznika [BRAK INFORMACJI]
        czesci_brak = czesc_bold.split('[BRAK INFORMACJI]')
        for index_brak, czesc_brak in enumerate(czesci_brak):
            if czesc_brak:
                run = paragraph.add_run(czesc_brak)
                if is_bold:
                    run.bold = True
            
            # Jeśli to nie jest ostatni element, oznacza to, że tu był [BRAK INFORMACJI]
            if index_brak < len(czesci_brak) - 1:
                run_alert = paragraph.add_run('[BRAK INFORMACJI]')
                run_alert.bold = True
                run_alert.font.color.rgb = RGBColor(220, 38, 38)  # Wyrazisty czerwony color

# Główna funkcja generująca zaawansowany plik .docx ze stylami MeatPoint
def konwertuj_do_docx(tekst_markdown):
    doc = Document()
    
    # Ustawienia marginesów (wąskie marginesy dla lepszego układu klinicznego)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Ustawienie globalnej czcionki Arial i interlinii 1.25
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10.5)
    style.paragraph_format.line_spacing = 1.25
    style.paragraph_format.space_after = Pt(4)

    # --- NAGŁÓWEK BRANDINGOWY (Tabela dwukolumnowa: Dane po lewej, Logo po prawej) ---
    tabela_naglowka = doc.add_table(rows=1, cols=2)
    tabela_naglowka.autofit = False
    
    # Brak obramowania tabeli
    tblPr = tabela_naglowka._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

    kol_lewa = tabela_naglowka.rows[0].cells[0]
    kol_prawa = tabela_naglowka.rows[0].cells[1]
    kol_lewa.width = Inches(4.5)
    kol_prawa.width = Inches(2.0)

    # Dane kontaktowe po lewej stronie nagłówka
    p_kontakt = kol_lewa.paragraphs[0]
    p_kontakt.paragraph_format.space_after = Pt(2)
    run_name = p_kontakt.add_run("Anna Michalska\n")
    run_name.bold = True
    run_name.font.size = Pt(12)
    
    run_sub = p_kontakt.add_run("Dietetyka Psów i Kotów\n")
    run_sub.font.size = Pt(10)
    run_sub.font.color.rgb = RGBColor(100, 116, 139) # Szary podtytuł
    
    run_det = p_kontakt.add_run("miesnepsokotki@gmail.com\nhttps://www.facebook.com/meatpoint.io")
    run_det.font.size = Pt(9.5)

    # Dodanie Logo po prawej stronie nagłówka (jeśli plik istnieje w repozytorium)
    if os.path.exists("logo.png"):
        p_logo = kol_prawa.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_logo.add_paragraph().add_run().add_picture("logo.png", width=Inches(1.3))
    
    # Linia odstępu pod nagłówkiem
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- PARSER LINII TEKSTU ---
    for linia in tekst_markdown.split('\n'):
        linia_strip = linia.strip()
        
        if not linia_strip:
            continue
            
        # Nagłówki Główne Sekcji (np. ## WYWIAD KLINICZNY) -> Kolor Ciemnopomarańczowy
        if linia_strip.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(linia_strip.replace('## ', ''))
            run.bold = True
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(194, 65, 12)  # Markowy pomarańcz MeatPoint (#C2410C)
            
        # Podnagłówki (###) -> Pogrubiony czarny tekst, lekko większy
        elif linia_strip.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(linia_strip.replace('### ', ''))
            run.bold = True
            run.font.size = Pt(11)
            
        # Listy punktowane -> Standardowy punktor Worda
        elif linia_strip.startswith('- ') or linia_strip.startswith('* '):
            czysty_tekst = linia_strip.lstrip('-* ').strip()
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(3)
            parsuj_i_formatuj_tekst(p, czysty_tekst)
            
        # Zwykły tekst akapitu
        else:
            p = doc.add_paragraph()
            parsuj_i_formatuj_tekst(p, linia_strip)
            
    bufor = BytesIO()
    doc.save(bufor)
    return bufor.getvalue()

# Konfiguracja Streamlit
st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

st.title("🐾 MeatPoint.io - Generator Protokołów Konsultacji")
st.write("Wklej surową transkrypcję, aby wygenerować profesjonalny dokument Word zgodny z identyfikacją wizualną marki.")

with st.sidebar:
    st.header("🔑 Autoryzacja")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-2.5-flash", "gemini-1.5-flash"])

# Rygorystyczny system instructions - usunięto stary kod HTML na rzecz czystego tagu
system_instruction = """
Jesteś elitarnym asystentem medycznym MeatPoint.io. Twoim zadaniem jest uzupełnianie protokołu wizyty.
REGUŁY:
1. Pisz TYLKO fakty podane w transkrypcji.
2. Jeśli brakuje jakiejkolwiek informacji, wstaw dokładnie tekst: [BRAK INFORMACJI]
3. NIE używaj żadnych tagów HTML (typu <span style...>). Wstawiaj wyłącznie czysty tekst [BRAK INFORMACJI].
"""

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("🔊 Surowa transkrypcja")
    transcript = st.text_area("Wklej tekst z Google Meet:", height=500)

with col2:
    st.subheader("📋 Wynikowy Protokół Wizyty")
    
    if st.button("🚀 Generuj i wypełnij szablon", type="primary"):
        if not api_key or not transcript:
            st.error("❌ Uzupełnij klucz API oraz transkrypcję!")
        else:
            with st.spinner("Przetwarzanie dokumentu..."):
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(model_name=model_choice, system_instruction=system_instruction)
                    
                    # Wstrzyknięcie pełnego szablonu do uzupełnienia
                    prompt = f"Przeanalizuj transkrypcję i uzupełnij dokładnie poniższy szablon structures:\n\n[Tutaj z poziomu kodu przesyłany jest Twój pełny szablon od sekcji DANE PACJENTA po HARMONOGRAM]\n\nTranskrypcja:\n{transcript}"
                    
                    response = model.generate_content(prompt)
                    
                    # Wyświetlenie czystego podglądu w aplikacji
                    st.text_area("Podgląd tekstu generowanego:", value=response.text, height=350)
                    
                    plik_docx = konwertuj_do_docx(response.text)
                    st.markdown("---")
                    st.download_button(
                        label="📥 POBIERZ PROFESJONALNY PLIK WORD (.DOCX)",
                        data=plik_docx,
                        file_name="Protokol_Konsultacji_MeatPoint.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"🚨 Błąd: {e}")
