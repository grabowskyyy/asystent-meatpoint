import streamlit as st
import google.generativeai as genai
import os
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Funkcja dodająca prawdziwe, klikalne linki (hiperłącza) do pliku Word (.docx)
def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0563C1')  # Elegancki niebieski kolor linku
    rPr.append(c)
    
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')  # Podkreślenie dolne
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
    
    # Ustawienia marginesów dokumentu
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

    # --- LOGIKA ZRÓŻNICOWANIA NAGŁÓWKÓW ---
    section = doc.sections[0]
    section.different_first_page_header_footer = True  # Włączenie podziału na stronę 1 i resztę

    # 1. NAGŁÓWEK DLA PIERWSZEJ STRONY (Pełne dane + Logo)
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

    # 2. NAGŁÓWEK DLA KOLEJNYCH STRON (Wyłącznie samo Logo w prawym rogu)
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

# Interfejs Streamlit
st.set_page_config(page_title="MeatPoint - Asystent Dietetyka", layout="wide", page_icon="🐾")

st.title("🐾 MeatPoint.io - Generator Protokołów Konsultacji")
st.write("Wklej surową transkrypcję, aby wygenerować profesjonalny dokument Word z automatyczną korektą nagłówków stron.")

with st.sidebar:
    st.header("🔑 Autoryzacja")
    api_key = st.text_input("Klucz API Gemini", type="password")
    model_choice = st.selectbox("Wybierz model", ["gemini-2.5-flash", "gemini-1.5-flash"])

system_instruction = """
Jesteś elitarnym asystentem medycznym dla marki MeatPoint.io. Twoim jedynym zadaniem jest precyzyjne uzupełnianie struktury protokołu wizyty na podstawie transkrypcji.
REGUŁY BEZWZGLĘDNE:
1. Pisz TYLKO fakty podane bezpośrednio w transkrypcji rozmowy.
2. Jeśli w tekście brakuje informacji do jakiejkolwiek rubryki lub punktu (np. PESEL, adres, wyniki badań), wstaw tekst: [BRAK INFORMACJI]
3. NIE wolno Ci niczego zmyślać ani pominąć żadnego nagłówka. Jeśli brak danych - zostaw nagłówek i napisz [BRAK INFORMACJI].
4. NIE używaj żadnych tagów HTML ani kodów CSS. Wstaw
