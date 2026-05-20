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
    biezaca_sekcja = "Nagłówek i Data wizyty"
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

# Funkcja dodająca klikalne hiperłącza do pliku Word
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

    # Nagłówek dla pierwszej strony (Dane + Logo)
    tabela_naglowka = section.first_page_header.add_table(1, 2, Inches(6.7))
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

    # Nagłówek dla kolejnych stron (Samo Logo)
    subsequent_header
