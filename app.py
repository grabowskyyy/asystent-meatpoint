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

# --- DANE SZABLONU KLINICZNEGO MEATPOINT ---
SZABLON_PRODUKCYJNY = """
Data wizyty: [Wpisz datę wizyty lub BRAK INFORMACJI]

## DANE FORMALNE OPIEKUNA
- **Imię i Nazwisko:** [Imię i nazwisko opiekuna]
- **Adres zamieszkania:** [Adres zamieszkania opiekuna (ulica, kod, miasto)]
- **Numer PESEL:** [Numer PESEL opiekuna]
- **Numer telefonu:** [Numer telefonu kontaktowego]
- **Adres e-mail:** [Adres e-mail opiekuna]

## DANE PACJENTA
- **Pacjent:** [Imię zwierzęcia]
- **Gatunek:** [Kot/Pies]
- **Rasa:** [Rasa pacjenta]
- **Wiek:** [Wiek pacjenta]
- **Waga:** [Aktualna waga, tendencje wagowe i waga docelowa]
- **BCS:** [Ocena kondycji w skali 1-9/9 oraz krótki opis fizyczny sylwetki]
- **Ilość zwierząt w domu:** [Liczba zwierząt w stadzie, status pacjenta i relacje]
- **Sterylizacja/kastracja:** [Tak/Nie + rok i miejsce zabiegu]

## WYWIAD KLINICZNY
- **Powód konsultacji:** [Historia schorzenia, zdiagnozowane jednostki chorobowe np. PNN, IBD, zapalenie trzustki oraz główne oczekiwania i cele opiekuna wobec diety]
- **Aktualne samopoczucie:** [Zachowanie pacjenta, przebyte niedawno zabiegi np. sanacja jamy ustnej, stan po zabiegu]
- **Aktywność:** [Umiarkowana/duża, adekwatność do aktualnego stanu zdrowia]
- **Apetyt:** [Stan apetytu, częstotliwość podawania karmy w ciągu doby, historia wybredności]
- **Pragnienie:** [Ilość samodzielnego picia, częstotliwość, stosowane kroplówki - objętość dobowa i rodzaj płynów, plany redukcji płynoterapii]

## WYPRÓŻNIENIA I OBJAWY GASTRYCZNE
- **Kał:** [Częstotliwość na dobę, uformowanie, zapach, konsystencja, opis jelit z USG pod kątem zmian typowych dla IBD]
- **Wymioty:** [Częstotliwość występowania, po jakich pokarmach lub lekach]
- **Mocz:** [Barwa, klarowność, ciężar właściwy, proteinuria/białko, obecność erytrocytów, infekcje, dobowy schemat mikcji]
- **Odrobaczanie:** [Ostatnia data odrobaczania, powód, forma podania i preparat]

## AKTUALNE BADANIA LABORATORYJNE
[Zestawienie najnowszych wyników wraz z datami i interpretacją trendu klinicznego]:
- **Kreatynina:** [Wartość + jednostka + trend]
- **Mocznik:** [Wartość + jednostka + trend]
- **Fosfor:** [Wartość + jednostka + trend]
- **T4 całkowita:** [Wartość + jednostka + trend]
- **Morfologia (HGB / Anemia):** [Wartość HGB, stan układu czerwonokrwinkowego, diagnoza niedokrwistości]
- **Albuminy:** [Wartość + jednostka + trend]
- **&alpha;-amylaza:** [Wartość + jednostka + trend]
- **Cholesterol:** [Wartość + jednostka + trend]
- **WBC (Leukocyty):** [Wartość + stan zapalny/infekcja]
- **Gospodarka cukrowa (Fruktozamina):** [Wartość fruktozaminy, glukoza w moczu, wykluczenie/potwierzenie cukrzycy]

## AKTUALNE LEKI I SUPLEMENTY MEDYCZNE
[Pełna lista przyjmowanych preparatów, dawkowanie, częstotliwość i od kiedy są stosowane]

## KOMENTARZ DO WYWIADU I GŁÓWNE ZAŁOŻENIA DIETY
- **Komentarz:** [Podsumowanie stopnia trudności pacjenta, tolerancji składników i wymaganych kompromisów klinicznych między nerkami a przewodem pokarmowym]
- **Główne założenia diety:** [Kluczowe cele makroskładnikowe: poziom fosforu, jakość i strawność białka, poziomy tłuszczów i węglowodanów pod kątem trzustki, zasada stopniowego wdrażania]

## EDUKACJA OPIEKUNA: CO SIĘ ZMIENI NA DIECIE BARF/BACF
- **Częstotliwość kału:** Zwierzę może oddawać mniejszy kał i może go oddawać co 2–3 dni. Na wysokomięsnej diecie to normalne. Ważne, żeby był dobrego kształtu i konsystencji (wdł skali bristolskiej).
- **UWAGA NA ZAPARCIA:** Należy odróżnić rzadkie oddawanie kału od zaparć. Jeśli pacjent na diecie BARF będzie miał: suchą kupę, twardą, bobki / rodzynki / kamyczki, z dużą ilością włosa… to może być zaparcie lub do niego prowadzić. Nie chodzi o samą częstotliwość oddawania stolca, ale o jego wygląd i o zachowanie w kuwecie.
- **Parametry krwi:** Parametry nerkowe krwi na wysoko mięsnej diecie mogą się różnić od zdrowych zwierząt (nie tylko z powodu choroby nerek), zwłaszcza mocznik i kreatynina. W zależności od pozostałych parametrów i samopoczucia - nie oznacza od razu pogorszenia choroby nerek. Ważna jest stała kontrola u nefrologa: badanie USG, SDMA, badania moczu i stanu ogólnego, być może FGF-23 - zgodnie z zaleceniami lekarza.
- **Objętość posiłku:** Początkowo może się wydawać, że diety jest mało. Dieta BARF/BACF nerkowa jest bardziej kaloryczna i treściwa w mniejszej objętości niż puszki i saszetki. Przyzwyczajanie się do tej zmianzonej ilości może zająć ok. 2–3 miesiące i to jest normalne.

## HISTORIA ŻYWIENIOWA I PREFERENCJE SMAKOWE
- **Dotychczasowe żywienie:** [Opis dotychczasowych modeli żywienia, stosowane wcześniej przepisy, źródła białka, używane marki, stopień akceptacji i przyczyny rezygnacji/modyfikacji]
- **KATEGORYCZNIE TAK (Ulubione smaki):** [Lista akceptowanych rodzajów mięs, części tuszy, podrobów, warzyw i forma podania. UWAGA: Podkreśl czy je potrawy mrożone czy tylko świeże/z lodówki]
- **KATEGORYCZNIE NIE (Odrzucone składniki):** [Lista absolutnie odrzucanych przez zwierzę składników, mięs, form wapnia lub suplementów wywołujących wymioty, niechęć lub całkowity bunt]

## SPECYFIKACJA NOW
