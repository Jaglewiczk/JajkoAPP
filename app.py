"""Generator checklist magazynowych dla firmy cateringowej.

Ta aplikacja nie odczytuje plikow Excel podczas pracy. Slownik i reguly sa
zapisane ponizej w Pythonie. Excel byl tylko materialem do ich przygotowania.
"""

from __future__ import annotations

import io
import json
import math
import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from slownik_pelny import DANE_SLOWNIKA


# ---------------------------------------------------------------------------
# KONFIGURACJA GEMINI
# Zostawiamy klucz w kodzie, zgodnie z ustaleniem. Wklej tu swoj aktualny klucz.
# ---------------------------------------------------------------------------
API_KEY = "AQ.Ab8RN6KjmjHiNd63tN84AzhbA89sbacPadEtlBMH-m0D-ozJfw"
MODEL_NAME = "gemini-3.7-flash"

st.set_page_config(page_title="Generator checklist cateringowych", layout="wide")


@dataclass(frozen=True)
class Artykul:
    kod: str
    nazwa: str
    jednostka: str = "SZT"
    opakowanie: int = 1
    uwagi: str = ""


def artykul(kod: str, nazwa: str, jednostka: str = "SZT", opakowanie: int = 1, uwagi: str = "") -> Artykul:
    return Artykul(kod, nazwa, jednostka, opakowanie, uwagi)


# ---------------------------------------------------------------------------
# SLOWNIK ARTYKULOW (zapisany w kodzie, bez zaleznosci od .xlsx)
# Kolejne pozycje dopisuje sie jedna linia wedlug wzoru artykul(...).
# ---------------------------------------------------------------------------
_ARTYKULY = [
    artykul("STOL_BUFET", "STÓŁ BUFETOWY 150"),
    artykul("STOL_KOKTAJL", "STÓŁ KOKTAJLOWY 80"),
    artykul("NACIAG_BUFET_CZARNY", "NACIĄG CZARNY NA STÓŁ BUFETOWY 150"),
    artykul("NACIAG_BUFET_BIALY", "NACIĄG BIAŁY NA STÓŁ BUFETOWY 150"),
    artykul("NACIAG_BUFET_SZARY", "NACIĄG SZARY NA STÓŁ BUFETOWY 150"),
    artykul("NACIAG_BUFET_ZOLTY", "NACIĄG ŻÓŁTY NA STÓŁ BUFETOWY 150"),
    artykul("NACIAG_BUFET_NIEBIESKI", "NACIĄG NIEBIESKI NA STÓŁ BUFETOWY 150"),
    artykul("CZAPKA_BUFET", "CZAPKA CZARNA NA STÓŁ BUFETOWY 150"),
    artykul("NACIAG_KOKTAJL_CZARNY", "NACIĄG CZARNY NA STÓŁ KOKTAJLOWY 80"),
    artykul("NACIAG_KOKTAJL_BIALY", "NACIĄG BIAŁY NA STÓŁ KOKTAJLOWY 80"),
    artykul("CZAPKA_KOKTAJL", "CZAPKA CZARNA NA STÓŁ KOKTAJLOWY 80"),
    artykul("OBRUS_BIALY", "OBRUS 240/120 BIAŁY"),
    artykul("EKSPRES", "EKSPRES DO KAWY SAECO DUŻY", uwagi="Urządzenie elektryczne"),
    artykul("WARNIK_DUZY", "WARNIK DUŻY", uwagi="Urządzenie elektryczne"),
    artykul("PODGRZEWACZ_DANIE", "PODGRZEWACZ DO II DANIA", uwagi="Urządzenie elektryczne"),
    artykul("PODGRZEWACZ_ZUPA", "PODGRZEWACZ DO ZUPY", uwagi="Urządzenie elektryczne"),
    artykul("PRZEDLUZACZ", "PRZEDŁUŻACZ KLASYCZNY"),
    artykul("FILIŻANKA", "FILIŻANKA", opakowanie=48),
    artykul("SPODEK", "SPODEK", opakowanie=80),
    artykul("LYZECZKA", "ŁYŻECZKA MAŁA", opakowanie=10),
    artykul("KUBKI_PAPIEROWE", "$ KUBKI PAPIEROWE"),
    artykul("TALERZ_DESER", "TALERZ DESEROWY 17", opakowanie=80),
    artykul("TALERZ_PRZEKASKA", "TALERZ PRZEKĄSKA 19", opakowanie=50),
    artykul("TALERZ_OBIAD", "TALERZ OBIADOWY 24", opakowanie=38),
    artykul("BULIONOWKA", "BULIONÓWKA", opakowanie=96),
    artykul("WIDELEC", "WIDELEC", opakowanie=10),
    artykul("NOZ", "NÓŻ", opakowanie=10),
    artykul("LYZKA", "ŁYŻKA", opakowanie=10),
    artykul("WIDELCZYK", "WIDELCZYK MAŁY DESEROWY", opakowanie=10),
    artykul("KIELISZEK_WINO", "KIELISZKI WINO", opakowanie=36),
    artykul("KIELISZEK_PROSECCO", "KIELISZKI PROSECCO", opakowanie=49),
    artykul("KIELISZEK_WODKA", "KIELISZKI WÓDKA", opakowanie=49),
    artykul("LONG", "LONGI", opakowanie=49),
    artykul("MELAMINA", "MELAMINA 1/1"),
    artykul("ETAZERKA", "WIESZAK ETAŻERKA STOJAK"),
    artykul("SERWIS_CHOCHLA", "SERWIS CHOCHLA"),
    artykul("SERWIS_LYZKA", "SERWIS ŁYŻKA"),
    artykul("SERWIS_SZCZYPCE", "SERWIS SZCZYPCE"),
    artykul("PODSTAWKA_SERWIS", "PODSTAWKA POD SERWIS"),
    artykul("SEPARATOR", "SEPARATOR NA SZTUĆCE GN 1/1"),
    artykul("SERWETA_KELNERSKA", "MATERIAŁOWA KELNERKA SERWETA 44X44 POCKET"),
    artykul("STEND_MLEKO", "STEND MAŁY OPIS DO MLEKA A7"),
    artykul("KARAFKA_MLEKO", "KARAFKA SZKLANA 1L NA MLEKO DO EXP"),
    artykul("DZBANEK_MLEKO", "DZBANKI PORCELANOWE 350ML BIAŁY"),
    artykul("MISECZKA", "MAŁA MISECZKA NA DODATKI"),
    artykul("DYSPENSER_HERBATA", "DYSPENSER NA HERBATĘ"),
    artykul("DYSPENSER_NAPOJE", "DYSPENSER NA NAPOJE"),
    artykul("KAWA", "$ KAWA ZIARNO 1kg", "KG"),
    artykul("MLEKO", "$ MLEKO ZWYKŁE 1 LITR", "LITR"),
    artykul("MLEKO_BEZ", "$ MLEKO BEZ LAKTOZY 1 LITR", "LITR"),
    artykul("MLEKO_ROSLINNE", "$ MLEKO ROŚLINNE SOJOWE 1 LITR", "LITR"),
    artykul("WODA_5L", "$ WODA 5L (BANIAK)", "BANIAK"),
    artykul("CUKIER_BIALY", "$ ZAPAS CUKRU BIAŁEGO", "SZT"),
    artykul("CUKIER_BRAZOWY", "$ ZAPAS CUKRU BRĄZOWEGO", "SZT"),
    artykul("KWIATY_BUFET", "$ KWIATY BUFET"),
    artykul("KWIATY_KOKTAJL", "$ KWIATY NA KOKTAJLE"),
    artykul("CYTRYNA", "$ CYTRYNA SPODEK + WIDELCZYK"),
    artykul("PSIK", "$ PSIK PSIK"),
    artykul("REKAWICZKI", "$ RĘKAWICZKI JEDNORAZOWE", "PARA"),
    artykul("CZYSCIWO", "$ PAPIER CZYŚCIWO", "OP"),
    artykul("WORKI", "$ WORKI NA ŚMIECI CIEMNE"),
    artykul("WIADERKO", "$ WIADERKO NA RESZTKI"),
    artykul("WOZEK", "WÓZEK TRANS"),
    artykul("KORKOCIAG", "KORKOCIĄG TRYBUSZON / OTWIERACZ DO NAPOI"),
    artykul("COOLER", "COOLER"),
    artykul("NOZ_TORT", "NÓŻ DO TORTU"),
    artykul("LOPATKA_TORT", "ŁOPATKA DO TORTU"),
    artykul("TERMOBOX", "TERMOBOX"),
    artykul("LOD", "$ LÓD"),
    artykul("GRILL", "GRILL"),
    artykul("BUTLA_GAZ", "BUTLA GAZOWA"),
    artykul("TACA", "TACA KELNERSKA"),
    artykul("SERWETNIK", "SERWETNIK"),
    artykul("SERWETKI", "$ SERWETKI JEDNORAZOWE MAŁE"),
    artykul("STEND_MENU", "STEND DO MENU A5"),
]

# Do recznego slownika dodajemy wszystkie pozostale pozycje z materialu
# zrodlowego. Core ma czytelne kody wykorzystywane przez silnik, a reszta jest
# dostepna w przycisku "Dodaj ze slownika".
_nazwy_core = {re.sub(r"[^A-Z0-9]", "", x.nazwa.upper()) for x in _ARTYKULY}
for _dane in DANE_SLOWNIKA:
    _porownanie = re.sub(r"[^A-Z0-9]", "", _dane["nazwa"].upper())
    if _porownanie not in _nazwy_core and _dane["nazwa"].strip().upper() != "Q":
        _ARTYKULY.append(
            artykul(
                f"BAZA_{_dane['kod']}",
                _dane["nazwa"],
                _dane["jednostka"],
                int(_dane["opakowanie"] or 1),
                _dane["uwagi"],
            )
        )
SLOWNIK = {pozycja.kod: pozycja for pozycja in _ARTYKULY}


class DaneWydarzenia(BaseModel):
    nazwa_wydarzenia: str = ""
    data_imprezy: str = ""
    godzina_rozpoczecia: str = ""
    godzina_zakonczenia: str = ""
    liczba_osob: int = 0
    miejsce: str = ""
    adres: str = ""
    osoba_kontaktowa: str = ""
    telefon_kontaktowy: str = ""
    modul_przerwy: str = Field(default="brak")
    przerwa_calodniowa: bool = False
    grill: bool = False
    liczba_stref_grill: int = 0
    obiad: str = Field(default="brak")
    tort: str = Field(default="brak")
    liczba_dan_cieplych: int = 0
    liczba_salat: int = 0
    liczba_deserow: int = 0
    sa_napoje_zimne: bool = False
    sa_wino_lub_prosecco: bool = False
    liczba_stolikow_koktajlowych: int = 0


PROMPT = """
Przeanalizuj ofertę cateringową w PDF. Zwróć tylko poprawny JSON zgodny ze
schematem. Podaj konkretne dane z dokumentu. Jeżeli danej nie ma, zwróć pusty
tekst, 0 albo false. Rozpoznaj typy przerwy (użyj dokładnie jednej wartości):
zwykla, z_napojem, przystawka, z_deserem albo komplet; obiad: lunch albo lunch_z_zupa;
tort: tort, tort_z_winem. Policz osobne dania ciepłe, sałatki i desery.
"""


def przedzial(liczba_osob: int, punkty: list[tuple[int, int]]) -> int:
    """Zwraca pierwszą wartość dla górnej granicy; powyżej skaluje ostatnią."""
    for granica, wartosc in punkty:
        if liczba_osob <= granica:
            return wartosc
    granica, wartosc = punkty[-1]
    return math.ceil(liczba_osob / granica) * wartosc


def mnoznik_zapasowy(liczba_osob: int, maly: float = 1.3, sredni: float = 1.2, duzy: float = 1.1) -> int:
    mnoznik = maly if liczba_osob <= 40 else sredni if liczba_osob <= 200 else duzy
    return math.ceil(liczba_osob * mnoznik)


def do_paczki(ilosc: int | float, opakowanie: int) -> int:
    return math.ceil(float(ilosc) / max(1, opakowanie)) * max(1, opakowanie)


def nowy_wiersz(kod: str, ilosc: int, sekcja: str, uwagi: str = "") -> dict[str, Any]:
    pozycja = SLOWNIK[kod]
    uwaga = " | ".join(x for x in (pozycja.uwagi, uwagi) if x)
    return {
        "Sekcja": sekcja,
        "Artykuł": pozycja.nazwa,
        "Jednostka": pozycja.jednostka,
        "Liczba zamówiona": do_paczki(ilosc, pozycja.opakowanie),
        "Liczba spakowana": "",
        "Uwagi": uwaga,
        "Opakowanie": pozycja.opakowanie,
    }


def dodaj(wiersze: list[dict[str, Any]], kod: str, ilosc: int, sekcja: str, uwagi: str = "") -> None:
    if ilosc > 0:
        wiersze.append(nowy_wiersz(kod, ilosc, sekcja, uwagi))


def dodaj_stoly(wiersze: list[dict[str, Any]], liczba_stolow: int, liczba_koktajlow: int, sekcja: str) -> None:
    dodaj(wiersze, "STOL_BUFET", liczba_stolow, sekcja)
    dodaj(wiersze, "NACIAG_BUFET_CZARNY", liczba_stolow, sekcja)
    dodaj(wiersze, "CZAPKA_BUFET", liczba_stolow, sekcja)
    dodaj(wiersze, "KWIATY_BUFET", max(1, liczba_stolow), sekcja)
    if liczba_koktajlow:
        dodaj(wiersze, "STOL_KOKTAJL", liczba_koktajlow, sekcja)
        dodaj(wiersze, "NACIAG_KOKTAJL_CZARNY", liczba_koktajlow, sekcja)
        dodaj(wiersze, "CZAPKA_KOKTAJL", liczba_koktajlow, sekcja)
        dodaj(wiersze, "KWIATY_KOKTAJL", liczba_koktajlow, sekcja)


def dodaj_serwis_kawy(wiersze: list[dict[str, Any]], osoby: int, sekcja: str) -> None:
    ekspresy = przedzial(osoby, [(50, 1), (100, 2), (150, 3), (200, 4), (500, 10)])
    dodaj(wiersze, "EKSPRES", ekspresy, sekcja)
    dodaj(wiersze, "FILIŻANKA", mnoznik_zapasowy(osoby), sekcja)
    dodaj(wiersze, "SPODEK", mnoznik_zapasowy(osoby), sekcja)
    dodaj(wiersze, "LYZECZKA", osoby, sekcja)
    dodaj(wiersze, "KUBKI_PAPIEROWE", math.ceil(mnoznik_zapasowy(osoby) / 2), sekcja)
    dodaj(wiersze, "WARNIK_DUZY", przedzial(osoby, [(100, 1), (500, 5)]), sekcja)
    for kod, ilosc in [("KAWA", przedzial(osoby, [(50, 1), (500, 10)])), ("MLEKO", przedzial(osoby, [(25, 2), (500, 40)])), ("MLEKO_BEZ", przedzial(osoby, [(25, 1), (500, 20)])), ("MLEKO_ROSLINNE", przedzial(osoby, [(25, 1), (500, 20)])), ("KARAFKA_MLEKO", ekspresy), ("DZBANEK_MLEKO", ekspresy * 2), ("MISECZKA", ekspresy * 2), ("STEND_MLEKO", ekspresy * 2), ("DYSPENSER_HERBATA", przedzial(osoby, [(100, 1), (500, 5)])), ("WODA_5L", ekspresy), ("CUKIER_BIALY", ekspresy), ("CUKIER_BRAZOWY", ekspresy), ("CYTRYNA", ekspresy)]:
        dodaj(wiersze, kod, ilosc, sekcja)
    dodaj(wiersze, "PRZEDLUZACZ", ekspresy + przedzial(osoby, [(100, 1), (500, 5)]), sekcja)


def kalkuluj(dane: dict[str, Any]) -> pd.DataFrame:
    osoby = max(1, int(dane.get("liczba_osob", 0)))
    wiersze: list[dict[str, Any]] = []
    koktajle = int(dane.get("liczba_stolikow_koktajlowych", 0))

    przerwa = dane.get("modul_przerwy", "brak")
    if przerwa != "brak":
        profily = {
            "zwykla": [(10, 1), (40, 2), (100, 3), (200, 4), (300, 5), (500, 6)],
            "z_napojem": [(10, 2), (40, 3), (100, 3), (200, 4), (300, 5), (500, 7)],
            "przystawka": [(10, 1), (40, 1), (100, 2), (200, 3), (300, 4), (500, 5)],
            "z_deserem": [(10, 2), (40, 3), (100, 3), (200, 4), (300, 5), (500, 7)],
            "komplet": [(10, 3), (40, 3), (100, 4), (200, 5), (300, 6), (500, 8)],
        }
        sekcja = "PRZERWA KAWOWA"
        dodaj_stoly(wiersze, przedzial(osoby, profily[przerwa]), koktajle or math.ceil(osoby / 8), sekcja)
        dodaj_serwis_kawy(wiersze, osoby, sekcja)
        if przerwa in {"z_deserem", "komplet"}:
            dodaj(wiersze, "TALERZ_DESER", osoby, sekcja)
            dodaj(wiersze, "WIDELCZYK", osoby, sekcja)
            dodaj(wiersze, "SERWIS_SZCZYPCE", 2, sekcja)
        if przerwa in {"przystawka", "komplet"}:
            dodaj(wiersze, "TALERZ_PRZEKASKA", osoby, sekcja)
            dodaj(wiersze, "WIDELCZYK", osoby, sekcja)
        if przerwa in {"z_napojem", "komplet"}:
            dodaj(wiersze, "DYSPENSER_NAPOJE", przedzial(osoby, [(100, 3), (200, 6), (500, 9)]), sekcja, "Po jednym na rodzaj napoju")
            dodaj(wiersze, "LONG", osoby, sekcja)

    obiad = dane.get("obiad", "brak")
    if obiad != "brak":
        sekcja = "OBIAD"
        tablica = [(10, 2), (40, 3), (100, 5 if obiad == "lunch_z_zupa" else 4), (200, 7 if obiad == "lunch_z_zupa" else 6), (500, 10 if obiad == "lunch_z_zupa" else 9)]
        dodaj_stoly(wiersze, przedzial(osoby, tablica), koktajle or math.ceil(osoby / 5), sekcja)
        dodaj(wiersze, "TALERZ_OBIAD", mnoznik_zapasowy(osoby), sekcja)
        dodaj(wiersze, "WIDELEC", mnoznik_zapasowy(osoby), sekcja)
        dodaj(wiersze, "NOZ", mnoznik_zapasowy(osoby), sekcja)
        dania = max(1, int(dane.get("liczba_dan_cieplych", 1)))
        dodaj(wiersze, "PODGRZEWACZ_DANIE", dania, sekcja)
        dodaj(wiersze, "SERWIS_LYZKA", dania, sekcja)
        if obiad == "lunch_z_zupa":
            dodaj(wiersze, "BULIONOWKA", osoby, sekcja)
            dodaj(wiersze, "LYZKA", osoby, sekcja)
            dodaj(wiersze, "PODGRZEWACZ_ZUPA", 1, sekcja)
            dodaj(wiersze, "SERWIS_CHOCHLA", 1, sekcja)
        salatki = int(dane.get("liczba_salat", 0))
        dodaj(wiersze, "SERWIS_SZCZYPCE", salatki, sekcja)
        dodaj(wiersze, "PODSTAWKA_SERWIS", dania + (1 if obiad == "lunch_z_zupa" else 0), sekcja)
        dodaj(wiersze, "SERWETNIK", 1, sekcja, "DUŻY")

    if dane.get("grill"):
        sekcja = "GRILL"
        dodaj_stoly(wiersze, przedzial(osoby, [(10, 4), (40, 4), (100, 6), (200, 7), (300, 8), (500, 10)]), koktajle or math.ceil(osoby / 5), sekcja)
        dodaj(wiersze, "GRILL", max(1, int(dane.get("liczba_stref_grill", 1))), sekcja)
        dodaj(wiersze, "BUTLA_GAZ", max(1, int(dane.get("liczba_stref_grill", 1))), sekcja)
        dodaj(wiersze, "TALERZ_OBIAD", mnoznik_zapasowy(osoby), sekcja)
        dodaj(wiersze, "WIDELEC", mnoznik_zapasowy(osoby), sekcja)
        dodaj(wiersze, "NOZ", mnoznik_zapasowy(osoby), sekcja)

    tort = dane.get("tort", "brak")
    if tort != "brak":
        sekcja = "TORT" if tort == "tort" else "TORT + WINO"
        dodaj_stoly(wiersze, przedzial(osoby, [(300, 1 if tort == "tort" else 2), (600, 2 if tort == "tort" else 3)]), koktajle or math.ceil(osoby / 8), sekcja)
        dodaj(wiersze, "TALERZ_DESER", mnoznik_zapasowy(osoby), sekcja)
        dodaj(wiersze, "WIDELCZYK", mnoznik_zapasowy(osoby, maly=2.0), sekcja)
        dodaj(wiersze, "NOZ_TORT", 1, sekcja)
        dodaj(wiersze, "LOPATKA_TORT", 1, sekcja)
        if tort == "tort_z_winem":
            dodaj(wiersze, "COOLER", 3, sekcja)
            dodaj(wiersze, "KIELISZEK_WINO", osoby, sekcja)
            dodaj(wiersze, "KORKOCIAG", 2, sekcja)
            dodaj(wiersze, "TERMOBOX", przedzial(osoby, [(100, 1), (500, 5)]), sekcja)
            dodaj(wiersze, "LOD", 1, sekcja)

    if dane.get("sa_wino_lub_prosecco") and tort != "tort_z_winem":
        dodaj(wiersze, "KORKOCIAG", 1, "DODATKI")
    if wiersze:
        dodaj(wiersze, "SEPARATOR", 1, "DODATKI")
        dodaj(wiersze, "WOZEK", 1, "DODATKI")
        for kod in ("PSIK", "REKAWICZKI", "CZYSCIWO", "WORKI", "WIADERKO"):
            dodaj(wiersze, kod, 1, "DODATKI")

    # Laczymy identyczne pozycje, aby jedna pozycja nie wystepowala kilka razy.
    wynik: dict[tuple[str, str], dict[str, Any]] = {}
    for wiersz in wiersze:
        klucz = (wiersz["Sekcja"], wiersz["Artykuł"])
        if klucz in wynik:
            wynik[klucz]["Liczba zamówiona"] += wiersz["Liczba zamówiona"]
        else:
            wynik[klucz] = wiersz
    return pd.DataFrame(wynik.values(), columns=["Sekcja", "Artykuł", "Jednostka", "Liczba zamówiona", "Liczba spakowana", "Uwagi", "Opakowanie"])


def fonty_pdf() -> tuple[str, str]:
    zwykly, pogrubiony = "Helvetica", "Helvetica-Bold"
    sciezki = [(Path("C:/Windows/Fonts/arial.ttf"), Path("C:/Windows/Fonts/arialbd.ttf")), (Path("C:/Windows/Fonts/DejaVuSans.ttf"), Path("C:/Windows/Fonts/DejaVuSans-Bold.ttf"))]
    for normalny, bold in sciezki:
        if normalny.exists() and bold.exists():
            pdfmetrics.registerFont(TTFont("Catering", str(normalny)))
            pdfmetrics.registerFont(TTFont("CateringBold", str(bold)))
            return "Catering", "CateringBold"
    return zwykly, pogrubiony


def tekst(wartosc: Any) -> str:
    return str(wartosc if wartosc not in (None, "nan") else "")


def generuj_pdf(dane: dict[str, Any], tabela: pd.DataFrame) -> bytes:
    normalny, bold = fonty_pdf()
    bufor = io.BytesIO()
    doc = SimpleDocTemplate(bufor, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    style = getSampleStyleSheet()
    tytul = ParagraphStyle("Tytul", parent=style["Normal"], fontName=bold, fontSize=15, leading=18, alignment=TA_CENTER, spaceAfter=9)
    szczegoly = ParagraphStyle("Szczegoly", parent=style["Normal"], fontName=bold, fontSize=10, leading=13)
    komorka = ParagraphStyle("Komorka", parent=style["Normal"], fontName=normalny, fontSize=8, leading=9)
    komorka_bold = ParagraphStyle("KomorkaBold", parent=komorka, fontName=bold)
    elementy = [Paragraph("CHECKLISTA SPRZĘTU CATERINGOWEGO", tytul)]
    info = [
        ["WYDARZENIE", tekst(dane.get("nazwa_wydarzenia")), "DATA", tekst(dane.get("data_imprezy"))],
        ["GODZINA", f"{tekst(dane.get('godzina_rozpoczecia'))} – {tekst(dane.get('godzina_zakonczenia'))}", "LICZBA GOŚCI", f"{tekst(dane.get('liczba_osob'))} os."],
        ["MIEJSCE / ADRES", f"{tekst(dane.get('miejsce'))} {tekst(dane.get('adres'))}", "", ""],
        ["OSOBA KONTAKTOWA", tekst(dane.get("osoba_kontaktowa")), "TELEFON", tekst(dane.get("telefon_kontaktowy"))],
    ]
    info_par = [[Paragraph(f"<b>{a}</b>", szczegoly), Paragraph(b or "&nbsp;", szczegoly), Paragraph(f"<b>{c}</b>", szczegoly), Paragraph(d or "&nbsp;", szczegoly)] for a, b, c, d in info]
    tabela_info = Table(info_par, colWidths=[104, 216, 104, 99])
    tabela_info.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#334155")), ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#e2e8f0")), ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#e2e8f0")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("SPAN", (1, 2), (3, 2))]))
    elementy += [tabela_info, Spacer(1, 9)]
    pola = Table([[Paragraph("<b>SAMOCHÓD:</b> ____________________________________", szczegoly), Paragraph("<b>KTO ZABIERA:</b> ____________________________________", szczegoly)]], colWidths=[261.5, 261.5])
    pola.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#334155")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8), ("LEFTPADDING", (0, 0), (-1, -1), 7)]))
    elementy += [pola, Spacer(1, 12)]
    dane_tabeli = [["L.p.", "Artykuł", "Liczba\nzamówiona", "Liczba\nspakowana", "Uwagi"]]
    poprzednia_sekcja = None
    lp = 0
    for _, r in tabela.iterrows():
        if r["Sekcja"] != poprzednia_sekcja:
            dane_tabeli.append([Paragraph(f"<b>{tekst(r['Sekcja'])}</b>", komorka_bold), "", "", "", ""])
            poprzednia_sekcja = r["Sekcja"]
        lp += 1
        dane_tabeli.append([str(lp), Paragraph(tekst(r["Artykuł"]), komorka), Paragraph(tekst(r["Liczba zamówiona"]), komorka), Paragraph("", komorka), Paragraph(tekst(r["Uwagi"]), komorka)])
    tabela_glowna = Table(dane_tabeli, colWidths=[30, 212, 80, 80, 121], repeatRows=1)
    style_tabeli = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f172a")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONTNAME", (0, 0), (-1, 0), bold), ("FONTSIZE", (0, 0), (-1, 0), 8), ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (3, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
    for i, wiersz in enumerate(dane_tabeli[1:], start=1):
        if isinstance(wiersz[0], Paragraph):
            style_tabeli.extend([("SPAN", (0, i), (-1, i)), ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#dbeafe")), ("TOPPADDING", (0, i), (-1, i), 4), ("BOTTOMPADDING", (0, i), (-1, i), 4)])
    tabela_glowna.setStyle(TableStyle(style_tabeli))
    elementy.append(tabela_glowna)
    doc.build(elementy)
    return bufor.getvalue()


def odczytaj_pdf(uploaded_file: Any) -> dict[str, Any]:
    if API_KEY.startswith("WKLEJ_"):
        raise ValueError("Wstaw swój klucz Gemini w zmiennej API_KEY na początku pliku app.py.")
    client = genai.Client(api_key=API_KEY)
    response = client.models.generate_content(
        model=MODEL_NAME,
        contents=[types.Part.from_bytes(data=uploaded_file.getvalue(), mime_type="application/pdf"), PROMPT],
        config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=DaneWydarzenia),
    )
    wynik = response.parsed
    if isinstance(wynik, DaneWydarzenia):
        return wynik.model_dump()
    if wynik:
        return DaneWydarzenia.model_validate(wynik).model_dump()
    return DaneWydarzenia.model_validate_json(response.text).model_dump()


def formularz_danych(dane: dict[str, Any]) -> dict[str, Any]:
    st.subheader("1. Dane wydarzenia i rodzaje serwisu")
    with st.form("dane_wydarzenia"):
        a, b, c = st.columns(3)
        dane["nazwa_wydarzenia"] = a.text_input("Wydarzenie", dane.get("nazwa_wydarzenia", ""))
        dane["data_imprezy"] = b.text_input("Data", dane.get("data_imprezy", ""))
        dane["liczba_osob"] = c.number_input("Liczba gości", min_value=1, value=max(1, int(dane.get("liczba_osob", 1))))
        a, b, c = st.columns(3)
        dane["godzina_rozpoczecia"] = a.text_input("Godzina rozpoczęcia", dane.get("godzina_rozpoczecia", ""))
        dane["godzina_zakonczenia"] = b.text_input("Godzina zakończenia", dane.get("godzina_zakonczenia", ""))
        dane["miejsce"] = c.text_input("Miejsce", dane.get("miejsce", ""))
        a, b, c = st.columns(3)
        dane["adres"] = a.text_input("Adres", dane.get("adres", ""))
        dane["osoba_kontaktowa"] = b.text_input("Osoba kontaktowa", dane.get("osoba_kontaktowa", ""))
        dane["telefon_kontaktowy"] = c.text_input("Telefon", dane.get("telefon_kontaktowy", ""))
        a, b, c, d = st.columns(4)
        opcje_przerwy = ["brak", "zwykla", "z_napojem", "przystawka", "z_deserem", "komplet"]
        opcje_obiadu = ["brak", "lunch", "lunch_z_zupa"]
        opcje_tortu = ["brak", "tort", "tort_z_winem"]
        dane["modul_przerwy"] = a.selectbox("Przerwa kawowa", opcje_przerwy, index=opcje_przerwy.index(dane["modul_przerwy"]) if dane.get("modul_przerwy") in opcje_przerwy else 0)
        dane["obiad"] = b.selectbox("Obiad", opcje_obiadu, index=opcje_obiadu.index(dane["obiad"]) if dane.get("obiad") in opcje_obiadu else 0)
        dane["tort"] = c.selectbox("Tort", opcje_tortu, index=opcje_tortu.index(dane["tort"]) if dane.get("tort") in opcje_tortu else 0)
        dane["grill"] = d.checkbox("Grill", value=bool(dane.get("grill", False)))
        a, b, c, d = st.columns(4)
        dane["liczba_dan_cieplych"] = a.number_input("Dania ciepłe", min_value=1, value=max(1, int(dane.get("liczba_dan_cieplych", 1))))
        dane["liczba_salat"] = b.number_input("Sałatki", min_value=0, value=int(dane.get("liczba_salat", 0)))
        dane["liczba_deserow"] = c.number_input("Rodzaje deserów", min_value=0, value=int(dane.get("liczba_deserow", 0)))
        dane["liczba_stolikow_koktajlowych"] = d.number_input("Stoliki koktajlowe (0 = automatycznie)", min_value=0, value=int(dane.get("liczba_stolikow_koktajlowych", 0)))
        dane["liczba_stref_grill"] = a.number_input("Strefy grillowe", min_value=0, max_value=99, value=int(dane.get("liczba_stref_grill", 1)))
        dane["sa_wino_lub_prosecco"] = b.checkbox("Wino / prosecco", value=bool(dane.get("sa_wino_lub_prosecco", False)))
        zatwierdz = st.form_submit_button("Wylicz checklistę", type="primary")
    if zatwierdz:
        st.session_state["dane"] = dane
        st.session_state["checklista"] = kalkuluj(dane)
        st.success("Checklista została przeliczona. Przejdź do jej edycji.")
    return dane


def main() -> None:
    st.title("📦 Generator checklist sprzętu cateringowego")
    st.caption("PDF oferty → analiza Gemini → reguły zapisane w Pythonie → edytowalna checklista → PDF A4")
    if "dane" not in st.session_state:
        st.session_state["dane"] = DaneWydarzenia().model_dump()

    uploaded = st.file_uploader("Wgraj ofertę / agendę PDF", type="pdf")
    if uploaded and st.button("Przeczytaj PDF przez Gemini", type="primary"):
        with st.spinner("Gemini analizuje dokument..."):
            try:
                st.session_state["dane"] = odczytaj_pdf(uploaded)
                st.success("Dane zostały odczytane. Sprawdź je i popraw poniżej.")
            except Exception as exc:
                st.error(f"Nie udało się odczytać PDF: {exc}")

    formularz_danych(dict(st.session_state["dane"]))
    if "checklista" not in st.session_state:
        return

    st.subheader("2. Edycja checklisty przed drukiem")
    col_ed, col_add = st.columns([3, 1])
    with col_ed:
        edytowana = st.data_editor(st.session_state["checklista"], num_rows="dynamic", use_container_width=True, hide_index=True, key="edytor", column_config={"Liczba zamówiona": st.column_config.NumberColumn(min_value=0, step=1), "Liczba spakowana": st.column_config.TextColumn(help="Pole do wpisania ręcznie po wydruku lub w aplikacji"), "Opakowanie": None})
        st.session_state["checklista"] = edytowana
    with col_add:
        st.markdown("#### Dodaj ze słownika")
        opcje = {f"{x.nazwa} ({x.jednostka}) [{x.kod}]": x.kod for x in _ARTYKULY}
        nazwa = st.selectbox("Artykuł", list(opcje), key="wybierz_slownik")
        ilosc = st.number_input("Ilość", min_value=1, value=1, key="dodaj_ilosc")
        uwagi = st.text_input("Uwagi", key="dodaj_uwagi")
        if st.button("Dodaj artykuł"):
            nowy = nowy_wiersz(opcje[nazwa], int(ilosc), "DODATKOWE", uwagi)
            st.session_state["checklista"] = pd.concat([st.session_state["checklista"], pd.DataFrame([nowy])], ignore_index=True)
            st.rerun()
        st.markdown("#### Pozycja własna")
        wlasna = st.text_input("Nazwa własna", key="nazwa_wlasna")
        if st.button("Dodaj pozycję własną") and wlasna.strip():
            nowy = {"Sekcja": "DODATKOWE", "Artykuł": wlasna.strip(), "Jednostka": "SZT", "Liczba zamówiona": 1, "Liczba spakowana": "", "Uwagi": "", "Opakowanie": 1}
            st.session_state["checklista"] = pd.concat([st.session_state["checklista"], pd.DataFrame([nowy])], ignore_index=True)
            st.rerun()

    st.subheader("3. PDF do druku")
    pdf = generuj_pdf(st.session_state["dane"], st.session_state["checklista"])
    st.download_button("Pobierz checklistę PDF A4", data=pdf, file_name="checklista_magazynowa.pdf", mime="application/pdf", type="primary")


if __name__ == "__main__":
    main()
