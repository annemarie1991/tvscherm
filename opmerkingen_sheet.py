"""Pony-opmerkingen uit de Google Sheet "TV-scherm opmerkingen".

Vroeger stonden de opmerkingen in pony_opmerkingen.json op de server van
Streamlit. Bij elke herstart van de app was dat bestand weg. Nu staan ze in de
sheet, en die blijft gewoon bestaan.

Indeling van de sheet (eerste tabblad, rij 1 is de kop):
- kolom A: de pony (of een deel van de naam)
- kolom B: de opmerking zoals die op het scherm komt (kort)
- kolom C: uitgebreide opmerking voor Annemarie zelf; die gebruikt de app niet
Rijen zonder pony of zonder opmerking in B worden overgeslagen.
"""

import csv
import io
import urllib.error
import urllib.request

SHEET_ID = "1eadWoSPPuGJLEioQGNKHONPDVb_gZxYTqfzKazYTIpQ"
SHEET_LINK = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
CSV_LINK = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"


def lees_rijen():
    """Alle rijen van het eerste tabblad, zonder de kop."""
    try:
        with urllib.request.urlopen(CSV_LINK, timeout=20) as antwoord:
            tekst = antwoord.read().decode("utf-8")
    except urllib.error.HTTPError as fout:
        if fout.code in (401, 403):
            raise RuntimeError("de sheet is niet te lezen: zet delen op 'iedereen met de link "
                               "kan bekijken'") from None
        raise
    if tekst.lstrip().lower().startswith("<!doctype html"):
        raise RuntimeError("de sheet is niet te lezen (staat hij op 'iedereen met de link'?)")
    return list(csv.reader(io.StringIO(tekst)))[1:]


def maak_opmerkingen(rijen):
    """Van de rijen in de sheet naar {pony: opmerking}, alleen kolom A en B."""
    opmerkingen = {}
    for rij in rijen:
        pony = rij[0].strip() if len(rij) > 0 else ""
        opmerking = rij[1].strip() if len(rij) > 1 else ""
        if pony and opmerking:
            opmerkingen[pony] = opmerking
    return opmerkingen


def lees_opmerkingen():
    return maak_opmerkingen(lees_rijen())
