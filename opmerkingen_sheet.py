"""Pony-opmerkingen uit de Google Sheet "TV-scherm opmerkingen".

Vroeger stonden de opmerkingen in pony_opmerkingen.json op de server van
Streamlit. Bij elke herstart van de app was dat bestand weg. Nu staan ze in de
sheet, en die blijft gewoon bestaan.

Indeling van de sheet (eerste tabblad, rij 1 is de kop):
- kolom A: de pony (of een deel van de naam)
- kolom B: de opmerking zoals die op het scherm komt (kort)
- kolom C: uitgebreide opmerking voor Annemarie zelf; die gebruikt de app niet
- kolom G-I: het rijtje "Als we gaan springen", net zo opgebouwd
- het vakje "Springen in bak" (kop in K1, de bakken in L1): in welke bakken er
  deze week gesprongen wordt, bijvoorbeeld "1+2" of "alle". Leeg = niet
  springen.
Rijen zonder pony of zonder opmerking in B worden overgeslagen.

Springen (sinds 09-11-2026): in een les in een van de springbakken geldt het
rijtje G-H, in de andere lessen gewoon A-B. Welke bak een les heeft, staat in
de rij "Bak" van de daglijst. Weet de app de bak niet (geen rij "Bak", of het
vakje is leeg), dan geldt het springrijtje: liever een keer geen elastiek dan
elastiek tijdens het springen. Een bak met "buiten" erin springt niet.
In een springles (Niveau "Spring ...") geldt het springrijtje altijd, ook in
een gewone week.
"""

import csv
import io
import re
import urllib.error
import urllib.request

SHEET_ID = "1eadWoSPPuGJLEioQGNKHONPDVb_gZxYTqfzKazYTIpQ"
SHEET_LINK = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
CSV_LINK = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

SPRINGVAKJE = "springen in bak"
ALLE = "alle"


def lees_alle_rijen():
    """Alle rijen van het eerste tabblad, met de kop."""
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
    return list(csv.reader(io.StringIO(tekst)))


def lees_rijen():
    """Alle rijen van het eerste tabblad, zonder de kop."""
    return lees_alle_rijen()[1:]


def maak_opmerkingen(rijen, pony_kolom=0, opmerking_kolom=1, ook_leeg=False):
    """Van de rijen in de sheet naar {pony: opmerking}, standaard kolom A en B."""
    opmerkingen = {}
    for rij in rijen:
        pony = rij[pony_kolom].strip() if len(rij) > pony_kolom else ""
        opmerking = rij[opmerking_kolom].strip() if len(rij) > opmerking_kolom else ""
        if pony and (opmerking or ook_leeg):
            opmerkingen[pony] = opmerking
    return opmerkingen


def lees_springvakje(alle_rijen):
    """Wat er naast (of onder) het vakje "Springen in bak" staat, of "" als het leeg is."""
    for r, rij in enumerate(alle_rijen):
        for k, cel in enumerate(rij):
            if cel.strip().lower().rstrip(":") == SPRINGVAKJE:
                rechts = rij[k + 1].strip() if len(rij) > k + 1 else ""
                if rechts:
                    return rechts
                onder = alle_rijen[r + 1] if len(alle_rijen) > r + 1 else []
                return onder[k].strip() if len(onder) > k else ""
    return ""


def maak_springbakken(tekst):
    """"1+2" -> {1, 2}; "alle" -> ALLE; leeg (of niet te lezen) -> None."""
    tekst = (tekst or "").strip().lower()
    if not tekst:
        return None
    if tekst.startswith("alle"):
        return ALLE
    nummers = {int(n) for n in re.findall(r"\d+", tekst)}
    return nummers or None


def bak_nummers(bak_cel):
    """Wat er in de rij "Bak" van de daglijst staat, als nummers.

    None = onbekend (leeg). Een lege set = wel iets, maar geen springbak
    ("buiten", "vlekkenbak", "stapritje").
    """
    if bak_cel is None:
        return None
    if isinstance(bak_cel, float):
        if bak_cel != bak_cel:  # NaN: leeg vakje
            return None
        if bak_cel.is_integer():
            return {int(bak_cel)}
    tekst = str(bak_cel).strip().lower()
    if not tekst or tekst == "nan":
        return None
    if "buiten" in tekst:
        return set()
    return {int(n) for n in re.findall(r"\d+", tekst)}


def niveau_en_bak_rij(labels, vanaf=0):
    """Rijnummers van "Niveau" en "Bak" onder het rooster (None als ze er niet zijn).

    labels: wat er per rij in kolom D staat. Op woensdag, donderdag en vrijdag
    staan de woorden "Niveau" en "Bak" er niet bij; dan zijn het de twee rijen
    meteen onder "Pauze" (gezien in periode 2 van 2026).
    """
    def schoon(i):
        if i >= len(labels):
            return None
        tekst = str(labels[i]).strip().lower()
        return "" if tekst == "nan" else tekst

    niveau = bak = pauze = None
    for i in range(vanaf, len(labels)):
        tekst = schoon(i)
        if tekst == "niveau" and niveau is None:
            niveau = i
        elif tekst == "bak" and bak is None:
            bak = i
        elif tekst == "pauze" and pauze is None:
            pauze = i
    if niveau is None and pauze is not None and schoon(pauze + 1) == "":
        niveau = pauze + 1
    if bak is None and niveau is not None and schoon(niveau + 1) == "":
        bak = niveau + 1
    return niveau, bak


def is_springles(niveau_cel):
    """Staat er "spring" bij Niveau ("Spring 1", "Springles 2")?"""
    return "spring" in str(niveau_cel or "").lower()


def springt_hier(bak_cel, springbakken, niveau_cel=None):
    """Geldt in deze les het springrijtje?

    Een springles altijd, elke week (Annemarie, 09-11-2026). Verder alleen in
    een week met springen, in de bakken uit het vakje "Springen in bak".
    """
    if is_springles(niveau_cel):
        return True
    if not springbakken:
        return False
    if springbakken == ALLE:
        return True
    if "buiten" in str(niveau_cel or "").lower():
        return False  # "P buitenbak" bij Niveau, ook als de bak leeg is
    nummers = bak_nummers(bak_cel)
    if nummers is None:
        return True  # onbekend: liever geen elastiek tijdens het springen
    return bool(nummers & springbakken)


def zoek_opmerking(pony, opmerkingen):
    """De opmerking voor deze pony, of "" (de eerste naam die in de ponynaam zit)."""
    for sleutel, tekst in opmerkingen.items():
        if sleutel.lower() in pony.lower():
            return tekst
    return ""


def lees_opmerkingen():
    return maak_opmerkingen(lees_rijen())


def lees_alles():
    """Alles wat de app uit de sheet nodig heeft, in een keer."""
    alle_rijen = lees_alle_rijen()
    gewoon = maak_opmerkingen(alle_rijen[1:])
    # Het springrijtje vult aan: een pony die er niet in staat, houdt zijn
    # gewone opmerking. Staat hij er wel in met een leeg vakje, dan geen.
    springen = dict(gewoon)
    springen.update(maak_opmerkingen(alle_rijen[1:], 6, 7, ook_leeg=True))
    vakje = lees_springvakje(alle_rijen)
    return {
        "gewoon": gewoon,
        "springen": springen,
        "springbakken": maak_springbakken(vakje),
        "springvakje": vakje,
    }
