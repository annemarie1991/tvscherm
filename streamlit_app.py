import streamlit as st
import pandas as pd
import datetime
import re
import locale
from pathlib import Path
import json
import google.auth
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Nederlandse datuminstelling
try:
    locale.setlocale(locale.LC_TIME, 'nl_NL.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'nl_NL')
    except:
        pass

# App configuratie
st.set_page_config(page_title="Het Zesspan Ponyplanner", layout="wide")
st.title("🐴 Het Zesspan Ponyplanner")

st.markdown("""
Upload hieronder het Excel-bestand met de planning. Kies daarna het juiste tabblad.
De app herkent automatisch de ponynamen, lestijden, kindernamen (geanonimiseerd) en juffen.
""")

tekstpad = Path("ondertekst.txt")
if "ondertekst" not in st.session_state:
    if tekstpad.exists():
        regels = tekstpad.read_text(encoding="utf-8").split("\n")
        st.session_state.ondertekst = regels[0] if regels else ""
        st.session_state.vet = regels[1] == "True" if len(regels) > 1 else False
        st.session_state.geel = regels[2] == "True" if len(regels) > 2 else False
    else:
        st.session_state.ondertekst = ""
        st.session_state.vet = False
        st.session_state.geel = False

st.sidebar.header("📝 Ondertekst instellen")
nieuwe_tekst = st.sidebar.text_area("Tekst onderaan elke sectie", st.session_state.ondertekst)
vet = st.sidebar.checkbox("Dikgedrukt", value=st.session_state.vet)
geel = st.sidebar.checkbox("Geel markeren", value=st.session_state.geel)
if st.sidebar.button("💾 Opslaan"):
    st.session_state.ondertekst = nieuwe_tekst
    st.session_state.vet = vet
    st.session_state.geel = geel
    tekstpad.write_text(f"{nieuwe_tekst}\n{vet}\n{geel}", encoding="utf-8")
    st.sidebar.success("Tekst opgeslagen!")

uploaded_file = st.file_uploader("📄 Upload je Excel-bestand", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📃 Kies een tabblad", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet, header=None)
    st.dataframe(df.head(20))

    ponynamen_kolom, ponynamen_start_index = None, 0
    for col in df.columns:
        telling, start_index = 0, None
        for idx, value in df[col].dropna().astype(str).items():
            if re.match(r"^[A-Za-z\- ]+$", value):
                if telling == 0:
                    start_index = idx
                telling += 1
                if telling >= 60:
                    ponynamen_kolom = col
                    ponynamen_start_index = start_index
                    break
            else:
                telling = 0
        if ponynamen_kolom is not None:
            break

    if ponynamen_kolom is not None:
        tijdrij, tijd_pattern = None, re.compile(r"\b\d{1,2}:\d{2}(\s*[-–−]\s*\d{1,2}:\d{2})?\b")
        for i in range(0, 5):
            if any(tijd_pattern.search(str(cell)) for cell in df.iloc[i]):
                tijdrij = i
                break

        if tijdrij is not None:
            tijd_dict = {col: str(df.iloc[tijdrij, col]).strip() for col in df.columns if tijd_pattern.match(str(df.iloc[tijdrij, col]))}
            tijd_items = sorted(tijd_dict.items(), key=lambda x: datetime.datetime.strptime(re.search(r"\d{1,2}:\d{2}", x[1]).group(), "%H:%M"))

            groepen_per_blok, huidige_blok, laatst_verwerkte_tijd = [], [], None
            for col, tijd in tijd_items:
                tijd_clean = re.search(r"\d{1,2}:\d{2}", tijd).group()
                tijd_dt = datetime.datetime.strptime(tijd_clean, "%H:%M")
                if laatst_verwerkte_tijd is None or (tijd_dt - laatst_verwerkte_tijd).total_seconds() > 30 * 60:
                    if huidige_blok:
                        groepen_per_blok.append(huidige_blok)
                    huidige_blok = [(col, tijd)]
                    laatst_verwerkte_tijd = tijd_dt
                else:
                    huidige_blok.append((col, tijd))
            if huidige_blok:
                groepen_per_blok.append(huidige_blok)

            datum_vandaag = datetime.datetime.today().strftime("%d-%m-%Y")
            eigen_pony_rij = next((r for r in range(ponynamen_start_index, len(df)) if "eigen pony" in str(df.iloc[r, ponynamen_kolom]).strip().lower()), None)
            reeds_in_bak, output_for_slides = set(), []

            for blok in groepen_per_blok:
                groep_markdown = [f"<div style='text-align:center; font-weight:bold; font-size:18px'>{datum_vandaag}</div>", "<br><div style='display:flex; gap:20px;'>"]
                for col, tijd in blok:
                    juf = str(df.iloc[eigen_pony_rij + 2, col]).strip().title() if eigen_pony_rij is not None and eigen_pony_rij + 2 < len(df) else "onbekend"
                    kind_pony_combinaties, namen_counter = [], {}
                    max_rij = eigen_pony_rij if eigen_pony_rij else len(df)
                    for r in range(ponynamen_start_index, max_rij):
                        naam = str(df.iloc[r, col])
                        pony = str(df.iloc[r, ponynamen_kolom])
                        if naam.lower() in ["", "nan", "x"]:
                            continue
                        delen = naam.strip().split()
                        voornaam = delen[0].capitalize() if delen else ""
                        achternaam = next((d.capitalize() for d in delen[1:] if d.lower() not in {"van", "de", "der", "den", "ter", "ten", "het", "te"}), "")
                        code = voornaam + (achternaam[:1].upper() if voornaam.lower() in namen_counter else "")
                        namen_counter[voornaam.lower()] = namen_counter.get(voornaam.lower(), 0) + 1
                        locatie = "(B)" if pony in reeds_in_bak else "(S)"
                        kind_pony_combinaties.append((code, f"{pony.title()} {locatie}"))
                        reeds_in_bak.add(pony)
                    kind_pony_combinaties.sort(key=lambda x: x[0].lower())
                    groep_markdown.append("<div><strong>Groep {}<br>Juf: {}</strong><br>{}</div>".format(
                        tijd,
                        juf,
                        "<br>".join(f"{n} – {p}" for n, p in kind_pony_combinaties)
                    ))
                groep_markdown.append("</div>")
                if st.session_state.ondertekst:
                    stijl = ""
                    if st.session_state.geel:
                        stijl += "background-color:yellow; padding:4px; border-radius:4px;"
                    if st.session_state.vet:
                        stijl += "font-weight:bold;"
                    groep_markdown.append(f"<div style='text-align:center; margin-top:1em; {stijl}'>{st.session_state.ondertekst}</div>")

                sectie_html = "\n".join(groep_markdown)
                st.markdown("<hr>" + sectie_html, unsafe_allow_html=True)
                output_for_slides.append(sectie_html)

            st.session_state["slides_output"] = output_for_slides

            # Upload knop naar Google Slides
            if st.button("📤 Upload naar (online) scherm"):
                from slides_uploader import upload_to_slides
                presentation_id = "1uhfZtV-ota0vNrmQyLi6hmxRMvkuo1qM"
                upload_to_slides(presentation_id, st.session_state["slides_output"])
                st.success("De planning is geüpload naar het online scherm!")
