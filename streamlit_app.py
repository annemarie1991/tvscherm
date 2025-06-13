import streamlit as st
import pandas as pd
import datetime
import re
import locale
import os
import json

# Pad voor notitie-opslag
notes_path = "/mnt/data/ponyplanner_notitie.json"
notitie_standaard = {"text": "", "bold": False, "highlight": False}

# Notitie laden
if os.path.exists(notes_path):
    with open(notes_path, "r") as f:
        opgeslagen_notitie = json.load(f)
else:
    opgeslagen_notitie = notitie_standaard

# Nederlandse datum
try:
    locale.setlocale(locale.LC_TIME, 'nl_NL.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'nl_NL')
    except:
        pass

st.set_page_config(page_title="Het Zesspan Ponyplanner", layout="wide")
st.title("🐴 Het Zesspan Ponyplanner")

st.markdown("""
Upload hieronder het Excel-bestand met de planning. Kies daarna het juiste tabblad.
De app herkent automatisch de ponynamen, lestijden, kindernamen (geanonimiseerd) en juffen.
""")

# 🟨 Notitiebeheer (boven de uploader)
st.markdown("---")
st.markdown("### 📌 Instellingen voor ondertekst")

with st.form("notitie_form"):
    tekst = st.text_input("Notitie voor onderaan iedere sectie:", value=opgeslagen_notitie.get("text", ""))
    col1, col2 = st.columns(2)
    bold = col1.checkbox("Dikgedrukt", value=opgeslagen_notitie.get("bold", False))
    highlight = col2.checkbox("Geel gemarkeerd", value=opgeslagen_notitie.get("highlight", False))
    opgeslagen = st.form_submit_button("🔖 Opslaan")

if opgeslagen:
    opgeslagen_notitie = {"text": tekst, "bold": bold, "highlight": highlight}
    with open(notes_path, "w") as f:
        json.dump(opgeslagen_notitie, f)
    st.success("Opgeslagen! Herlaad de pagina om te testen.")

st.markdown("---")

uploaded_file = st.file_uploader("📄 Upload je Excel-bestand", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📃 Kies een tabblad", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet, header=None)
    st.dataframe(df.head(20))

    ponynamen_kolom = None
    ponynamen_start_index = 0
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
        tijdrij = None
        tijd_pattern = re.compile(r"\b\d{1,2}:\d{2}(\s*[-\u2013\u2212]\s*\d{1,2}:\d{2})?\b")
        for i in range(0, 5):
            if any(tijd_pattern.search(str(cell)) for cell in df.iloc[i]):
                tijdrij = i
                break

        if tijdrij is not None:
            tijd_dict = {}
            for col in df.columns:
                cel = str(df.iloc[tijdrij, col])
                if tijd_pattern.match(cel):
                    tijd_dict[col] = cel.strip()

            tijd_items = sorted(
                tijd_dict.items(),
                key=lambda x: datetime.datetime.strptime(
                    re.search(r"\d{1,2}:\d{2}", x[1]).group(), "%H:%M"
                )
            )

            groepen_per_blok = []
            huidige_blok = []
            laatst_verwerkte_tijd = None

            for col, tijd in tijd_items:
                tijd_clean_match = re.search(r"\d{1,2}:\d{2}", tijd)
                if not tijd_clean_match:
                    continue
                tijd_clean = tijd_clean_match.group()
                try:
                    tijd_dt = datetime.datetime.strptime(tijd_clean, "%H:%M")
                except ValueError:
                    continue

                if laatst_verwerkte_tijd is None or (tijd_dt - laatst_verwerkte_tijd).total_seconds() > 30 * 60:
                    if huidige_blok:
                        groepen_per_blok.append(huidige_blok)
                    huidige_blok = [(col, tijd)]
                    laatst_verwerkte_tijd = tijd_dt
                else:
                    huidige_blok.append((col, tijd))

            if huidige_blok:
                groepen_per_blok.append(huidige_blok)

            st.markdown("### 🗓️ Planning per groep")

            # Alleen datum (geen weekdag)
            datum_vandaag = datetime.datetime.today().strftime("%d-%m-%Y")

            eigen_pony_rij = None
            for r in range(ponynamen_start_index, len(df)):
                waarde = str(df.iloc[r, ponynamen_kolom]).strip().lower()
                if "eigen pony" in waarde:
                    eigen_pony_rij = r
                    break

            reeds_in_bak = set()

            for blok in groepen_per_blok:
                st.markdown("---")
                st.markdown(f"<div style='text-align: center; font-size: 20px;'><strong>{datum_vandaag}</strong></div>", unsafe_allow_html=True)
                cols = st.columns(len(blok))

                for (col, tijd), container in zip(blok, cols):
                    max_rij = eigen_pony_rij if eigen_pony_rij else len(df)
                    juf = "onbekend"
                    if eigen_pony_rij is not None and eigen_pony_rij + 2 < len(df):
                        jufcel = df.iloc[eigen_pony_rij + 2, col]
                        juf = str(jufcel).strip().title() if pd.notna(jufcel) else "onbekend"

                    kind_pony_combinaties = []
                    namen_counter = {}

                    for r in range(ponynamen_start_index, max_rij):
                        naam = str(df.iloc[r, col]) if r < len(df) else ""
                        pony = str(df.iloc[r, ponynamen_kolom]) if r < len(df) else ""
                        if not naam or naam.strip().lower() in ["", "nan", "x"]:
                            continue

                        delen = naam.strip().split()
                        voornaam = delen[0].capitalize() if delen else ""
                        achternaam = ""
                        tussenvoegsels = {"van", "de", "der", "den", "ter", "ten", "het", "te"}
                        for deel in delen[1:]:
                            if deel.lower() not in tussenvoegsels:
                                achternaam = deel.capitalize()
                                break
                        code = voornaam
                        key = voornaam.lower()
                        if key in namen_counter:
                            code += achternaam[:1].upper()
                        namen_counter[key] = namen_counter.get(key, 0) + 1

                        locatie = "(B)" if pony in reeds_in_bak else "(S)"
                        kind_pony_combinaties.append((code, f"{pony.title()} {locatie}"))
                        reeds_in_bak.add(pony)

                    kind_pony_combinaties.sort(key=lambda x: x[0].lower())

                    with container:
                        st.markdown(f"<strong>Groep {tijd}</strong>", unsafe_allow_html=True)
                        st.markdown(f"<strong>Juf:</strong> {juf}</strong>", unsafe_allow_html=True)
                        for naam, pony in kind_pony_combinaties:
                            st.markdown(f"- {naam} – {pony}")

                # Ondertekst weergeven
                if opgeslagen_notitie["text"]:
                    stijl = "text-align: center;"
                    if opgeslagen_notitie["highlight"]:
                        stijl += " background-color: yellow;"
                    if opgeslagen_notitie["bold"]:
                        opgeslagen_notitie["text"] = f"<strong>{opgeslagen_notitie['text']}</strong>"
                    st.markdown(f"<div style='{stijl} margin-top: 10px;'>{opgeslagen_notitie['text']}</div>", unsafe_allow_html=True)

        else:
            st.warning("Kon geen rij met lestijden vinden.")
    else:
        st.warning("Kon geen kolom met >60 ponynamen vinden.")
else:
    st.info("Upload eerst een Excel-bestand om verder te gaan.")
