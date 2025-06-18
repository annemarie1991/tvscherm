import streamlit as st
import pandas as pd
import datetime
import re
import locale
from pathlib import Path
from slides_uploader import upload_to_slides
import json

st.set_page_config(page_title="Het Zesspan TV Scherm", layout="wide")

# 👉 Pony-opmerkingen initialiseren
pony_opmerkingen_pad = Path("pony_opmerkingen.json")
if "pony_opmerkingen" not in st.session_state:
    try:
        if pony_opmerkingen_pad.exists():
            with pony_opmerkingen_pad.open("r", encoding="utf-8") as f:
                st.session_state.pony_opmerkingen = json.load(f)
        else:
            st.session_state.pony_opmerkingen = {}
    except Exception:
        st.session_state.pony_opmerkingen = {}

if "verwijder_sleutel" not in st.session_state:
    st.session_state.verwijder_sleutel = None

# 👉 Pony-opmerkingen beheren in de zijbalk
if st.sidebar.checkbox("✏️ Pony-opmerkingen beheren"):
    st.sidebar.markdown("Voeg hier opmerkingen toe aan pony's. Opmerkingen worden getoond in de planning.")
    nieuwe_pony = st.sidebar.text_input("Pony-naam (of deel van naam)")
    nieuwe_opmerking = st.sidebar.text_input("Opmerking bij deze pony")
    if st.sidebar.button("➕ Opslaan/aanpassen"):
        if nieuwe_pony.strip():
            st.session_state.pony_opmerkingen[nieuwe_pony.strip()] = nieuwe_opmerking.strip()
            with pony_opmerkingen_pad.open("w", encoding="utf-8") as f:
                json.dump(st.session_state.pony_opmerkingen, f, ensure_ascii=False, indent=2)
            st.sidebar.success("Opmerking opgeslagen!")
    if st.session_state.pony_opmerkingen:
        st.sidebar.markdown("### 📋 Huidige opmerkingen")
        for naam, opm in st.session_state.pony_opmerkingen.items():
            cols = st.sidebar.columns([4, 1])
            cols[0].markdown(f"- **{naam}**: {opm}")
            if cols[1].button("🗑️", key=f"verwijder_{naam}"):
                st.session_state.verwijder_sleutel = naam
    # Verwijder indien nodig
    if st.session_state.verwijder_sleutel:
        st.session_state.pony_opmerkingen.pop(st.session_state.verwijder_sleutel, None)
        with pony_opmerkingen_pad.open("w", encoding="utf-8") as f:
            json.dump(st.session_state.pony_opmerkingen, f, ensure_ascii=False, indent=2)
        st.session_state.verwijder_sleutel = None
        st.sidebar.success("Opmerking verwijderd!")

# 👉 Basisinstellingen
try:
    locale.setlocale(locale.LC_TIME, 'nl_NL.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'nl_NL')
    except:
        pass

# 👉 Link-knoppen bovenaan
col1, col2 = st.columns(2)
with col1:
    st.link_button("📄 Bewerk presentatie", "https://docs.google.com/presentation/d/1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4/edit?slide=id.slide_27146aef")
with col2:
    st.link_button("🌐 Bekijk online", "https://www.hetzesspan.nl/tv")

st.title("🏄 Het Zesspan TV Scherm")

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

st.sidebar.header("📜 Ondertekst instellen")
nieuwe_tekst = st.sidebar.text_area("Tekst onderaan elke sectie", st.session_state.ondertekst or "")
vet = st.sidebar.checkbox("Dikgedrukt", value=st.session_state.vet)
geel = st.sidebar.checkbox("Geel markeren", value=st.session_state.geel)
if st.sidebar.button("📂 Opslaan"):
    st.session_state.ondertekst = nieuwe_tekst
    st.session_state.vet = vet
    st.session_state.geel = geel
    tekstpad.write_text(f"{nieuwe_tekst}\n{vet}\n{geel}", encoding="utf-8")
    st.sidebar.success("Tekst opgeslagen!")

st.markdown("Upload hieronder het Excel-bestand met de planning. Kies daarna het juiste tabblad.")

uploaded_file = st.file_uploader("📄 Upload je Excel-bestand", type=["xlsx"])
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📘 Kies een tabblad", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet, header=None)
    st.dataframe(df.head(20))

    # Zoek de rij waar 'eigen pony' staat in kolom D
    eigen_pony_rij = None
    for i in range(2, len(df)):
        cell = str(df.iloc[i, 3]).strip().lower() # kolom D = index 3
        if cell.startswith("eigen pony"):
            eigen_pony_rij = i
            break
        if cell.replace(" ", "").startswith("eigenpony"):
            eigen_pony_rij = i
            break

    if eigen_pony_rij is None:
        st.warning("Kon geen rij met 'eigen pony' vinden in kolom D.")
    else:
        # Bepaal kolommen met tijden in rij 2 (index 1), vanaf kolom E (index 4)
        tijd_kolommen = []
        tijd_pattern = re.compile(r"\d{1,2}:\d{2}")
        for col in range(4, df.shape[1]):
            val = str(df.iloc[1, col]).strip()
            if tijd_pattern.match(val):
                tijd_kolommen.append(col)
            else:
                break # Stop bij eerste lege/geen tijds-cel

        ponynamen_kolom = 3 # kolom D
        ponynamen_start_index = 2 # Rij 3 (index 2)
        max_rij = eigen_pony_rij

        groepen_per_blok = []
        blok = []
        laatst_verwerkte_tijd = None
        tijd_dict = {}
        for col in tijd_kolommen:
            tijd = str(df.iloc[1, col]).strip()
            tijd_dict[col] = tijd
        tijd_items = sorted(
            tijd_dict.items(),
            key=lambda x: datetime.datetime.strptime(
                re.search(r"\d{1,2}:\d{2}", x[1]).group(), "%H:%M"
            )
        )
        for col, tijd in tijd_items:
            tijd_match = re.search(r"\d{1,2}:\d{2}", tijd)
            if not tijd_match:
                continue
            tijd_dt = datetime.datetime.strptime(tijd_match.group(), "%H:%M")
            if laatst_verwerkte_tijd is None or (tijd_dt - laatst_verwerkte_tijd).total_seconds() > 30 * 60:
                if blok:
                    groepen_per_blok.append(blok)
                blok = [(col, tijd)]
                laatst_verwerkte_tijd = tijd_dt
            else:
                blok.append((col, tijd))
        if blok:
            groepen_per_blok.append(blok)

        datum_vandaag = datetime.datetime.today().strftime("%d-%m-%Y")
        slides_data = []

        # --- S/B-logica met 10-minutenregel ---
        pony_last_end = {}

        for blok in groepen_per_blok:
            blok_kolommen = []
            for col, tijd in blok:
                # Juf staat altijd 2 rijen onder de rij van 'eigen pony'
                juf_rij = eigen_pony_rij + 2
                juf = str(df.iloc[juf_rij, col]).strip().title() if pd.notna(df.iloc[juf_rij, col]) else "Onbekend"
                kind_pony_combinaties = []
                namen_counter = {}
                for r in range(ponynamen_start_index, max_rij):
                    naam = str(df.iloc[r, col])
                    pony = str(df.iloc[r, ponynamen_kolom])
                    if not naam.strip() or naam.strip().lower() in ["", "nan", "x"]:
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

                    opmerking = ""
                    for sleutel, tekst in st.session_state.pony_opmerkingen.items():
                        if sleutel.lower() in pony.lower():
                            opmerking = f" ({tekst})"
                            break

                    # --- Hier de aangepaste S/B-logica ---
                    starttijd_match = re.search(r"\d{1,2}:\d{2}", tijd)
                    if starttijd_match:
                        starttijd_dt = datetime.datetime.strptime(starttijd_match.group(), "%H:%M")
                        eindtijd_dt = starttijd_dt + datetime.timedelta(minutes=30)  # Pas aan als lesduur anders is
                    else:
                        starttijd_dt = None
                        eindtijd_dt = None

                    in_bak = False
                    if pony in pony_last_end and starttijd_dt:
                        tijdverschil = (starttijd_dt - pony_last_end[pony]).total_seconds()
                        if 0 < tijdverschil <= 600:
                            in_bak = True
                    locatie = "(B)" if in_bak else "(S)"
                    pony_last_end[pony] = eindtijd_dt
                    # --- einde S/B-logica ---

                    pony_tekst = f"{pony.title()} {locatie}{opmerking}"
                    kind_pony_combinaties.append((code, pony_tekst))
                kind_pony_combinaties.sort(key=lambda x: x[0].lower())
                blok_kolommen.append({
                    "tijd": tijd,
                    "juf": juf,
                    "kinderen": kind_pony_combinaties
                })
            for i in range(0, len(blok_kolommen), 3):
                slides_data.append({
                    "title": f"Planning {datum_vandaag}",
                    "columns": blok_kolommen[i:i + 3],
                    "ondertekst": st.session_state.ondertekst.strip(),
                    "vet": st.session_state.vet,
                    "geel": st.session_state.geel
                })
        st.session_state["slides_data"] = slides_data
        st.success("Planning is verwerkt. Je kunt nu uploaden.")

        if st.button("📄 Upload naar (online) scherm"):
            upload_to_slides()

        st.markdown("### 📋 Voorbeeld weergave van slides")
        for idx, blok in enumerate(slides_data):
            st.markdown(f"**Slide {idx + 1}: {blok['title']}**")
            cols = st.columns(3)
            for i, coldata in enumerate(blok["columns"]):
                with cols[i]:
                    st.markdown(f"**{coldata['tijd']}**")
                    st.markdown(f"**Juf: {coldata['juf']}**")
                    for kind, pony in coldata["kinderen"]:
                        st.markdown(f"{kind} – {pony}")
            if blok.get("ondertekst"):
                stijl = "**" if blok.get("vet") else ""
                kleur = '' if blok.get("geel") else ""
                einde = "" if kleur else ""
                st.markdown(f"{kleur}{stijl}{blok['ondertekst']}{stijl}{einde}", unsafe_allow_html=True)
else:
    st.info("Upload eerst een Excel-bestand om verder te gaan.")
