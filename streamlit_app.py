import streamlit as st
import pandas as pd
import datetime
import re

st.set_page_config(page_title="Het Zesspan Ponyplanner", layout="wide")

st.title("🐴 Het Zesspan Ponyplanner")

st.markdown("""
Upload hieronder het Excel-bestand met de planning. Kies daarna het juiste tabblad.
De app herkent automatisch de ponynamen, lestijden, kindernamen (geanonimiseerd) en juffen.
""")

uploaded_file = st.file_uploader("📄 Upload je Excel-bestand", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📃 Kies een tabblad", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet, header=None)

    st.markdown("### 📊 Voorbeeld van de data")
    st.dataframe(df.head(20))

    # Zoek de kolom met >60 aaneengeschakelde pony-namen (niet leeg, niet numeriek)
    ponynamen_kolom = None
    for col in df.columns:
        tekst_rijen = df[col].dropna().astype(str)
        telling = 0
        for value in tekst_rijen:
            if re.match(r"^[A-Za-z\- ]+$", value):
                telling += 1
                if telling >= 60:
                    ponynamen_kolom = col
                    break
            else:
                telling = 0
        if ponynamen_kolom is not None:
            break

    if ponynamen_kolom is not None:
        # Zoek de rij met tijden (formaat HH:MM of HH:MM-HH:MM)
        tijdrij = None
        tijd_pattern = re.compile(r"\b\d{1,2}:\d{2}(\s*-\s*\d{1,2}:\d{2})?\b")
        for i in range(0, 5):
            if any(tijd_pattern.search(str(cell)) for cell in df.iloc[i]):
                tijdrij = i
                break

        if tijdrij is not None:
            tijden = []
            kolommen = []
            for col in df.columns:
                cel = str(df.iloc[tijdrij, col])
                if tijd_pattern.match(cel):
                    tijden.append(cel)
                    kolommen.append(col)

            st.markdown("### 📅 Planning per groep")
            datum_vandaag = datetime.datetime.today().strftime("%A %d-%m-%Y")
            st.markdown(f"**Datum:** {datum_vandaag}")

            for tijd, col in zip(tijden, kolommen):
                kind_pony_combinaties = []
                juf = None
                for i in range(len(df)):
                    naam = str(df.iloc[i, col]) if col in df.columns and i < len(df) else ""
                    if not naam or naam.strip() == "" or naam.strip().lower() == "nan" or naam.strip().lower() == "x":
                        continue
                    if "juf" in naam.lower():
                        juf = naam
                        continue
                    pony = str(df.iloc[i, ponynamen_kolom]) if ponynamen_kolom in df.columns and i < len(df) else ""
                    delen = naam.strip().split()
                    voornaam = delen[0] if delen else ""
                    achternaam = ""
                    tussenvoegsels = {"van", "de", "der", "den", "ter", "ten", "het", "te"}
                    for deel in delen[1:]:
                        if deel.lower() not in tussenvoegsels:
                            achternaam = deel
                            break
                    code = voornaam
                    if sum(n.startswith(voornaam) for n, _ in kind_pony_combinaties) > 0:
                        code += achternaam[:1].upper()
                    kind_pony_combinaties.append((code, pony))

                if kind_pony_combinaties:
                    with st.container():
                        st.markdown(f"**Groep {tijd}**")
                        if juf:
                            st.markdown(f"Juf: **{juf}**")
                        else:
                            st.markdown("Juf: _onbekend_")
                        for naam, pony in kind_pony_combinaties:
                            st.markdown(f"- {naam} – {pony}")
        else:
            st.warning("Kon geen rij met lestijden vinden.")
    else:
        st.warning("Kon geen kolom met >60 ponynamen vinden.")
else:
    st.info("Upload eerst een Excel-bestand om verder te gaan.")
