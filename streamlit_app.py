import streamlit as st
import pandas as pd
import datetime
import re
import locale

# Stel Nederlandse taal in voor datumnotatie
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

uploaded_file = st.file_uploader("📄 Upload je Excel-bestand", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📃 Kies een tabblad", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet, header=None)

    st.markdown("### 📊 Voorbeeld van de data")
    st.dataframe(df.head(20))

    # Zoek de kolom met >60 aaneengeschakelde pony-namen (niet leeg, niet numeriek)
    ponynamen_kolom = None
    ponynamen_start_index = 0
    for col in df.columns:
        tekst_rijen = df[col].dropna().astype(str)
        telling = 0
        start_index = None
        for idx, value in tekst_rijen.items():
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
                juf = "onbekend"
                max_rij = len(df)
                eigen_pony_rij = None

                # Zoek de rij waar "eigen pony" staat in de ponynamen-kolom, om af te kappen
                for i in range(ponynamen_start_index, len(df)):
                    waarde = str(df.iloc[i, ponynamen_kolom]).strip().lower()
                    if "eigen pony" in waarde:
                        eigen_pony_rij = i
                        max_rij = i
                        break

                # Nieuw: kijk 2 rijen onder 'eigen pony' in de groepskolom voor juf
                if eigen_pony_rij is not None and eigen_pony_rij + 2 < len(df):
                    mogelijke_juf = str(df.iloc[eigen_pony_rij + 2, col]).strip()
                    if mogelijke_juf and mogelijke_juf
