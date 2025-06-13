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

    # Zoek de kolom met >60 aaneengeschakelde pony-namen
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
        # Zoek de rij met tijden
        tijdrij = None
        tijd_pattern = re.compile(r"\b\d{1,2}:\d{2}(\s*[-–−]\s*\d{1,2}:\d{2})?\b")
        for i in range(0, 5):
            if any(tijd_pattern.search(str(cell)) for cell in df.iloc[i]):
                tijdrij = i
                break

        if tijdrij is not None:
            tijd_dict = {}
            for col in df.columns:
                cel = str(df.iloc[tijdrij, col])
                if tijd_pattern.match(cel):
                    tijd_dict.setdefault(cel.strip(), []).append(col)

            tijd_lijst = list(tijd_dict.items())
            tijd_lijst.sort(key=lambda x: x[0])

            st.markdown("### 📅 Planning per groep")
            datum_vandaag = datetime.datetime.today().strftime("%A %d-%m-%Y")
            st.markdown(f"**Datum:** {datum_vandaag}")

            gebruikte_tijden = set()
            i = 0
            while i < len(tijd_lijst):
                hoofd_tijd, _ = tijd_lijst[i]
                if hoofd_tijd in gebruikte_tijden:
                    i += 1
                    continue

                basis_tijd_clean = re.sub(r"[–−]", "-", hoofd_tijd)
                basis_tijd_clean = re.split(r"[-\s]", basis_tijd_clean)[0].strip()
                try:
                    basis_tijd = datetime.datetime.strptime(basis_tijd_clean, "%H:%M")
                except ValueError:
                    i += 1
                    continue

                gekoppelde_kolommen = []
                j = i
                while j < len(tijd_lijst):
                    tijd, kolset = tijd_lijst[j]
                    tijd_clean = re.sub(r"[–−]", "-", tijd)
                    tijd_clean = re.split(r"[-\s]", tijd_clean)[0].strip()
                    try:
                        test_tijd = datetime.datetime.strptime(tijd_clean, "%H:%M")
                        verschil = (test_tijd - basis_tijd).total_seconds() / 60
                        if 0 <= verschil <= 30:
                            gekoppelde_kolommen.append((tijd, kolset))
                            gebruikte_tijden.add(tijd)
                            j += 1
                        else:
                            break
                    except ValueError:
                        break

                # Nu renderen we deze set van groepen naast elkaar
                columns = st.columns(len(gekoppelde_kolommen))
                for (tijd, kolset), col_container in zip(gekoppelde_kolommen, columns):
                    kind_pony_combinaties = []
                    juf = "onbekend"
                    max_rij = len(df)
                    eigen_pony_rij = None

                    for r in range(ponynamen_start_index, len(df)):
                        waarde = str(df.iloc[r, ponynamen_kolom]).strip().lower()
                        if "eigen pony" in waarde:
                            eigen_pony_rij = r
                            max_rij = r
                            break

                    if eigen_pony_rij is not None and eigen_pony_rij + 2 < len(df):
                        for col in kolset:
                            mogelijke_juf = str(df.iloc[eigen_pony_rij + 2, col]).strip()
                            if mogelijke_juf and mogelijke_juf.lower() != "nan":
                                juf = mogelijke_juf.title()
                                break

                    namen_counter = {}

                    for col in kolset:
                        for r in range(ponynamen_start_index, max_rij):
                            ponycel = str(df.iloc[r, ponynamen_kolom]) if r < len(df) else ""
                            naam = str(df.iloc[r, col]) if col in df.columns and r < len(df) else ""

                            if not naam or naam.strip().lower() in ["", "nan", "x"]:
                                continue

                            pony = ponycel.strip().title()
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
                            kind_pony_combinaties.append((code, pony))

                    kind_pony_combinaties.sort(key=lambda x: x[0].lower())

                    with col_container:
                        st.markdown(f"<strong>Groep {tijd}</strong>", unsafe_allow_html=True)
                        st.markdown(f"<strong>Juf:</strong> {juf}", unsafe_allow_html=True)
                        for naam, pony in kind_pony_combinaties:
                            st.markdown(f"- {naam} – {pony}")

                i = j
        else:
            st.warning("Kon geen rij met lestijden vinden.")
    else:
        st.warning("Kon geen kolom met >60 ponynamen vinden.")
else:
    st.info("Upload eerst een Excel-bestand om verder te gaan.")
