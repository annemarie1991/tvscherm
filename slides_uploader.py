from google.oauth2 import service_account
from googleapiclient.discovery import build
import uuid
import streamlit as st
import re
import json
from pathlib import Path

SCOPES = ['https://www.googleapis.com/auth/presentations']
PRESENTATION_ID = '1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4'

def parse_markdown_to_text_elements(text):
    elements = []
    bold = False
    buffer = ""

    def flush():
        nonlocal buffer, bold
        if buffer:
            style = {}
            if bold:
                style["bold"] = True
            elements.append({"textRun": {"content": buffer, "style": style}})
            buffer = ""

    i = 0
    while i < len(text):
        if text[i:i+2] == "**":
            flush()
            bold = not bold
            i += 2
        else:
            buffer += text[i]
            i += 1
    flush()
    return elements

def pony_opmerking(pony_naam: str) -> str:
    pad = Path("pony_opmerkingen.json")
    if not pad.exists():
        return ""
    try:
        with pad.open("r", encoding="utf-8") as f:
            opmerkingen = json.load(f)
        for sleutel, tekst in opmerkingen.items():
            if sleutel.lower() in pony_naam.lower() and tekst.lower() not in pony_naam.lower():
                return f" ({tekst})"
    except Exception:
        pass
    return ""

def upload_to_slides():
    if "slides_data" not in st.session_state or not st.session_state["slides_data"]:
        st.error("Geen slides om te uploaden.")
        return

    try:
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["google_service_account"], scopes=SCOPES
        )
        service = build('slides', 'v1', credentials=credentials)

        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        base_slide_id = presentation.get('slides', [])[0].get('objectId')  # ← eerste slide als sjabloon
        slide_ids = [slide['objectId'] for slide in presentation.get('slides', [])[1:]]
        delete_requests = [{"deleteObject": {"objectId": slide_id}} for slide_id in slide_ids]

        if delete_requests:
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        requests = []

        for index, blok in enumerate(st.session_state["slides_data"]):
            slide_id = f"slide_{uuid.uuid4().hex[:8]}"
            requests.append({
                "duplicateObject": {
                    "objectId": base_slide_id,
                    "objectIds": {
                        base_slide_id: slide_id
                    }
                }
            })

            x_offset = 50
            y_offset = 60
            column_width = 200

            for i, col in enumerate(blok["columns"]):
                box_x = x_offset + i * (column_width + 40)
                box_y = y_offset

                content = f"**{col['tijd']}**\n**Juf: {col['juf']}**\n\n"
                for kind, pony in col["kinderen"]:
                    opm = pony_opmerking(pony)
                    content += f"{kind} – {pony}{opm}\n"

                textbox_id = f"textbox_{uuid.uuid4().hex[:8]}"
                requests.append({
                    "createShape": {
                        "objectId": textbox_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 300, "unit": "PT"},
                                     "width": {"magnitude": column_width, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": box_x, "translateY": box_y,
                                "unit": "PT"
                            }
                        }
                    }
                })

                parsed = parse_markdown_to_text_elements(content)
                full_text = "".join([e["textRun"]["content"] for e in parsed])
                requests.append({
                    "insertText": {
                        "objectId": textbox_id,
                        "insertionIndex": 0,
                        "text": full_text
                    }
                })

                index_start = 0
                for element in parsed:
                    length = len(element["textRun"]["content"])
                    if element["textRun"].get("style"):
                        requests.append({
                            "updateTextStyle": {
                                "objectId": textbox_id,
                                "textRange": {"type": "FIXED_RANGE", "startIndex": index_start, "endIndex": index_start + length},
                                "style": element["textRun"]["style"],
                                "fields": ",".join(element["textRun"]["style"].keys())
                            }
                        })
                    index_start += length

            # Ondertekst
            if blok.get("ondertekst"):
                onder_id = f"onder_{uuid.uuid4().hex[:8]}"
                requests.append({
                    "createShape": {
                        "objectId": onder_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 80, "unit": "PT"},
                                     "width": {"magnitude": 600, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": 50, "translateY": 330,
                                "unit": "PT"
                            }
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": onder_id,
                        "insertionIndex": 0,
                        "text": blok["ondertekst"]
                    }
                })

                style = {}
                if blok.get("vet"):
                    style["bold"] = True
                if blok.get("geel"):
                    style["foregroundColor"] = {
                        "opaqueColor": {"rgbColor": {"red": 1, "green": 0.84, "blue": 0}}
                    }

                if style:
                    requests.append({
                        "updateTextStyle": {
                            "objectId": onder_id,
                            "textRange": {"type": "ALL"},
                            "style": style,
                            "fields": ",".join(style.keys())
                        }
                    })

        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": requests}
        ).execute()

        st.success("Slides succesvol geüpload!")

    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")


En hier de code van streamlit_app

import streamlit as st
import pandas as pd
import datetime
import re
import locale
from pathlib import Path
from slides_uploader import upload_to_slides
import json

st.set_page_config(page_title="Het Zesspan TV Scherm", layout="wide")  # ← Moet als eerste st.-commando

# 👉 Pony-opmerkingen initialiseren
pony_opmerkingen_pad = Path("pony_opmerkingen.json")
if "pony_opmerkingen" not in st.session_state:
    if pony_opmerkingen_pad.exists():
        with pony_opmerkingen_pad.open("r", encoding="utf-8") as f:
            st.session_state.pony_opmerkingen = json.load(f)
    else:
        st.session_state.pony_opmerkingen = {}

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
            st.sidebar.markdown(f"- **{naam}**: {opm}")

# 👉 Basisinstellingen
try:
    locale.setlocale(locale.LC_TIME, 'nl_NL.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'nl_NL')
    except:
        pass

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
        tijd_pattern = re.compile(r"\b\d{1,2}:\d{2}(\s*[-\u2013−]\s*\d{1,2}:\d{2})?\b")
        tijdrij = next((i for i in range(0, 5) if any(tijd_pattern.search(str(cell)) for cell in df.iloc[i])), None)

        if tijdrij is not None:
            tijd_dict = {
                col: str(df.iloc[tijdrij, col]).strip()
                for col in df.columns
                if tijd_pattern.match(str(df.iloc[tijdrij, col]))
            }

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
                tijd_match = re.search(r"\d{1,2}:\d{2}", tijd)
                if not tijd_match:
                    continue
                tijd_dt = datetime.datetime.strptime(tijd_match.group(), "%H:%M")

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
            eigen_pony_rij = next((r for r in range(ponynamen_start_index, len(df))
                                   if "eigen pony" in str(df.iloc[r, ponynamen_kolom]).strip().lower()), None)

            reeds_in_bak = set()
            slides_data = []

            for blok in groepen_per_blok:
                blok_kolommen = []
                for col, tijd in blok:
                    max_rij = eigen_pony_rij if eigen_pony_rij else len(df)
                    juf = str(df.iloc[eigen_pony_rij + 2, col]).strip().title() if eigen_pony_rij and pd.notna(df.iloc[eigen_pony_rij + 2, col]) else "Onbekend"

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

                        # Zoek opmerking bij deze pony
                        opmerking = ""
                        for sleutel, tekst in st.session_state.pony_opmerkingen.items():
                            if sleutel.lower() in pony.lower():
                                opmerking = f" ({tekst})"
                                break

                        locatie = "(B)" if pony in reeds_in_bak else "(S)"
                        pony_tekst = f"{pony.title()} {locatie}{opmerking}"
                        kind_pony_combinaties.append((code, pony_tekst))
                        reeds_in_bak.add(pony)

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
                    kleur = '<span style="color:gold">' if blok.get("geel") else ""
                    einde = "</span>" if kleur else ""
                    st.markdown(f"{kleur}{stijl}{blok['ondertekst']}{stijl}{einde}", unsafe_allow_html=True)
        else:
            st.warning("Kon geen rij met lestijden vinden.")
    else:
        st.warning("Kon geen kolom met >60 ponynamen vinden.")
else:
    st.info("Upload eerst een Excel-bestand om verder te gaan.")
