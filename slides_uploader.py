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
        slides = presentation.get('slides', [])

        if not slides:
            st.error("Geen slides gevonden in de presentatie.")
            return

        base_slide_id = slides[-1]['objectId']
        slides_to_delete = [s['objectId'] for s in slides[:-1]]

        if slides_to_delete:
            delete_requests = [{"deleteObject": {"objectId": sid}} for sid in slides_to_delete]
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        requests = []

        # 🔄 Aangepaste volgorde: reversed() voor correcte slide-positie
        for blok in reversed(st.session_state["slides_data"]):  # ← Kernwijziging hier
            slide_id = f"slide_{uuid.uuid4().hex[:8]}"
            requests.append({
                "duplicateObject": {
                    "objectId": base_slide_id,
                    "objectIds": {base_slide_id: slide_id}
                }
            })

            # ... (zelfde datum/textbox-code als eerder) ...

        # ... (zelfde ondertekst-logica) ...

        # Uitvoeren van alle requests
        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": requests}
        ).execute()

        # 🔄 Definitieve sjabloonpositie
        total_slides = len(service.presentations().get(presentationId=PRESENTATION_ID).execute().get('slides', []))
        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={
                "requests": [{
                    "updateSlidesPosition": {
                        "slideObjectIds": [base_slide_id],
                        "insertionIndex": total_slides  # Zet sjabloon altijd als laatste
                    }
                }]
            }
        ).execute()

        st.success("Slides succesvol geüpload!")

    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
