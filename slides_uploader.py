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

        # Presentatie ophalen
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slides = presentation.get('slides', [])
        if not slides:
            st.error("Geen slides gevonden in de presentatie.")
            return

        base_slide_id = slides[-1]['objectId']  # Laatste slide = sjabloon
        slides_to_delete = [s['objectId'] for s in slides[:-1]]

        if slides_to_delete:
            delete_requests = [{"deleteObject": {"objectId": sid}} for sid in slides_to_delete]
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        requests = []

        # Slides in juiste volgorde (eerste slide eerst)
        for blok in reversed(st.session_state["slides_data"]):
            slide_id = f"slide_{uuid.uuid4().hex[:8]}"
            requests.append({
                "duplicateObject": {
                    "objectId": base_slide_id,
                    "objectIds": {base_slide_id: slide_id}
                }
            })

            datum_id = f"datum_{uuid.uuid4().hex[:8]}"
            requests.append({
                "createShape": {
                    "objectId": datum_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"height": {"magnitude": 30, "unit": "PT"},
                                 "width": {"magnitude": 600, "unit": "PT"}},
                        "transform": {
                            "scaleX": 1, "scaleY": 1,
                            "translateX": 50, "translateY": 5,
                            "unit": "PT"
                        }
                    }
                }
            })
            requests.append({
                "insertText": {
                    "objectId": datum_id,
                    "insertionIndex": 0,
                    "text": blok["title"].replace("Planning ", "")
                }
            })
            requests.append({
                "updateTextStyle": {
                    "objectId": datum_id,
                    "textRange": {"type": "ALL"},
                    "style": {"fontSize": {"magnitude": 14, "unit": "PT"}, "bold": True},
                    "fields": "fontSize,bold"
                }
            })
            requests.append({
                "updateParagraphStyle": {
                    "objectId": datum_id,
                    "textRange": {"type": "ALL"},
                    "style": {"alignment": "CENTER"},
                    "fields": "alignment"
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

                requests.append({
                    "updateParagraphStyle": {
                        "objectId": onder_id,
                        "textRange": {"type": "ALL"},
                        "style": {"alignment": "CENTER"},
                        "fields": "alignment"
                    }
                })

        # Alle slides aanmaken
        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": requests}
        ).execute()

        # 🟨 Zet sjabloonslide weer als laatste
        # Herlaad presentatie zodat we het correcte totaal hebben NA toevoeging van nieuwe slides
        updated_presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        total_slide_count = len(updated_presentation.get("slides", []))

        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={
                "requests": [
                    {
                        "updateSlidesPosition": {
                            "slideObjectIds": [base_slide_id],
                            "insertionIndex": total_slide_count - 1
                        }
                    }
                ]
            }
        ).execute()

        st.success("Slides succesvol geüpload!")

    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
