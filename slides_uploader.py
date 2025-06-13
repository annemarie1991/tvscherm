from google.oauth2 import service_account
from googleapiclient.discovery import build
import uuid
import streamlit as st
import re

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

def upload_to_slides():
    if "slides_data" not in st.session_state or not st.session_state["slides_data"]:
        st.error("Geen slides om te uploaden.")
        return

    try:
        # Gebruik credentials uit secrets.toml
        credentials_dict = st.secrets["google_service_account"]
        credentials = service_account.Credentials.from_service_account_info(
            credentials_dict, scopes=SCOPES
        )
        service = build('slides', 'v1', credentials=credentials)

        # Verwijder alle bestaande slides behalve de eerste
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
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
                "createSlide": {
                    "objectId": slide_id,
                    "insertionIndex": str(index + 1),
                    "slideLayoutReference": {"predefinedLayout": "BLANK"}
                }
            })

            # Titel bovenaan in het midden
            title_id = f"title_{uuid.uuid4().hex[:8]}"
            requests.append({
                "createShape": {
                    "objectId": title_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"height": {"magnitude": 40, "unit": "PT"},
                                 "width": {"magnitude": 400, "unit": "PT"}},
                        "transform": {
                            "scaleX": 1, "scaleY": 1,
                            "translateX": 150, "translateY": 10,
                            "unit": "PT"
                        }
                    }
                }
            })
            requests.append({
                "insertText": {
                    "objectId": title_id,
                    "insertionIndex": 0,
                    "text": blok["title"]
                }
            })
            requests.append({
                "updateTextStyle": {
                    "objectId": title_id,
                    "textRange": {"type": "ALL"},
                    "style": {"fontSize": {"magnitude": 18, "unit": "PT"}, "bold": True},
                    "fields": "fontSize,bold"
                }
            })

            # Kolommen (kinderen en juf)
            x_offset = 50
            y_offset = 60
            column_width = 200

            for i, col in enumerate(blok["columns"]):
                box_x = x_offset + i * (column_width + 40)
                box_y = y_offset

                content = f"**{col['tijd']}**\n**Juf: {col['juf']}**\n\n"
                for kind, pony in col["kinderen"]:
                    content += f"{kind} – {pony}\n"

                textbox_id = f"textbox_{uuid.uuid4().hex[:8]}"
                requests.append({
                    "createShape": {
                        "objectId": textbox_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "height": {"magnitude": 300, "unit": "PT"},
                                "width": {"magnitude": column_width, "unit": "PT"}
                            },
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
                                "textRange": {
                                    "type": "FIXED_RANGE",
                                    "startIndex": index_start,
                                    "endIndex": index_start + length
                                },
                                "style": element["textRun"]["style"],
                                "fields": ",".join(element["textRun"]["style"].keys())
                            }
                        })
                    index_start += length

            # Ondertekst (hoger geplaatst zodat zichtbaar)
            if blok.get("ondertekst"):
                onder_id = f"onder_{uuid.uuid4().hex[:8]}"
                requests.append({
                    "createShape": {
                        "objectId": onder_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "height": {"magnitude": 40, "unit": "PT"},
                                "width": {"magnitude": 600, "unit": "PT"}
                            },
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": 50, "translateY": 340,
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

        # Stuur alle requests naar de Slides API
        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": requests}
        ).execute()

        st.success("Slides succesvol geüpload!")

    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
