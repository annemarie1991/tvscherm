from google.oauth2 import service_account
from googleapiclient.discovery import build
import uuid
import os
import streamlit as st

SCOPES = ['https://www.googleapis.com/auth/presentations']
SERVICE_ACCOUNT_FILE = 'service_account.json'
PRESENTATION_ID = '1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4'

def upload_to_slides():
    if "slides_data" not in st.session_state or not st.session_state["slides_data"]:
        st.error("Geen slides om te uploaden.")
        return

    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        service = build('slides', 'v1', credentials=credentials)

        # Verwijder eerst alle bestaande slides behalve de eerste
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slide_ids = [slide['objectId'] for slide in presentation.get('slides', [])[1:]]

        delete_requests = [{
            "deleteObject": {
                "objectId": slide_id
            }
        } for slide_id in slide_ids]

        if delete_requests:
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        # Voeg nieuwe slides toe
        requests = []
        for index, blok in enumerate(st.session_state["slides_data"]):
            slide_id = f"slide_{uuid.uuid4().hex[:8]}"
            requests.append({
                "createSlide": {
                    "objectId": slide_id,
                    "insertionIndex": str(index + 1),
                    "slideLayoutReference": {
                        "predefinedLayout": "BLANK"
                    }
                }
            })

            x_offset = 50
            y_offset = 50
            column_width = 200
            box_height = 20
            spacing = 30

            for i, col in enumerate(blok["columns"]):
                box_y = y_offset
                box_x = x_offset + i * (column_width + 40)

                # Tijd en juf
                content = f"**{col['tijd']}**\nJuf: {col['juf']}\n\n"
                for kind, pony in col["kinderen"]:
                    content += f"{kind} – {pony}\n"

                requests.append({
                    "createShape": {
                        "objectId": f"textbox_{uuid.uuid4().hex[:8]}",
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
                requests.append({
                    "insertText": {
                        "objectId": requests[-1]["createShape"]["objectId"],
                        "insertionIndex": 0,
                        "text": content
                    }
                })

            # Ondertekst
            if blok.get("ondertekst"):
                tekst = blok["ondertekst"]
                bold = blok.get("vet", False)
                geel = blok.get("geel", False)

                requests.append({
                    "createShape": {
                        "objectId": f"ondertekst_{uuid.uuid4().hex[:8]}",
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "height": {"magnitude": 40, "unit": "PT"},
                                "width": {"magnitude": 600, "unit": "PT"}
                            },
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": 50, "translateY": 400,
                                "unit": "PT"
                            }
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": requests[-1]["createShape"]["objectId"],
                        "insertionIndex": 0,
                        "text": tekst
                    }
                })
                style = {}
                if bold:
                    style["bold"] = True
                if geel:
                    style["foregroundColor"] = {
                        "opaqueColor": {"rgbColor": {"red": 1, "green": 0.84, "blue": 0}}
                    }

                requests.append({
                    "updateTextStyle": {
                        "objectId": requests[-2]["insertText"]["objectId"],
                        "style": style,
                        "textRange": {"type": "ALL"},
                        "fields": ",".join(style.keys())
                    }
                })

        # Upload naar Google Slides
        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": requests}
        ).execute()

        st.success("Slides succesvol geüpload!")

    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
