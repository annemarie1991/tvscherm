from google.oauth2 import service_account
from googleapiclient.discovery import build
import uuid
import streamlit as st

PRESENTATION_ID = '1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4'

def upload_to_slides():
    if "slides_data" not in st.session_state or not st.session_state["slides_data"]:
        st.error("Geen slides om te uploaden.")
        return

    try:
        creds_dict = st.secrets["google_service_account"]
        credentials = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=["https://www.googleapis.com/auth/presentations"]
        )
        service = build('slides', 'v1', credentials=credentials)

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

            x_offset = 50
            y_offset = 50
            column_width = 200

            for i, col in enumerate(blok["columns"]):
                content = f"**{col['tijd']}**\nJuf: {col['juf']}\n\n"
                for kind, pony in col["kinderen"]:
                    content += f"{kind} – {pony}\n"

                box_id = f"textbox_{uuid.uuid4().hex[:8]}"
                box_x = x_offset + i * (column_width + 40)

                requests.append({
                    "createShape": {
                        "objectId": box_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "height": {"magnitude": 300, "unit": "PT"},
                                "width": {"magnitude": column_width, "unit": "PT"}
                            },
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": box_x, "translateY": y_offset,
                                "unit": "PT"
                            }
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": box_id,
                        "insertionIndex": 0,
                        "text": content
                    }
                })

            if blok.get("ondertekst"):
                tekst = blok["ondertekst"]
                bold = blok.get("vet", False)
                geel = blok.get("geel", False)

                onder_id = f"ondertekst_{uuid.uuid4().hex[:8]}"
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
                                "translateX": 50, "translateY": 400,
                                "unit": "PT"
                            }
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": onder_id,
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

                if style:
                    requests.append({
                        "updateTextStyle": {
                            "objectId": onder_id,
                            "style": style,
                            "textRange": {"type": "ALL"},
                            "fields": ",".join(style.keys())
                        }
                    })

        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": requests}
        ).execute()

        st.success("✅ Slides succesvol geüpload!")

    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
