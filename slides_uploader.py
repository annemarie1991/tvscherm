import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from uuid import uuid4

# ID van je bestaande presentatie
PRESENTATION_ID = "1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4"

def create_textbox(object_id, slide_id, x, y, width, height):
    return {
        "createShape": {
            "objectId": object_id,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": slide_id,
                "size": {
                    "height": {"magnitude": height, "unit": "PT"},
                    "width": {"magnitude": width, "unit": "PT"},
                },
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": x,
                    "translateY": y,
                    "unit": "PT",
                },
            }
        }
    }

def upload_to_slides():
    try:
        st.info("Uploaden naar Google Slides gestart...")

        creds_dict = st.secrets["google_service_account"]
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/presentations"],
        )

        service = build("slides", "v1", credentials=creds)

        # Verwijder bestaande slides behalve de eerste
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slide_ids = [slide["objectId"] for slide in presentation.get("slides", [])[1:]]
        delete_requests = [{"deleteObject": {"objectId": sid}} for sid in slide_ids]
        all_requests = delete_requests.copy()

        blokken = [b for b in st.session_state.get("slides_data", []) if b.get("columns")]

        for blok in blokken:
            slide_id = f"slide_{uuid4().hex}"
            all_requests.append({
                "createSlide": {
                    "objectId": slide_id,
                    "insertionIndex": 1,
                    "slideLayoutReference": {"predefinedLayout": "BLANK"},
                }
            })

            # Titel (datum)
            datum_id = f"datum_{uuid4().hex}"
            all_requests.append(create_textbox(datum_id, slide_id, 50, 20, 600, 50))
            all_requests.append({
                "insertText": {
                    "objectId": datum_id,
                    "insertionIndex": 0,
                    "text": blok.get("title", "Planning"),
                }
            })
            all_requests.append({
                "updateTextStyle": {
                    "objectId": datum_id,
                    "textRange": {"type": "ALL"},
                    "style": {"fontSize": {"magnitude": 20, "unit": "PT"}, "bold": True},
                    "fields": "fontSize,bold",
                }
            })

            # Elke groep in kolommen (maximaal 3 per slide)
            columns = blok.get("columns", [])
            for i, col in enumerate(columns):
                x = 50 + i * 190
                y = 80
                col_id = f"col_{uuid4().hex}"
                all_requests.append(create_textbox(col_id, slide_id, x, y, 180, 400))

                regels = [f"**{col['tijd']}**", f"Juf: {col['juf']}"]
                regels += [f"{k} – {p}" for k, p in col.get("kinderen", [])]

                tekst = "\n".join(regels)
                all_requests.append({
                    "insertText": {
                        "objectId": col_id,
                        "insertionIndex": 0,
                        "text": tekst,
                    }
                })

            # Ondertekst
            ondertekst = blok.get("ondertekst", "")
            if ondertekst:
                onder_id = f"onder_{uuid4().hex}"
                all_requests.append(create_textbox(onder_id, slide_id, 50, 490, 600, 50))
                all_requests.append({
                    "insertText": {
                        "objectId": onder_id,
                        "insertionIndex": 0,
                        "text": ondertekst,
                    }
                })
                fields = []
                style = {"fontSize": {"magnitude": 14, "unit": "PT"}}
                if blok.get("vet"):
                    style["bold"] = True
                    fields.append("bold")
                if blok.get("geel"):
                    style["foregroundColor"] = {
                        "opaqueColor": {
                            "rgbColor": {"red": 1, "green": 0.8, "blue": 0}
                        }
                    }
                    fields.append("foregroundColor")
                fields.append("fontSize")
                all_requests.append({
                    "updateTextStyle": {
                        "objectId": onder_id,
                        "textRange": {"type": "ALL"},
                        "style": style,
                        "fields": ",".join(fields),
                    }
                })

        # Upload naar Google Slides
        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": all_requests}
        ).execute()

        st.success("✅ Succesvol geüpload naar Google Slides!")

    except HttpError as error:
        st.error(f"Fout bij Google Slides API: {error}")
    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
