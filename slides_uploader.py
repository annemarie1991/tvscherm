import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ID van je bestaande presentatie
PRESENTATION_ID = "1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4"

MAX_KOLOMMEN_PER_SLIDE = 4


def upload_to_slides():
    try:
        st.info("Uploaden naar Google Slides gestart...")

        creds_dict = st.secrets["google_service_account"]
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/presentations"]
        )

        service = build("slides", "v1", credentials=creds)

        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slide_ids = [slide["objectId"] for slide in presentation.get("slides", [])[1:]]
        delete_requests = [{"deleteObject": {"objectId": sid}} for sid in slide_ids]
        if delete_requests:
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        # Genereer slides op basis van blokken, maximaal 4 per slide
        slides_data = st.session_state.get("slides_data", [])
        for page_index in range(0, len(slides_data), MAX_KOLOMMEN_PER_SLIDE):
            subset = slides_data[page_index:page_index + MAX_KOLOMMEN_PER_SLIDE]
            slide_id = f"slide_{page_index}"
            requests = [
                {
                    "createSlide": {
                        "objectId": slide_id,
                        "insertionIndex": page_index,
                        "slideLayoutReference": {"predefinedLayout": "BLANK"}
                    }
                }
            ]

            # Titel (datum) gecentreerd bovenaan
            title_id = f"title_{page_index}"
            requests.append({
                "createShape": {
                    "objectId": title_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"height": {"magnitude": 50, "unit": "PT"}, "width": {"magnitude": 500, "unit": "PT"}},
                        "transform": {
                            "scaleX": 1, "scaleY": 1,
                            "translateX": 100, "translateY": 20,
                            "unit": "PT"
                        }
                    }
                }
            })
            requests.append({
                "insertText": {
                    "objectId": title_id,
                    "insertionIndex": 0,
                    "text": subset[0]["title"]
                }
            })
            requests.append({
                "updateTextStyle": {
                    "objectId": title_id,
                    "textRange": {"type": "ALL"},
                    "style": {"bold": True, "fontSize": {"magnitude": 24, "unit": "PT"}},
                    "fields": "bold,fontSize"
                }
            })

            # Kolommen
            kolombreedte = 150
            spacing = 20
            start_x = 50
            for i, blok in enumerate(subset):
                content_id = f"content_{page_index}_{i}"
                x = start_x + i * (kolombreedte + spacing)
                requests.append({
                    "createShape": {
                        "objectId": content_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 400, "unit": "PT"}, "width": {"magnitude": kolombreedte, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": x, "translateY": 100,
                                "unit": "PT"
                            }
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": content_id,
                        "insertionIndex": 0,
                        "text": blok["content"]
                    }
                })

            # Ondertekst
            onder = st.session_state.get("ondertekst", "")
            if onder:
                ondertekst_id = f"onder_{page_index}"
                requests.append({
                    "createShape": {
                        "objectId": ondertekst_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 50, "unit": "PT"}, "width": {"magnitude": 600, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": 50, "translateY": 520,
                                "unit": "PT"
                            }
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": ondertekst_id,
                        "insertionIndex": 0,
                        "text": onder
                    }
                })
                stijl = {}
                if st.session_state.get("vet"):
                    stijl["bold"] = True
                if st.session_state.get("geel"):
                    stijl["foregroundColor"] = {
                        "opaqueColor": {
                            "rgbColor": {"red": 1, "green": 0.8, "blue": 0}
                        }
                    }
                if stijl:
                    requests.append({
                        "updateTextStyle": {
                            "objectId": ondertekst_id,
                            "textRange": {"type": "ALL"},
                            "style": stijl,
                            "fields": ",".join(stijl.keys())
                        }
                    })

            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": requests}
            ).execute()

        st.success("✅ Succesvol geüpload naar Google Slides!")

    except HttpError as error:
        st.error(f"Fout bij Google Slides API: {error}")
    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
