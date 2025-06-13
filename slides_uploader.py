import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ID van je bestaande presentatie
PRESENTATION_ID = "1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4"

def upload_to_slides():
    try:
        st.info("Uploaden naar Google Slides gestart...")

        # Haal credentials op uit Streamlit secrets
        creds_dict = st.secrets["google_service_account"]
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/presentations"]
        )

        service = build("slides", "v1", credentials=creds)

        # Verwijder alle bestaande slides behalve de eerste
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slide_ids = [slide["objectId"] for slide in presentation.get("slides", [])[1:]]
        delete_requests = [{"deleteObject": {"objectId": sid}} for sid in slide_ids]
        if delete_requests:
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        # Voeg nieuwe slides toe
        for i, blok in enumerate(st.session_state.get("slides_data", []), start=1):
            slide_id = f"slide_{i}"
            title_id = f"title_{i}"
            content_id = f"content_{i}"

            requests = [
                {
                    "createSlide": {
                        "objectId": slide_id,
                        "insertionIndex": i,
                        "slideLayoutReference": {"predefinedLayout": "BLANK"}
                    }
                },
                {
                    "createShape": {
                        "objectId": title_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 100, "unit": "PT"}, "width": {"magnitude": 500, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": 50, "translateY": 20,
                                "unit": "PT"
                            }
                        }
                    }
                },
                {
                    "insertText": {
                        "objectId": title_id,
                        "insertionIndex": 0,
                        "text": blok["title"]
                    }
                },
                {
                    "createShape": {
                        "objectId": content_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 400, "unit": "PT"}, "width": {"magnitude": 600, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": 50, "translateY": 130,
                                "unit": "PT"
                            }
                        }
                    }
                },
                {
                    "insertText": {
                        "objectId": content_id,
                        "insertionIndex": 0,
                        "text": blok["content"]
                    }
                }
            ]

            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": requests}
            ).execute()

        st.success("✅ Succesvol geüpload naar Google Slides!")

    except HttpError as error:
        st.error(f"Fout bij Google Slides API: {error}")
    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
