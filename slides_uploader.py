from google.oauth2 import service_account
from googleapiclient.discovery import build
import streamlit as st

# Google Slides presentatie ID
PRESENTATION_ID = "1uhfZtV-ota0vNrmQyLi6hmxRMvkuo1qM"

def upload_to_slides():
    st.info("Uploaden naar Google Slides gestart...")

    try:
        credentials = service_account.Credentials.from_service_account_file(
            "service_account.json",
            scopes=["https://www.googleapis.com/auth/presentations"]
        )
        service = build("slides", "v1", credentials=credentials)

        # Voorbeeldbewerking: lege dia toevoegen
        requests = [
            {
                "createSlide": {
                    "insertionIndex": "1"
                }
            }
        ]

        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID, body={"requests": requests}
        ).execute()

        st.success("Upload voltooid (test-dias toegevoegd)!")
    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
