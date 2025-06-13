import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ID van je bestaande presentatie
PRESENTATION_ID = "1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4"

def parse_text_with_style(text):
    """
    Converteer eenvoudige markdownstijl (**vet** en [GEEL]) naar text elements voor Google Slides.
    """
    elements = []
    bold = False
    yellow = False
    buffer = ""

    def flush():
        nonlocal buffer, bold, yellow
        if buffer:
            style = {"fontSize": {"magnitude": 16, "unit": "PT"}}
            if bold:
                style["bold"] = True
            if yellow:
                style["foregroundColor"] = {
                    "opaqueColor": {
                        "rgbColor": {"red": 1, "green": 0.8, "blue": 0}
                    }
                }
            elements.append({"textRun": {"content": buffer, "style": style}})
            buffer = ""

    i = 0
    while i < len(text):
        if text[i:i+2] == "**":
            flush()
            bold = not bold
            i += 2
        elif text[i:i+6] == "[GEEL]":
            flush()
            yellow = True
            i += 6
        elif text[i:i+7] == "[/GEEL]":
            flush()
            yellow = False
            i += 7
        else:
            buffer += text[i]
            i += 1
    flush()
    return elements

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
                        "text": " "  # placeholder zodat styles kunnen worden toegepast
                    }
                },
                {
                    "updateTextStyle": {
                        "objectId": content_id,
                        "style": {"fontSize": {"magnitude": 16, "unit": "PT"}},
                        "textRange": {"type": "ALL"},
                        "fields": "fontSize"
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

            # Meer geavanceerde opmaak: splits tekst in runs met stijlen
            text_elements = parse_text_with_style(blok["content"])
            if text_elements:
                requests.append({
                    "deleteText": {
                        "objectId": content_id,
                        "textRange": {"type": "ALL"}
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": content_id,
                        "insertionIndex": 0,
                        "text": ""
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": content_id,
                        "insertionIndex": 0,
                        "text": "\n".join([e["textRun"]["content"] for e in text_elements])
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
