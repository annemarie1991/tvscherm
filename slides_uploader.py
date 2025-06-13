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
        if text[i:i + 2] == "**":
            flush()
            bold = not bold
            i += 2
        elif text[i:i + 6] == "[GEEL]":
            flush()
            yellow = True
            i += 6
        elif text[i:i + 7] == "[/GEEL]":
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

        # Verwijder alle bestaande slides behalve de eerste
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slide_ids = [slide["objectId"] for slide in presentation.get("slides", [])[1:]]
        delete_requests = [{"deleteObject": {"objectId": sid}} for sid in slide_ids]
        if delete_requests:
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        # Hernoem eerste slide als titel slide
        slide_index = 0
        blokken = st.session_state.get("slides_data", [])

        # Verdeel blokken in sets van 4 groepen per slide
        for blok_index in range(0, len(blokken), 4):
            blokgroep = blokken[blok_index:blok_index + 4]
            slide_id = f"slide_{slide_index}"
            requests = []

            if slide_index > 0:
                requests.append({
                    "createSlide": {
                        "objectId": slide_id,
                        "insertionIndex": slide_index,
                        "slideLayoutReference": {"predefinedLayout": "BLANK"}
                    }
                })

            # Voeg datum toe
            requests.append({
                "createShape": {
                    "objectId": f"datum_{slide_index}",
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"height": {"magnitude": 50, "unit": "PT"}, "width": {"magnitude": 600, "unit": "PT"}},
                        "transform": {
                            "scaleX": 1, "scaleY": 1,
                            "translateX": 50, "translateY": 20,
                            "unit": "PT"
                        }
                    }
                }
            })
            requests.append({
                "insertText": {
                    "objectId": f"datum_{slide_index}",
                    "insertionIndex": 0,
                    "text": blokgroep[0]["title"]
                }
            })

            # Bereken fontgrootte afhankelijk van aantal groepen
            font_size = 16 if len(blokgroep) <= 2 else 14 if len(blokgroep) == 3 else 12

            # Voeg elke groep toe als kolom
            for j, blok in enumerate(blokgroep):
                content_id = f"groep_{slide_index}_{j}"
                x_offset = 50 + j * 160

                requests.append({
                    "createShape": {
                        "objectId": content_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 400, "unit": "PT"}, "width": {"magnitude": 150, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": x_offset, "translateY": 80,
                                "unit": "PT"
                            }
                        }
                    }
                })
                text_elements = parse_text_with_style(blok["content"] + "\n")
                if text_elements:
                    requests.append({
                        "insertText": {
                            "objectId": content_id,
                            "insertionIndex": 0,
                            "text": "".join([e["textRun"]["content"] for e in text_elements])
                        }
                    })
                    for k, e in enumerate(text_elements):
                        requests.append({
                            "updateTextStyle": {
                                "objectId": content_id,
                                "textRange": {
                                    "type": "FIXED_RANGE",
                                    "startIndex": sum(len(te["textRun"]["content"]) for te in text_elements[:k]),
                                    "endIndex": sum(len(te["textRun"]["content"]) for te in text_elements[:k + 1])
                                },
                                "style": e["textRun"]["style"],
                                "fields": ",".join(e["textRun"]["style"].keys())
                            }
                        })

            slide_index += 1

            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": requests}
            ).execute()

        st.success("✅ Succesvol geüpload naar Google Slides!")

    except HttpError as error:
        st.error(f"Fout bij Google Slides API: {error}")
    except Exception as e:
        st.error(f"Fout tijdens up
