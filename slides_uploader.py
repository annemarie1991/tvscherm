import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ID van je bestaande presentatie
PRESENTATION_ID = "1vuVUa8oVsXYNoESTGdZH0NYqJJnNF_HgguSsdAGOkk4"

def parse_text_with_style(text, font_size=16):
    elements = []
    bold = False
    yellow = False
    buffer = ""

    def flush():
        nonlocal buffer, bold, yellow
        if buffer:
            style = {"fontSize": {"magnitude": font_size, "unit": "PT"}}
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

def create_textbox(object_id, slide_id, x, y, width, height):
    return {
        "createShape": {
            "objectId": object_id,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": slide_id,
                "size": {"height": {"magnitude": height, "unit": "PT"}, "width": {"magnitude": width, "unit": "PT"}},
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": x, "translateY": y,
                    "unit": "PT"
                }
            }
        }
    }

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

        all_requests = delete_requests.copy()

        blokken = st.session_state.get("slides_data", [])

        for slide_index, blok in enumerate(blokken, start=1):
            slide_id = f"slide_{slide_index}"

            all_requests.append({
                "createSlide": {
                    "objectId": slide_id,
                    "insertionIndex": slide_index,
                    "slideLayoutReference": {"predefinedLayout": "BLANK"}
                }
            })

            # Titel/datum bovenaan
            datum_id = f"datum_{slide_index}"
            all_requests.append(create_textbox(datum_id, slide_id, 50, 20, 600, 50))
            all_requests.append({
                "insertText": {
                    "objectId": datum_id,
                    "insertionIndex": 0,
                    "text": blok["title"]
                }
            })
            all_requests.append({
                "updateTextStyle": {
                    "objectId": datum_id,
                    "textRange": {"type": "ALL"},
                    "style": {"fontSize": {"magnitude": 20, "unit": "PT"}, "bold": True},
                    "fields": "fontSize,bold"
                }
            })

            groepen = blok["content"].split("----") if isinstance(blok["content"], str) else [blok["content"]]
            font_size = 16 if len(groepen) <= 2 else 14 if len(groepen) == 3 else 12

            for j, groep in enumerate(groepen):
                content_id = f"groep_{slide_index}_{j}"
                x_offset = 50 + j * 160
                text_elements = parse_text_with_style(groep.strip(), font_size)

                all_requests.append(create_textbox(content_id, slide_id, x_offset, 80, 150, 400))
                all_requests.append({
                    "insertText": {
                        "objectId": content_id,
                        "insertionIndex": 0,
                        "text": "".join([e["textRun"]["content"] for e in text_elements])
                    }
                })

                index = 0
                for e in text_elements:
                    length = len(e["textRun"]["content"])
                    all_requests.append({
                        "updateTextStyle": {
                            "objectId": content_id,
                            "textRange": {"type": "FIXED_RANGE", "startIndex": index, "endIndex": index + length},
                            "style": e["textRun"]["style"],
                            "fields": ",".join(e["textRun"]["style"].keys())
                        }
                    })
                    index += length

            ondertekst = st.session_state.get("ondertekst", "").strip()
            if ondertekst:
                onder_id = f"onder_{slide_index}"
                geel = st.session_state.get("geel")
                vet = st.session_state.get("vet")
                styled_text = f"{'[GEEL]' if geel else ''}{'**' if vet else ''}{ondertekst}{'**' if vet else ''}{'[/GEEL]' if geel else ''}"

                all_requests.append(create_textbox(onder_id, slide_id, 50, 500, 600, 50))
                styled_elements = parse_text_with_style(styled_text, 14)

                all_requests.append({
                    "insertText": {
                        "objectId": onder_id,
                        "insertionIndex": 0,
                        "text": "".join([e["textRun"]["content"] for e in styled_elements])
                    }
                })

                index = 0
                for e in styled_elements:
                    length = len(e["textRun"]["content"])
                    all_requests.append({
                        "updateTextStyle": {
                            "objectId": onder_id,
                            "textRange": {"type": "FIXED_RANGE", "startIndex": index, "endIndex": index + length},
                            "style": e["textRun"]["style"],
                            "fields": ",".join(e["textRun"]["style"].keys())
                        }
                    })
                    index += length

        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": all_requests}
        ).execute()

        st.success("✅ Succesvol geüpload naar Google Slides!")

    except HttpError as error:
        st.error(f"Fout bij Google Slides API: {error}")
    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
