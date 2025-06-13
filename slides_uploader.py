import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from uuid import uuid4
import textwrap

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

        # Verwijder bestaande slides behalve de eerste
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slide_ids = [slide["objectId"] for slide in presentation.get("slides", [])[1:]]
        delete_requests = [{"deleteObject": {"objectId": sid}} for sid in slide_ids]

        all_requests = delete_requests.copy()
        blokken = st.session_state.get("slides_data", [])
        max_columns = 3

        for blok in blokken:
            lines = blok["content"].split("\n")
            title = blok["title"]

            # Deel in subslides van max_columns groepen per slide
            current_slide_lines = []
            group_count = 0

            for line in lines:
                if line.strip().startswith("Groep"):
                    if group_count >= max_columns:
                        # Voeg vorige slide toe
                        all_requests += build_slide_requests(title, current_slide_lines)
                        current_slide_lines = []
                        group_count = 0
                if line.strip():
                    current_slide_lines.append(line)
                    if line.strip().startswith("Groep"):
                        group_count += 1

            if current_slide_lines:
                all_requests += build_slide_requests(title, current_slide_lines)

        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": all_requests}
        ).execute()

        st.success("✅ Succesvol geüpload naar Google Slides!")

    except HttpError as error:
        st.error(f"Fout bij Google Slides API: {error}")
    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")

def build_slide_requests(title, content_lines):
    requests = []
    slide_id = f"slide_{uuid4().hex}"

    requests.append({
        "createSlide": {
            "objectId": slide_id,
            "insertionIndex": 1,
            "slideLayoutReference": {"predefinedLayout": "BLANK"}
        }
    })

    # Titel
    datum_id = f"datum_{uuid4().hex}"
    requests.append(create_textbox(datum_id, slide_id, 50, 20, 600, 50))
    requests.append({
        "insertText": {
            "objectId": datum_id,
            "insertionIndex": 0,
            "text": title
        }
    })
    requests.append({
        "updateTextStyle": {
            "objectId": datum_id,
            "textRange": {"type": "ALL"},
            "style": {"fontSize": {"magnitude": 20, "unit": "PT"}, "bold": True},
            "fields": "fontSize,bold"
        }
    })

    # Inhoud
    content_id = f"content_{uuid4().hex}"
    full_text = "\n".join(content_lines)
    line_count = full_text.count("\n") + 1
    font_size = 16 if line_count <= 10 else 14 if line_count <= 15 else 12

    requests.append(create_textbox(content_id, slide_id, 50, 80, 600, 400))
    text_elements = parse_text_with_style(full_text, font_size)
    flat_text = "".join([e["textRun"]["content"] for e in text_elements])
    requests.append({
        "insertText": {
            "objectId": content_id,
            "insertionIndex": 0,
            "text": flat_text
        }
    })

    index = 0
    for e in text_elements:
        length = len(e["textRun"]["content"])
        requests.append({
            "updateTextStyle": {
                "objectId": content_id,
                "textRange": {"type": "FIXED_RANGE", "startIndex": index, "endIndex": index + length},
                "style": e["textRun"]["style"],
                "fields": ",".join(e["textRun"]["style"].keys())
            }
        })
        index += length

    return requests
