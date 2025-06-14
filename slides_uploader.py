def upload_to_slides():
    if "slides_data" not in st.session_state or not st.session_state["slides_data"]:
        st.error("Geen slides om te uploaden.")
        return

    try:
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["google_service_account"], scopes=SCOPES
        )
        service = build('slides', 'v1', credentials=credentials)

        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        slides = presentation.get('slides', [])
        if len(slides) < 1:
            st.error("Geen slides gevonden in de presentatie.")
            return

        base_slide_id = slides[-1].get('objectId')  # laatste slide is sjabloon
        slide_ids = [slide['objectId'] for slide in slides[:-1]]  # alles behalve laatste

        delete_requests = [{"deleteObject": {"objectId": slide_id}} for slide_id in slide_ids]
        if delete_requests:
            service.presentations().batchUpdate(
                presentationId=PRESENTATION_ID,
                body={"requests": delete_requests}
            ).execute()

        # 🟢 omgekeerde volgorde: eerste blok als eerste slide
        blocks = list(reversed(st.session_state["slides_data"]))

        requests = []
        for blok in blocks:
            slide_id = f"slide_{uuid.uuid4().hex[:8]}"
            requests.append({
                "duplicateObject": {
                    "objectId": base_slide_id,
                    "objectIds": {base_slide_id: slide_id}
                }
            })

            # Bovenaan: datum in het midden en vetgedrukt
            datum_id = f"datum_{uuid.uuid4().hex[:8]}"
            requests.append({
                "createShape": {
                    "objectId": datum_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"height": {"magnitude": 30, "unit": "PT"},
                                 "width": {"magnitude": 600, "unit": "PT"}},
                        "transform": {
                            "scaleX": 1, "scaleY": 1,
                            "translateX": 50, "translateY": 5,
                            "unit": "PT"
                        }
                    }
                }
            })
            requests.append({
                "insertText": {
                    "objectId": datum_id,
                    "insertionIndex": 0,
                    "text": blok["title"].replace("Planning ", "")
                }
            })
            requests.append({
                "updateTextStyle": {
                    "objectId": datum_id,
                    "textRange": {"type": "ALL"},
                    "style": {"fontSize": {"magnitude": 14, "unit": "PT"}, "bold": True},
                    "fields": "fontSize,bold"
                }
            })
            requests.append({
                "updateParagraphStyle": {
                    "objectId": datum_id,
                    "textRange": {"type": "ALL"},
                    "style": {"alignment": "CENTER"},
                    "fields": "alignment"
                }
            })

            # Kolommen per groep
            x_offset = 50
            y_offset = 60
            column_width = 200
            for i, col in enumerate(blok["columns"]):
                box_x = x_offset + i * (column_width + 40)
                box_y = y_offset
                content = f"**{col['tijd']}**\n**Juf: {col['juf']}**\n\n"
                for kind, pony in col["kinderen"]:
                    content += f"{kind} – {pony}\n"

                textbox_id = f"textbox_{uuid.uuid4().hex[:8]}"
                requests.append({
                    "createShape": {
                        "objectId": textbox_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 300, "unit": "PT"},
                                     "width": {"magnitude": column_width, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": box_x, "translateY": box_y,
                                "unit": "PT"
                            }
                        }
                    }
                })

                parsed = parse_markdown_to_text_elements(content)
                full_text = "".join([e["textRun"]["content"] for e in parsed])
                requests.append({
                    "insertText": {
                        "objectId": textbox_id,
                        "insertionIndex": 0,
                        "text": full_text
                    }
                })

                index_start = 0
                for element in parsed:
                    length = len(element["textRun"]["content"])
                    if element["textRun"].get("style"):
                        requests.append({
                            "updateTextStyle": {
                                "objectId": textbox_id,
                                "textRange": {
                                    "type": "FIXED_RANGE",
                                    "startIndex": index_start,
                                    "endIndex": index_start + length
                                },
                                "style": element["textRun"]["style"],
                                "fields": ",".join(element["textRun"]["style"].keys())
                            }
                        })
                    index_start += length

            # Ondertekst onderaan
            if blok.get("ondertekst"):
                onder_id = f"onder_{uuid.uuid4().hex[:8]}"
                requests.append({
                    "createShape": {
                        "objectId": onder_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"height": {"magnitude": 80, "unit": "PT"},
                                     "width": {"magnitude": 600, "unit": "PT"}},
                            "transform": {
                                "scaleX": 1, "scaleY": 1,
                                "translateX": 50, "translateY": 330,
                                "unit": "PT"
                            }
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": onder_id,
                        "insertionIndex": 0,
                        "text": blok["ondertekst"]
                    }
                })

                style = {}
                if blok.get("vet"):
                    style["bold"] = True
                if blok.get("geel"):
                    style["foregroundColor"] = {
                        "opaqueColor": {"rgbColor": {"red": 1, "green": 0.84, "blue": 0}}
                    }

                if style:
                    requests.append({
                        "updateTextStyle": {
                            "objectId": onder_id,
                            "textRange": {"type": "ALL"},
                            "style": style,
                            "fields": ",".join(style.keys())
                        }
                    })

                requests.append({
                    "updateParagraphStyle": {
                        "objectId": onder_id,
                        "textRange": {"type": "ALL"},
                        "style": {"alignment": "CENTER"},
                        "fields": "alignment"
                    }
                })

        service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={"requests": requests}
        ).execute()

        st.success("Slides succesvol geüpload!")

    except Exception as e:
        st.error(f"Fout tijdens uploaden naar Slides: {e}")
