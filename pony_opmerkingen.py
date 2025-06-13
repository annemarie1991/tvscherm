import streamlit as st
import json
from pathlib import Path

st.set_page_config(page_title="Pony-opmerkingen", layout="wide")
st.title("🐴 Pony-opmerkingen beheren")

json_pad = Path("pony_opmerkingen.json")

if not json_pad.exists():
    json_pad.write_text("{}", encoding="utf-8")

with json_pad.open("r", encoding="utf-8") as f:
    opmerkingen = json.load(f)

ponynaam = st.text_input("Pony (bijv. Parel of pare)")
opmerking = st.text_input("Opmerking voor deze pony")

if st.button("➕ Opslaan of bijwerken"):
    opmerkingen[ponynaam.strip()] = opmerking.strip()
    with json_pad.open("w", encoding="utf-8") as f:
        json.dump(opmerkingen, f, ensure_ascii=False, indent=2)
    st.success("Opmerking opgeslagen!")

st.markdown("### 📋 Bestaande opmerkingen")
for pony, tekst in sorted(opmerkingen.items()):
    st.markdown(f"**{pony}**: {tekst}")
