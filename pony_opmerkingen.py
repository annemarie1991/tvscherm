import streamlit as st
import json
from pathlib import Path

st.set_page_config(page_title="Pony-opmerkingen", layout="wide")

OPMERKINGEN_PATH = Path("pony_opmerkingen.json")

def load_opmerkingen():
    try:
        if OPMERKINGEN_PATH.exists():
            with open(OPMERKINGEN_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        return {}
    return {}

def save_opmerkingen(data):
    with open(OPMERKINGEN_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

st.title("📝 Pony-opmerkingen beheren")

opmerkingen = load_opmerkingen()

# Nieuwe opmerking toevoegen
with st.form("toevoegen", clear_on_submit=True):
    pony = st.text_input("Pony-naam (of deel ervan)").strip()
    opmerking = st.text_input("Opmerking of eigenschap").strip()
    toevoegen = st.form_submit_button("Toevoegen")

    if toevoegen:
        if pony and opmerking:
            opmerkingen[pony] = opmerking
            save_opmerkingen(opmerkingen)
            st.success(f"Opmerking toegevoegd voor '{pony}'")
        else:
            st.warning("Vul zowel een pony-naam als een opmerking in.")

st.markdown("---")
st.subheader("📋 Bestaande opmerkingen")

if not opmerkingen:
    st.info("Nog geen opmerkingen toegevoegd.")
else:
    verwijder_key = None
    for pony, eigenschap in opmerkingen.items():
        col1, col2, col3 = st.columns([3, 6, 1])
        col1.markdown(f"**{pony}**")
        col2.markdown(eigenschap)
        if col3.button("🗑️", key=f"verwijder_{pony}"):
            verwijder_key = pony

    if verwijder_key:
        opmerkingen.pop(verwijder_key)
        save_opmerkingen(opmerkingen)
        st.experimental_rerun()
