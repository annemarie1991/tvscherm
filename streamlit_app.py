import streamlit as st
import pandas as pd

st.set_page_config(page_title="TV scherm maken - Het Zesspan", layout="wide")

st.title("🐴 Ponyplanner Het Zesspan")
st.write("Upload hieronder je Excel-bestand en kies het juiste tabblad om de planning te bekijken.")

uploaded_file = st.file_uploader("📤 Upload Excel-bestand", type=["xlsx"])

if uploaded_file:
    excel = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("📑 Kies een tabblad", excel.sheet_names)
    df = pd.read_excel(excel, sheet_name=sheet, header=None)

    st.subheader("📋 Voorbeeld van ruwe gegevens")
    st.dataframe(df.head(15))
    
    st.warning("Dit is een demo. De herkenning van pony’s, kinderen en juffen wordt in de volgende versie toegevoegd.")
else:
    st.info("Wacht op upload...")
