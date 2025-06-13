import streamlit as st
import datetime

# Zorg dat deze app draait op een aparte pagina, bijvoorbeeld via Streamlit multipage of een sub-app
# Bijvoorbeeld: https://ponyplanner.streamlit.app/scherm

st.set_page_config(page_title="TV Scherm - Het Zesspan", layout="wide")
st.markdown("""
<style>
body, .stApp {
    background-color: white;
}
.scherm-container {
    max-width: 1000px;
    margin: auto;
    font-family: 'Arial', sans-serif;
}
.scherm-groep {
    margin-bottom: 2em;
    padding: 1em;
    border-bottom: 1px solid #ddd;
}
.scherm-datum {
    text-align: center;
    font-size: 28px;
    font-weight: bold;
    margin-bottom: 1em;
}
.scherm-titel {
    font-weight: bold;
    font-size: 18px;
    margin-top: 1em;
    margin-bottom: 0.3em;
}
.scherm-tekst {
    white-space: pre-wrap;
    font-size: 16px;
    margin-left: 0.5em;
}
.ondertekst {
    margin-top: 2em;
    text-align: center;
    padding: 0.7em;
    border-radius: 6px;
    font-size: 16px;
}
</style>
""", unsafe_allow_html=True)


# 🧠 Ophalen uit session_state
slides_data = st.session_state.get("slides_data", [])
ondertekst = st.session_state.get("ondertekst", "")
vet = st.session_state.get("vet", False)
geel = st.session_state.get("geel", False)
extra_notitie = st.session_state.get("extra_notitie", "")

# 📝 Bewerkbare tekst voor snelle aanvullingen
st.sidebar.header("Snelle wijziging op scherm")
st.session_state.extra_notitie = st.sidebar.text_area("Extra notitie of toevoeging (verschijnt onderaan)", value=extra_notitie)


# 📺 Schermweergave
with st.container():
    st.markdown("<div class='scherm-container'>", unsafe_allow_html=True)

    for blok in slides_data:
        st.markdown("<div class='scherm-groep'>", unsafe_allow_html=True)

        st.markdown(f"<div class='scherm-datum'>{blok['title']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='scherm-tekst'>{blok['content'].replace('-', '•')}</div>", unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

    # 🟨 Ondertekst
    if ondertekst:
        stijl = ""
        if geel:
            stijl += "background-color: yellow;"
        if vet:
            stijl += "font-weight: bold;"

        st.markdown(f"<div class='ondertekst' style='{stijl}'>{ondertekst}</div>", unsafe_allow_html=True)

    # 🟧 Extra notitie
    if extra_notitie:
        st.markdown("<hr>")
        st.markdown(f"<div class='scherm-tekst'><strong>🔔 Opmerking:</strong><br>{extra_notitie}</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)
