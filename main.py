import streamlit as st
from modules.homepage import render_home_page
from modules.splitter import render_splitter
from modules.verifier import render_verifier

def set_page_config():
    """Set the Streamlit page configuration."""
    st.set_page_config(page_title="Splitter App", page_icon="🔗", layout="wide")

    # CSS untuk menyembunyikan footer dan menu
    hide_streamlit_style = """
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        </style>
    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)

def main():
    set_page_config()
    menu = ["Home", "Splitter", "Verifier"]
    choice = st.sidebar.selectbox("Main Menu", menu)

    if choice == "Home":
        render_home_page()
    elif choice == "Splitter":
        render_splitter()
    elif choice == "Verifier":
        render_verifier()

if __name__ == "__main__":
    main()
