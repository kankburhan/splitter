import streamlit as st

def render_home_page():
    """Render the Home page."""
    import streamlit as st
    st.title("Home")
    st.write("Welcome to the Home page of the Splitter App!")

    footer_html = """<div style='text-align: center;'>
      <p>Developed with 🔥 by Puti Andini</p>
    </div>"""
    st.markdown(footer_html, unsafe_allow_html=True)