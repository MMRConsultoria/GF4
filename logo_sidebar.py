# logo_sidebar.py
from PIL import Image
import streamlit as st

def exibir_logo(path="logo_cliente.png"):
    try:
        logo = Image.open(path)
        st.sidebar.image(logo, use_column_width=True)
    except FileNotFoundError:
        st.sidebar.warning("⚠️ Logo não encontrado.")
