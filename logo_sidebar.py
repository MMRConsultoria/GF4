# logo_sidebar.py
import streamlit as st
from PIL import Image

def mostrar_logo_cliente():
    imagem = Image.open("logo_cliente.png")
    st.sidebar.image(imagem, use_column_width=True)
