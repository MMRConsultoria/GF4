# logo_sidebar.py
import streamlit as st
from PIL import Image
import os

def mostrar_logo_cliente():
    caminho_logo = os.path.join(os.path.dirname(__file__), "logo_cliente.png")
    try:
        imagem = Image.open(caminho_logo)
        st.sidebar.image(imagem, width=120)  # Logo no topo
        st.sidebar.markdown("---")  # Separador visual abaixo
    except Exception as e:
        st.sidebar.warning(f"⚠️ Erro ao carregar logo: {e}")
