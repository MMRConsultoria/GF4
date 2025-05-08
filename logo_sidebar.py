import streamlit as st

def mostrar_logo_cliente():
    st.markdown(
        """
        <style>
            /* Ajuste na sidebar para adicionar o logo no topo */
            [data-testid="stSidebar"] > div:first-child {
                padding-top: 100px;
                background-image: url("https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_cliente.png");
                background-repeat: no-repeat;
                background-position: top center;
                background-size: 80px;
            }
        </style>
        """,
        unsafe_allow_html=True
    )
