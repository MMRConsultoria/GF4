import streamlit as st

def mostrar_logo_cliente():
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] > div:first-child {
                padding-top: 100px;
                background-image: url("https://raw.githubusercontent.com/MMRConsultoria/MMRConsultoria/principal/logo_cliente.png");
                background-repeat: no-repeat;
                background-position: top center;
                background-size: 80px;
            }
        </style>
        """,
        unsafe_allow_html=True
    )
