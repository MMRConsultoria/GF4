import streamlit as st

def mostrar_logo_cliente():
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] > div:first-child {
                display: flex;
                flex-direction: column;
                align-items: center;
            }
            .logo-container img {
                max-width: 100px;
                margin-bottom: 10px;
                margin-top: 5px;
            }
        </style>
        <div class="logo-container">
            <img src="https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_cliente.png" alt="Logo Cliente">
        </div>
        """,
        unsafe_allow_html=True
    )
