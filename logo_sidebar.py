import streamlit as st

def mostrar_logo_cliente():
    st.markdown(
        """
        <style>
            /* Aplica o logo no TOPO da sidebar, fixo e visível */
            [data-testid="stSidebar"]::before {
                content: "";
                display: block;
                height: 80px;
                background-image: url("https://raw.githubusercontent.com/MMRConsultoria/MMRConsultoria/principal/logo_cliente1.png");
                background-repeat: no-repeat;
                background-position: center top;
                background-size: 60px;
                margin: 20px 0;
            }
        </style>
        """,
        unsafe_allow_html=True
    )
