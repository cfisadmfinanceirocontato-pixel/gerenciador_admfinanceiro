import streamlit as st
import pandas as pd

st.set_page_config(page_title="Cfis App Financeiro")

st.markdown("""
# Boas-vindas ao Cfis App Financeiro!

## O app foi desenvolvido para auxiliar no controle financeiro.
            
""")

st.sidebar.header("🏤 Cfis App Financeiro")


file_uplpad = st.file_uploader("Faça o upload do seu arquivo CSV para começar a usar o app.", type=["csv"])
if file_uplpad is not None:
    df = pd.read_csv(file_uplpad)
    st.dataframe(df, hide_index=True)

