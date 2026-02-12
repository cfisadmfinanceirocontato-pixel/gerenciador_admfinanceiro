# --- IMportando as bibliotecas ---
import streamlit as st
import pandas as pd
import numpy as np

# --- Configurações da página ---
st.set_page_config(page_title="Dados gerais!", layout="wide")

# --- Título da página ---
st.title("Dados gerais dos termos de colaboração.")


# --- Pldanilha de dados gerais ----

df_instrumentos = pd.read_csv("lista_instrumentos.csv", low_memory=True)
st.write(df_instrumentos)


file_uplpad = st.file_uploader("Faça o upload do seu arquivo CSV para começar a usar o app.", type=["csv"])
if file_uplpad is not None:
    df = pd.read_csv(file_uplpad)
    st.dataframe(df, hide_index=True)

