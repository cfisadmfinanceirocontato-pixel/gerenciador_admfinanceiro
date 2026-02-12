import streamlit as st
import pandas as pd
import numpy as np

st.title("Esta é a minha página de cadastro de usuários!")

file_uplpad = st.file_uploader("Faça o upload do seu arquivo CSV para começar a usar o app.", type=["csv"])
if file_uplpad is not None:
    df = pd.read_csv(file_uplpad)
    st.dataframe(df, hide_index=True)