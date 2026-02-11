import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Cfis App Financeiro")

pg = st.navigation([
    #st.markdown("# Cfis App Financeiro"), 
    st.Page("home.py", title="Home"),
    st.Page("usuarios.py", title="Usuários"),
    st.Page("pginstrumentos.py", title="Demonstrativo financeiro"),
    st.Page("pgprovisionamento.py", title="Provisionamento mensal"),
    st.Page("pgdiarias.py", title="Controle de Diárias"),
    st.Page("pgpessoal.py", title="Pagamento de pessoal"),
    st.Page("pgconsumo.py", title="Pagamento de consumo"),
    st.Page("pgservicos.py", title="Pagamento de serviços"),
    st.Page("pgfuncionarios.py", title="Cadastro de Funcionários"),
    st.Page("pgfornecedores.py", title="Cadastro de Fornecedores"),
])

pg.run()




