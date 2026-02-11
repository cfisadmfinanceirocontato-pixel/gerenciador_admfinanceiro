import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Cfis App Financeiro")

pg = st.navigation([
    #st.markdown("# Cfis App Financeiro"), 
    st.Page("pages/home.py", title="Home"),
    st.Page("pages/usuarios.py", title="Usuários"),
    st.Page("pages/pginstrumentos.py", title="Demonstrativo financeiro"),
    st.Page("pages/pgprovisionamento.py", title="Provisionamento mensal"),
    st.Page("pages/pgdiarias.py", title="Controle de Diárias"),
    st.Page("pages/pgpessoal.py", title="Pagamento de pessoal"),
    st.Page("pages/pgconsumo.py", title="Pagamento de consumo"),
    st.Page("pages/pgservicos.py", title="Pagamento de serviços"),
    st.Page("pages/pgfuncionarios.py", title="Cadastro de Funcionários"),
    st.Page("pages/pgfornecedores.py", title="Cadastro de Fornecedores"),
])

pg.run()




