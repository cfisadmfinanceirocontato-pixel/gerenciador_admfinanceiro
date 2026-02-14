import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

st.set_page_config(page_title="Pagto Diárias", page_icon="📋", layout="wide", initial_sidebar_state="expanded")

st.title("📋 Pagamento de Diárias")
st.markdown("---")

def formatar_data(data_obj):
    if pd.isna(data_obj) or data_obj == '':
        return datetime.now().strftime("%d/%m/%Y")
    try:
        return pd.to_datetime(data_obj, dayfirst=True).strftime("%d/%m/%Y")
    except:
        return str(data_obj)

# ✅ FORMATAÇÃO MOEDA REAL BR
def formatar_moeda(valor):
    try:
        valor_num = float(valor)
        return f"R$ {valor_num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"

@st.cache_data
def carregar_dados():
    try:
        dados_diarias = pd.read_csv("diarias_data.csv")
        if 'Data' in dados_diarias.columns:
            dados_diarias['Data'] = dados_diarias['Data'].apply(formatar_data)
    except:
        dados_diarias = pd.DataFrame(columns=[
            'Termo', 'Instrumento', 'Numero_Termo', 'Funcionario', 'CPF', 'Cargo', 
            'Qtd', 'Valor', 'Objetivo', 'Localidades', 'Periodo', 'Oficio', 'Data'
        ])
    
    dados_apoio = pd.DataFrame({
        'Numero_Termo': ['001/2024', '001/2024', '002/2024', '002/2024'],
        'Termo': ['TERMO1', 'TERMO1', 'TERMO2', 'TERMO2'],
        'Instrumento': ['INST 001/2024', 'INST 001/2024', 'INST 002/2024', 'INST 002/2024'],
        'Funcionario': ['João Silva', 'Maria Santos', 'Pedro Oliveira', 'Ana Costa'],
        'CPF': ['123.456.789-00', '987.654.321-00', '111.222.333-44', '555.666.777-88'],
        'Cargo': ['Analista', 'Técnica', 'Coordenador', 'Assistente']
    })
    return dados_apoio, dados_diarias

def salvar_dados(dados_diarias):
    if 'Data' in dados_diarias.columns:
        dados_diarias['Data'] = dados_diarias['Data'].apply(formatar_data)
    dados_diarias.to_csv("diarias_data.csv", index=False)

dados_apoio, dados_diarias = carregar_dados()

# Lista suspensa quantidade
opcoes_quantidade = ['0,0', '0,5', '1,5', '2,5', '3,5', '4,5']

# Sidebar
with st.sidebar:
    st.header("🔍 Filtros")
    termos_unicos = sorted(dados_apoio['Termo'].dropna().unique())
    termo_selecionado = st.selectbox("Termo:", [''] + termos_unicos)
    
    funcionarios = []
    if termo_selecionado:
        mask = dados_apoio['Termo'] == termo_selecionado
        funcionarios = sorted(dados_apoio.loc[mask, 'Funcionario'].dropna().unique())
    funcionario_selecionado = st.selectbox("Funcionário:", [''] + funcionarios)

# Interface principal
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📝 Novo Registro")
    
    st.markdown("**📋 Dados do Termo de Colaboração**")
    termo_input = st.text_input("📄 **Termo de Colaboração**:")
    
    # Instrumento | Nº Termo lado a lado
    col_instrumento_numero = st.columns([1, 1])
    with col_instrumento_numero[0]:
        instrumento = st.text_input("🎯 **Instrumento**:", key="instrumento")
    with col_instrumento_numero[1]:
        numero_termo = st.text_input("📍 **Nº Termo**:", key="numero_termo")
    
    # Funcionário → CPF | Cargo lado a lado
    st.markdown("**👤 Dados do Funcionário**")
    funcionario_input = st.text_input("👤 **Funcionário**:", value=funcionario_selecionado)
    
    col_cpf_cargo = st.columns([1, 1])
    with col_cpf_cargo[0]:
        cpf = st.text_input("🆔 **CPF**:")
    with col_cpf_cargo[1]:
        cargo = st.text_input("💼 **Cargo**:")
    
    # Auto-preenchimento
    if funcionario_selecionado:
        mask_func = dados_apoio['Funcionario'] == funcionario_selecionado
        if mask_func.any():
            cpf_default = dados_apoio.loc[mask_func, 'CPF'].iloc[0]
            cargo_default = dados_apoio.loc[mask_func, 'Cargo'].iloc[0]
            st.info(f"💡 CPF: {cpf_default} | Cargo: {cargo_default}")
    
    # ✅ QUANTIDADE | VALOR NÃO EDITÁVEL | DATA LADO A LADO
    st.markdown("**💰 Valores e Data**")
    col_qtd_valor_data = st.columns([1, 1, 1])
    
    with col_qtd_valor_data[0]:
        qtd = st.selectbox("🔢 **Quantidade**:", 
                          options=opcoes_quantidade, 
                          index=2, 
                          key="qtd_select")
    
    with col_qtd_valor_data[1]:
        # ✅ VALOR AUTOMÁTICO NÃO EDITÁVEL = Quantidade × 140
        qtd_num = float(qtd.replace(',', '.'))
        valor_auto = qtd_num * 140
        valor_formatado = formatar_moeda(valor_auto)
        
        # ✅ TEXTBOX NÃO EDITÁVEL com disabled=True
        st.text_input("💰 **Valor (R$)**:", 
                     value=valor_formatado,
                     disabled=True,
                     help="Valor automático: Quantidade × R$ 140,00")
        valor = valor_formatado  # Para salvar
    
    with col_qtd_valor_data[2]:
        data_input = st.text_input("📅 **Data (DD/MM/AAAA)**:", 
                                  value=datetime.now().strftime("%d/%m/%Y"),
                                  placeholder="14/02/2026")
    
    data_formatada = formatar_data(data_input)
    
    st.markdown("**📋 Detalhes da Viagem**")
    objetivo = st.text_area("🎯 **Objetivo**:", height=50)
    localidades = st.text_area("📍 **Localidades**:", height=50)
    periodo = st.text_input("📊 **Período**:")
    oficio = st.text_input("📋 **Ofício**:")
    
    if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
        novo_registro = pd.DataFrame([{
            'Termo': termo_input,
            'Instrumento': instrumento,
            'Numero_Termo': numero_termo,
            'Funcionario': funcionario_input,
            'CPF': cpf,
            'Cargo': cargo,
            'Qtd': qtd,
            'Valor': valor,  # ✅ Valor calculado automaticamente
            'Objetivo': objetivo,
            'Localidades': localidades,
            'Periodo': periodo,
            'Oficio': oficio,
            'Data': data_formatada
        }])
        
        dados_diarias = pd.concat([dados_diarias, novo_registro], ignore_index=True)
        salvar_dados(dados_diarias)
        st.success("✅ Registro salvo!")
        st.rerun()

with col2:
    st.subheader("⚡ Ações")
    if st.button("🔄 Atualizar"): st.rerun()
    if st.button("🗑️ Limpar"): st.rerun()
    
    st.markdown("---")
    st.subheader("📊 Estatísticas")
    if not dados_diarias.empty:
        valores = [float(v.replace('R$', '').replace('.', '').replace(',', '.')) for v in dados_diarias['Valor'] if v]
        st.metric("Total Registros", len(dados_diarias))
        st.metric("Total Valor", formatar_moeda(sum(valores)))

# Tabela
st.markdown("---")
st.subheader("📋 Registros")

if not dados_diarias.empty:
    colunas_ordenadas = ['Termo', 'Instrumento', 'Numero_Termo', 'Funcionario', 'CPF', 'Cargo', 'Qtd', 'Valor', 'Data']
    colunas_display = [col for col in colunas_ordenadas if col in dados_diarias.columns]
    
    df_display = dados_diarias[colunas_display].copy()
    if 'Data' in df_display.columns:
        df_display['Data'] = df_display['Data'].apply(formatar_data)
    
    edited_df = st.data_editor(
        df_display,
        num_rows="dynamic",
        use_container_width=True
    )
    
    col_actions1, col_actions2, col_actions3 = st.columns(3)
    with col_actions1:
        csv = io.BytesIO()
        edited_df.to_csv(csv, index=False)
        csv.seek(0)
        st.download_button("📥 Excel", csv, f"diarias_{datetime.now().strftime('%d%m%Y_%H%M')}.csv", "text/csv")
    
    with col_actions2:
        if st.button("💾 Salvar"):
            salvar_dados(edited_df)
            st.success("✅ Salvo!")
            st.rerun()
    
    with col_actions3:
        if st.button("🗑️ Limpar Tudo", type="secondary"):
            dados_diarias = pd.DataFrame(columns=dados_diarias.columns)
            salvar_dados(dados_diarias)
            st.rerun()

else:
    st.info("👆 Cadastre o primeiro registro!")

st.markdown("---")
st.caption("✅ Valor AUTO NÃO EDITÁVEL: Quantidade × R$ 140")
