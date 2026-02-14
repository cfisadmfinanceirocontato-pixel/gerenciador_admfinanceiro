import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

# --- Configurações da página ---
st.set_page_config(
    page_title="Provisionamento mensal", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Estado da sessão ---
if 'df_editavel' not in st.session_state:
    st.session_state.df_editavel = pd.DataFrame()

# --- Funções auxiliares ---
@st.cache_data
def carregar_dados_csv(nome_arquivo):
    """Carrega dados do CSV local"""
    try:
        return pd.read_csv(nome_arquivo, low_memory=False)
    except:
        return pd.DataFrame()

def obter_valores_unicos(df, coluna):
    """Obtém valores únicos da coluna para dropdown"""
    if coluna in df.columns:
        valores = df[coluna].dropna().unique()
        valores_str = pd.Series(valores).astype(str).unique()
        return sorted(valores_str)
    return []

def aplicar_filtros(df, coluna, ordem, valor_filtro=""):
    """Aplica filtros e ordenações"""
    df_filtrado = df.copy()
    
    if valor_filtro and coluna in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado[coluna].astype(str).str.contains(valor_filtro, case=False, na=False)]
    
    if coluna in df_filtrado.columns:
        if ordem == "Maior → Menor":
            df_filtrado = df_filtrado.sort_values(coluna, ascending=False)
        elif ordem == "Menor → Maior":
            df_filtrado = df_filtrado.sort_values(coluna, ascending=True)
    
    return df_filtrado.reset_index(drop=True)

def criar_dashboard(df):
    """Dashboard com LISTAS COMPLETAS de todos os termos"""
    if df.empty:
        return None
    
    colunas_numericas = df.select_dtypes(include=[np.number]).columns.tolist()
    colunas_categoricas = df.select_dtypes(include=['object']).columns.tolist()
    
    if not colunas_numericas:
        return None
    
    col_valor = colunas_numericas[0]
    col_termo = colunas_categoricas[0] if colunas_categoricas else df.columns[0]
    col_fornecedor = colunas_categoricas[1] if len(colunas_categoricas) > 1 else col_termo
    
    termos_total = df.groupby(col_termo)[col_valor].sum().round(2).sort_values(ascending=False)
    fornecedores_total = df.groupby(col_fornecedor)[col_valor].sum().round(2).sort_values(ascending=False)
    
    total_faturado = df[col_valor].sum()
    total_pago = total_faturado * 0.7
    total_a_faturar = total_faturado - total_pago
    
    return {
        'total_faturado': total_faturado,
        'total_pago': total_pago,
        'total_a_faturar': total_a_faturar,
        'termos_total': termos_total,
        'fornecedores_total': fornecedores_total,
        'col_valor': col_valor,
        'col_termo': col_termo,
        'col_fornecedor': col_fornecedor
    }

# --- Título ---
st.title("📊 Provisionamento Mensal - Dashboard Completo")

# --- DASHBOARD SUPERIOR ---
st.markdown("---")
st.subheader("📈 DASHBOARD - Todos os Termos e Fornecedores")

if not st.session_state.df_editavel.empty:
    dashboard_data = criar_dashboard(st.session_state.df_editavel)
    
    if dashboard_data:
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("💰 Total Faturado", f"R$ {dashboard_data['total_faturado']:,.2f}")
        with col2: st.metric("💵 Total Pago", f"R$ {dashboard_data['total_pago']:,.2f}")
        with col3: st.metric("🎯 A Faturar", f"R$ {dashboard_data['total_a_faturar']:,.2f}")
        with col4: st.metric("📊 Registros", f"{len(st.session_state.df_editavel):,}")
        
        col_l1, col_l2 = st.columns(2)
        
        with col_l1:
            st.markdown("**🏆 TOTAL FATURADO POR TERMO**")
            df_termos = pd.DataFrame({
                'Termo': dashboard_data['termos_total'].index,
                'Total R$': dashboard_data['termos_total'].values
            })
            st.dataframe(df_termos, use_container_width=True, height=400, hide_index=True,
                        column_config={'Total R$': st.column_config.NumberColumn(format="R$ %.2f")})
        
        with col_l2:
            st.markdown("**🏢 TOTAL POR FORNECEDOR**")
            df_fornecedores = pd.DataFrame({
                'Fornecedor': dashboard_data['fornecedores_total'].index,
                'Total R$': dashboard_data['fornecedores_total'].values
            })
            st.dataframe(df_fornecedores, use_container_width=True, height=400, hide_index=True,
                        column_config={'Total R$': st.column_config.NumberColumn(format="R$ %.2f")})
        
        st.markdown("---")
else:
    st.info("👆 Carregue dados para ver dashboard")

# --- Sidebar: Upload e Filtros ---
with st.sidebar:
    st.header("📁 CARREGAR DADOS")
    
    file_upload = st.file_uploader("Upload CSV", type=["csv"])
    
    if st.button("📂 provisionamento_mensal_cd.csv"):
        df_local = carregar_dados_csv("provisionamento_mensal_cd.csv")
        if not df_local.empty:
            st.session_state.df_editavel = df_local
            st.success(f"✅ {len(df_local)} linhas!")
            st.rerun()
    
    st.header("🔍 FILTROS")
    colunas = st.session_state.df_editavel.columns.tolist() if not st.session_state.df_editavel.empty else []
    
    if colunas:
        coluna_filtro = st.selectbox("🎯 Coluna:", ["-- Todas --"] + colunas)
        if coluna_filtro != "-- Todas --":
            valores_unicos = obter_valores_unicos(st.session_state.df_editavel, coluna_filtro)
            valor_filtro = st.selectbox("📋 Selecionar:", ["-- Todos --"] + valores_unicos[:50])
            
            ordem = st.selectbox("📊 Ordem:", ["Maior → Menor", "Menor → Maior"])
            
            if st.button("🚀 FILTRAR", type="primary"):
                if valor_filtro != "-- Todos --":
                    df_filtrado = st.session_state.df_editavel[
                        st.session_state.df_editavel[coluna_filtro].astype(str).str.contains(valor_filtro, case=False, na=False)
                    ].sort_values(coluna_filtro, ascending=(ordem == "Menor → Maior"))
                    st.session_state.df_editavel = df_filtrado.reset_index(drop=True)
                    st.success("✅ Dashboard atualizado!")
                    st.rerun()

# --- Tabela Principal (SEM ações laterais) ---
st.subheader("📋 Tabela Completa")

if file_upload is not None:
    df_upload = pd.read_csv(file_upload)
    st.session_state.df_editavel = df_upload
    st.success("✅ Dados carregados!")
    st.rerun()

if not st.session_state.df_editavel.empty:
    # Configuração das colunas
    config_colunas = {}
    for col in st.session_state.df_editavel.columns:
        if st.session_state.df_editavel[col].dtype in ['int64', 'float64']:
            config_colunas[col] = st.column_config.NumberColumn(col, format="R$ %.2f")
        else:
            config_colunas[col] = st.column_config.TextColumn(col)
    
    # Tabela editável
    edited_df = st.data_editor(
        st.session_state.df_editavel,
        num_rows="dynamic",
        column_config=config_colunas,
        use_container_width=True,
        hide_index=True
    )
    
    # === AÇÕES RÁPIDAS MOVIDAS PARA BAIXO DA TABELA ===
    st.markdown("---")
    st.subheader("🛠️ AÇÕES RÁPIDAS")
    
    col_acao1, col_acao2, col_acao3 = st.columns(3)
    
    with col_acao1:
        if st.button("➕ ➕ Nova Linha", type="secondary", use_container_width=True):
            nova_linha = {col: 0 if st.session_state.df_editavel[col].dtype in ['int64', 'float64'] else "" 
                         for col in st.session_state.df_editavel.columns}
            st.session_state.df_editavel = pd.concat([st.session_state.df_editavel, pd.DataFrame([nova_linha])], ignore_index=True)
            st.success("✅ Nova linha adicionada!")
            st.rerun()
    
    with col_acao2:
        if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
            st.session_state.df_editavel = edited_df
            st.success("✅ Alterações salvas no dashboard!")
            st.rerun()
    
    with col_acao3:
        if st.button("🗑️ Limpar Tudo", type="secondary", use_container_width=True):
            st.session_state.df_editavel = pd.DataFrame()
            st.success("✅ Dados limpos!")
            st.rerun()
    
    st.markdown("---")
else:
    st.warning("⚠️ Carregue um arquivo CSV primeiro!")

# --- Downloads ---
if not st.session_state.df_editavel.empty:
    st.subheader("💾 DOWNLOADS")
    
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        csv = st.session_state.df_editavel.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
        st.download_button("📥 CSV Brasil", csv, "provisionamento.csv", "text/csv", use_container_width=True)
    
    with col_dl2:
        @st.cache_data
        def gerar_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Dados')
            return output.getvalue()
        
        st.download_button("📊 Excel", gerar_excel(st.session_state.df_editavel), 
                          "provisionamento.xlsx", 
                          "application/vnd.openxmlformats-officerspreadsheetml.sheet", 
                          use_container_width=True)

# --- Info ---
with st.expander("ℹ️ Como Usar"):
    st.markdown("""
    1. **Upload CSV** → Dashboard carrega automaticamente
    2. **Filtros** na sidebar atualizam listas
    3. **Editar tabela** → Clique "Salvar Alterações" abaixo
    4. **➕ Nova Linha** → Adiciona linha vazia
    5. **🗑️ Limpar** → Remove todos os dados
    
    **✅ Ações movidas para BAIXO da tabela!**
    """)
