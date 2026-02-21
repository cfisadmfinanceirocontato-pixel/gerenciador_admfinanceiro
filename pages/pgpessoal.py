import streamlit as st
import pandas as pd
import os
import glob
import io
from pathlib import Path

# Configuração da página
st.set_page_config(
    page_title="Importador Excel Completo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

@st.cache_data
def processar_arquivo_excel(arquivo):
    """Processa um arquivo Excel individual"""
    try:
        # Tenta aba 'data' primeiro (VBA original), depois primeira aba
        try:
            df = pd.read_excel(arquivo, sheet_name='data', engine='openpyxl')
        except:
            df = pd.read_excel(arquivo, engine='openpyxl')
        
        return df
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        return pd.DataFrame()

def ordenar_e_formatar(df):
    """Ordenação e formatação VBA"""
    df = df.copy()
    if 'Valor OBT' in df.columns:
        df = df.sort_values('Valor OBT', ascending=True).reset_index(drop=True)
    df.insert(0, 'ID', range(1, len(df) + 1))
    return df

def main():
    st.title("📊 Importador Excel Completo VBA → Streamlit")
    st.markdown("**Upload individual + Pasta múltipla + Visualização completa**")
    
    # Estado da sessão
    if 'df_total' not in st.session_state:
        st.session_state.df_total = pd.DataFrame()
    
    # ========== SIDEBAR ==========
    with st.sidebar:
        st.header("📁 OPÇÕES DE IMPORTAÇÃO")
        
        # 1. UPLOAD INDIVIDUAL
        st.subheader("1. Arquivo Individual")
        uploaded_files = st.file_uploader(
            "Escolha arquivo(s) .xlsx", 
            type=['xlsx'], 
            accept_multiple_files=True,
            key="upload_individual"
        )
        
        # 2. SELEÇÃO DE PASTA
        st.subheader("2. Pasta Completa")
        pasta_path = st.text_input(
            "Caminho da pasta:",
            placeholder="C:/Users/SeuNome/PastaExcel",
            key="pasta_input"
        )
        
        # Botão importar pasta
        if st.button("📂 IMPORTAR TODOS da Pasta", type="secondary"):
            if pasta_path and os.path.exists(pasta_path):
                arquivos = glob.glob(os.path.join(pasta_path, "*.xlsx"))
                if arquivos:
                    st.session_state.arquivos_pasta = arquivos
                    st.success(f"✅ {len(arquivos)} arquivos encontrados!")
                else:
                    st.error("Nenhum .xlsx na pasta!")
            else:
                st.error("Pasta inválida!")
    
    # ========== PROCESSAMENTO ==========
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Processar uploads individuais
        if uploaded_files:
            dfs = []
            for file in uploaded_files:
                with st.spinner(f"Lendo {file.name}..."):
                    df_temp = processar_arquivo_excel(file)
                    if not df_temp.empty:
                        dfs.append(df_temp)
                        st.info(f"✅ {file.name}: {len(df_temp)} linhas")
            
            if dfs:
                df_novo = pd.concat(dfs, ignore_index=True)
                st.session_state.df_total = pd.concat([
                    st.session_state.df_total, 
                    df_novo
                ], ignore_index=True)
                st.success(f"✅ {len(dfs)} arquivos importados!")
    
    with col2:
        # Processar pasta
        if 'arquivos_pasta' in st.session_state:
            if st.button("🚀 IMPORTAR PASTA INTEIRA", type="primary", use_container_width=True):
                progress_bar = st.progress(0)
                dfs = []
                
                for i, arquivo in enumerate(st.session_state.arquivos_pasta):
                    df_temp = processar_arquivo_excel(arquivo)
                    if not df_temp.empty:
                        dfs.append(df_temp)
                    progress_bar.progress((i + 1) / len(st.session_state.arquivos_pasta))
                
                if dfs:
                    df_novo = pd.concat(dfs, ignore_index=True)
                    st.session_state.df_total = pd.concat([
                        st.session_state.df_total, 
                        df_novo
                    ], ignore_index=True)
                    st.success(f"✅ Pasta importada: {len(dfs)} arquivos!")
    
    # ========== EXIBIÇÃO DOS DADOS ==========
    if not st.session_state.df_total.empty:
        st.markdown("---")
        st.subheader(f"📋 RESULTADO FINAL ({len(st.session_state.df_total):,})")
        
        # Aplicar formatação VBA
        df_exibicao = ordenar_e_formatar(st.session_state.df_total)
        
        # Dataframe com formatação
        st.dataframe(
            df_exibicao,
            use_container_width=True,
            column_config={
                "ID": st.column_config.NumberColumn("ID", width="small"),
                "Valor OBT": st.column_config.NumberColumn(
                    "Valor OBT", 
                    format="R$ %.2f",
                    width="medium"
                )
            },
            hide_index=True
        )
        
        # ========== DOWNLOADS ==========
        col1, col2, col3 = st.columns(3)
        with col1:
            csv = df_exibicao.to_csv(
                index=False, sep=';', decimal=',', encoding='utf-8-sig'
            )
            st.download_button(
                "📥 CSV (Excel BR)",
                csv,
                "dados_finais.csv",
                "text/csv"
            )
        
        with col2:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_exibicao.to_excel(writer, index=False, sheet_name='data')
            st.download_button(
                "📊 Excel",
                output.getvalue(),
                "dados_finais.xlsx",
                "application/vnd.openxmlformats-officerspreadsheetml.sheet"
            )
        
        with col3:
            if st.button("🗑️ LIMPAR TUDO", type="secondary"):
                st.session_state.df_total = pd.DataFrame()
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
    
    # ========== ESTATÍSTICAS ==========
    if not st.session_state.df_total.empty:
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📈 Total Linhas", f"{len(st.session_state.df_total):,}")
        with col2:
            st.metric("📁 Colunas", len(st.session_state.df_total.columns))
        with col3:
            st.metric("🎯 Ordenado por", "Valor OBT")
        with col4:
            st.metric("🆔 IDs", "1 a N")
    
    # ========== INSTRUÇÕES ==========
    with st.expander("📖 COMO USAR"):
        st.markdown("""
        **OPÇÃO 1 - Arquivos Individuais:**
        1. Sidebar → **Upload arquivo(s)**
        2. Selecione um ou mais .xlsx
        3. Dados aparecem automaticamente ✅
        
        **OPÇÃO 2 - Pasta Completa:**
        1. Sidebar → **Cole caminho da pasta**
        2. Clique **"IMPORTAR PASTA INTEIRA"**
        3. Todos os .xlsx são processados 🚀
        
        **✅ VBA 100% replicado:**
        - Aba "data" automática
        - Ordenação "Valor OBT" 
        - Coluna ID sequencial
        - CSV com ';' (Excel BR)
        """)

if __name__ == "__main__":
    main()
