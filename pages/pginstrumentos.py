# ---- IMPORTANDO AS BIBLIOTECAS ----
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
import re
from datetime import datetime
import os
import glob
import io
from pathlib import Path

# ============== PRIMEIRA ETAPA - IMPORTAR ARQUIVOS OFX ==============

def is_ofx_sgml(content):
    """Detecta se é SGML (OFX antigo) por tags como <OFX> sem <?xml>"""
    return b'<OFX>' in content and b'<?xml' not in content

def converter_data(data_str):
    if not data_str:
        return ""
    try:
        # OFX padrão: YYYYMMDDHHMMSS
        return datetime.strptime(data_str[:8], '%Y%m%d').strftime("%d/%m/%Y")
    except:
        return data_str[:8] if len(data_str) >= 8 else ""

def parse_ofx_sgml(file_content):
    """Parse manual SGML OFX (sem bibliotecas extras)"""
    content = file_content.read().decode('utf-8', errors='ignore')
    
    # Extrai seção BANKMSGSRSV1 > STMTTRNRS > STMTRS > BANKTRANLIST > STMTTRN
    stmttrn_pattern = r'<STMTTRN>(.*?)</STMTTRN>'
    matches = re.findall(stmttrn_pattern, content, re.DOTALL | re.IGNORECASE)
    
    dados = []
    acct_id = re.search(r'<ACCTID[^>]*>([^<]+)', content, re.IGNORECASE)
    acct_id = acct_id.group(1) if acct_id else 'N/A'
    
    for match in matches[:100]:  # Limita para performance
        trnamt = re.search(r'<TRNAMT[^>]*>([^<]+)', match, re.IGNORECASE)
        dtposted = re.search(r'<DTPOSTED[^>]*>([^<]+)', match, re.IGNORECASE)
        memo = re.search(r'<MEMO[^>]*>([^<]+)', match, re.IGNORECASE)
        checknum = re.search(r'<CHECKNUM[^>]*>([^<]+)', match, re.IGNORECASE)
        trntype = re.search(r'<TRNTYPE[^>]*>([^<]+)', match, re.IGNORECASE)
        
        amount = float((trnamt.group(1).replace(',', '.') if trnamt else '0'))
        
        dados.append({
            "BankID": "Seu_BankID",
            "AccountID": acct_id,
            "Date": converter_data(dtposted.group(1) if dtposted else ''),
            "CheckNum": checknum.group(1) if checknum else '',
            "Memo": memo.group(1) if memo else '',
            "Amount": amount,
            "TransactionType": (trntype.group(1)[:1] if trntype else '')
        })
    
    return pd.DataFrame(dados)

def parse_ofx_xml(file_content):
    """Parse XML OFX (ElementTree)"""
    try:
        tree = ET.parse(file_content)
        root = tree.getroot()
        acct_id = root.findtext('.//ACCTID', 'N/A')
        
        dados = []
        for trn in root.findall('.//STMTTRN'):
            dados.append({
                "BankID": "Seu_BankID",
                "AccountID": acct_id,
                "Date": converter_data(trn.findtext('.//DTPOSTED')),
                "CheckNum": trn.findtext('.//CHECKNUM', ''),
                "Memo": trn.findtext('.//MEMO', ''),
                "Amount": float((trn.findtext('.//TRNAMT', '0').replace(',', '.'))),
                "TransactionType": trn.findtext('.//TRNTYPE', '')[:1]
            })
        return pd.DataFrame(dados)
    except ET.ParseError:
        return None

def processar_arquivo_ofx(ofx_file):
    content = BytesIO(ofx_file.read())
    
    # Detecta formato e parse
    content.seek(0)
    if is_ofx_sgml(content.read(1024)):  # SGML detectado
        content.seek(0)
        return parse_ofx_sgml(content)
    else:
        content.seek(0)
        return parse_ofx_xml(content)

def main():
    st.title("🚀 Extrator OFX Universal")
    
    ofx_file = st.file_uploader("📁 Upload OFX", type=["ofx", "txt"])
    
    if ofx_file is not None:
        with st.spinner("🔄 Analisando formato OFX..."):
            df = processar_arquivo_ofx(ofx_file)
            
            if df is not None and not df.empty:
                st.success(f"✅ {len(df)} transações extraídas!")
                st.subheader("📋 Dados")
                st.dataframe(df, use_container_width=True)
                
                # Download
                csv = df.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
                st.download_button("💾 Baixar CSV (Excel BR)", csv, "extrato_ofx.csv", "text/csv")
                
                st.info("Formato: DD/MM/YYYY | Sep: ';' | Decimal: ','")
            else:
                st.warning("⚠️ Arquivo vazio ou formato não reconhecido. Envie as primeiras linhas para debug.")

if __name__ == "__main__":
    main()
 
 # ============== SEGUNDA ETAPA - IMPORTAR RELATÓRIOS DE OBT'S ==============

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