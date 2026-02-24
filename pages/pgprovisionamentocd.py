import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import os

# --- Configurações da página ---
st.set_page_config(
    page_title="Provisionamento mensal", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Estado da sessão ---
if 'df_editavel' not in st.session_state:
    st.session_state.df_editavel = pd.DataFrame()
    
if 'dados_formulario' not in st.session_state:
    st.session_state.dados_formulario = {}
    
if 'modo_edicao' not in st.session_state:
    st.session_state.modo_edicao = False
    
if 'linha_editando' not in st.session_state:
    st.session_state.linha_editando = None

# --- Funções para carregar e salvar dados dos instrumentos ---
@st.cache_data(ttl=60)
def carregar_dados_instrumentos():
    """Carrega os dados completos do arquivo CSV de instrumentos"""
    try:
        # Tentar diferentes caminhos para compatibilidade com Streamlit Cloud
        caminhos_possiveis = ['itens_instrumento.csv', './itens_instrumento.csv', 'data/itens_instrumento.csv']
        
        for caminho in caminhos_possiveis:
            if os.path.exists(caminho):
                df_instrumentos = pd.read_csv(caminho, dtype=str)
                # Garantir nomes de colunas padronizados
                df_instrumentos.columns = [col.strip().upper() for col in df_instrumentos.columns]
                return df_instrumentos
        
        st.error("Arquivo itens_instrumento.csv não encontrado. Verifique se o arquivo está no diretório correto.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao carregar itens_instrumento.csv: {e}")
        return pd.DataFrame()

def salvar_dados_instrumentos(df):
    """Salva os dados no arquivo itens_instrumento.csv"""
    try:
        caminhos_possiveis = ['itens_instrumento.csv', './itens_instrumento.csv', 'data/itens_instrumento.csv']
        
        for caminho in caminhos_possiveis:
            try:
                df.to_csv(caminho, index=False, encoding='utf-8-sig')
                st.success(f"Dados salvos com sucesso em {caminho}!")
                st.cache_data.clear()
                return True
            except:
                continue
        
        st.error("Não foi possível salvar o arquivo. Verifique as permissões.")
        return False
    except Exception as e:
        st.error(f"Erro ao salvar itens_instrumento.csv: {e}")
        return False

def carregar_dados_provisionamento():
    """Carrega dados do arquivo provisionamentocd.csv"""
    try:
        caminhos_possiveis = ['provisionamentocd.csv', './provisionamentocd.csv', 'data/provisionamentocd.csv']
        
        for caminho in caminhos_possiveis:
            if os.path.exists(caminho):
                df = pd.read_csv(caminho)
                return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao carregar provisionamentocd.csv: {e}")
        return pd.DataFrame()

def salvar_dados_provisionamento(df):
    """Salva dados no arquivo provisionamentocd.csv"""
    try:
        caminhos_possiveis = ['provisionamentocd.csv', './provisionamentocd.csv', 'data/provisionamentocd.csv']
        
        for caminho in caminhos_possiveis:
            try:
                df.to_csv(caminho, index=False, encoding='utf-8-sig')
                st.success("Dados salvos com sucesso em provisionamentocd.csv!")
                return True
            except:
                continue
                
        st.error("Não foi possível salvar o arquivo provisionamentocd.csv")
        return False
    except Exception as e:
        st.error(f"Erro ao salvar provisionamentocd.csv: {e}")
        return False

def obter_lista_instrumentos():
    """Obtém lista única de instrumentos da primeira coluna"""
    df = carregar_dados_instrumentos()
    if not df.empty and len(df.columns) > 0:
        primeira_coluna = df.columns[0]
        return df[primeira_coluna].dropna().unique().tolist()
    return []

def obter_dados_por_instrumento(instrumento):
    """Obtém todas as linhas relacionadas a um instrumento"""
    df = carregar_dados_instrumentos()
    if not df.empty and len(df.columns) > 0:
        primeira_coluna = df.columns[0]
        return df[df[primeira_coluna] == instrumento]
    return pd.DataFrame()

def obter_dados_por_item(instrumento, item):
    """Obtém dados específicos por instrumento e item"""
    df = carregar_dados_instrumentos()
    if not df.empty and len(df.columns) >= 5:
        primeira_coluna = df.columns[0]
        quarta_coluna = df.columns[3]
        return df[(df[primeira_coluna] == instrumento) & (df[quarta_coluna].astype(str) == str(item))]
    return pd.DataFrame()

def obter_valores_unicos_por_coluna(df, numero_coluna):
    """Obtém valores únicos de uma coluna específica do DataFrame"""
    if not df.empty and len(df.columns) > numero_coluna:
        return df.iloc[:, numero_coluna].dropna().unique().tolist()
    return []

# --- Função para formatar moeda BR ---
def formatar_moeda_br(valor):
    """Formata valor para Real Brasileiro"""
    try:
        if pd.isna(valor) or valor == "" or valor is None:
            return "R$ 0,00"
        # Se já for número, converte direto
        if isinstance(valor, (int, float)):
            return f"R$ {abs(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        # Se for string, tenta converter
        valor_str = str(valor).replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
        valor_float = float(valor_str)
        return f"R$ {abs(valor_float):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return f"R$ 0,00"

def converter_moeda_para_float(valor_str):
    """Converte string de moeda BR para float"""
    if isinstance(valor_str, (int, float)):
        return float(valor_str)
    try:
        if pd.isna(valor_str) or valor_str == "" or valor_str is None:
            return 0.0
        # Remove 'R$ ' e substitui vírgula por ponto
        valor_limpo = str(valor_str).replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
        return float(valor_limpo)
    except:
        return 0.0

def preparar_df_para_dashboard(df):
    """Prepara o DataFrame para o dashboard convertendo tipos corretamente"""
    if df.empty:
        return df
    
    df_dash = df.copy()
    
    # Converter coluna VALOR para float
    if 'VALOR' in df_dash.columns:
        df_dash['VALOR'] = df_dash['VALOR'].apply(converter_moeda_para_float)
    
    # Converter outras colunas numéricas se existirem
    for col in df_dash.columns:
        if col not in ['VALOR', 'NF', 'DATA PGTO']:
            try:
                # Tenta converter para float se possível
                df_dash[col] = pd.to_numeric(df_dash[col], errors='ignore')
            except:
                pass
    
    return df_dash

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
    """Dashboard com listas completas"""
    if df.empty:
        return None
    
    # Preparar DataFrame para o dashboard
    df_dash = preparar_df_para_dashboard(df)
    
    # Verificar se temos a coluna VALOR
    if 'VALOR' not in df_dash.columns:
        st.warning("Coluna 'VALOR' não encontrada no DataFrame")
        return None
    
    # Calcular totais
    total_faturado = df_dash['VALOR'].sum()
    
    # Calcular total pago baseado no status
    if 'STATUS' in df_dash.columns:
        total_pago = df_dash[df_dash['STATUS'] == 'PAGA']['VALOR'].sum()
    else:
        total_pago = 0
    
    total_a_faturar = total_faturado - total_pago
    
    # Agrupamentos
    if 'Nº DO TERMO' in df_dash.columns:
        termos_total = df_dash.groupby('Nº DO TERMO')['VALOR'].sum().round(2).sort_values(ascending=False)
    else:
        termos_total = pd.Series()
    
    if 'EMPRESA' in df_dash.columns:
        empresas_total = df_dash.groupby('EMPRESA')['VALOR'].sum().round(2).sort_values(ascending=False)
    else:
        empresas_total = pd.Series()
    
    return {
        'total_faturado': total_faturado,
        'total_pago': total_pago,
        'total_a_faturar': total_a_faturar,
        'termos_total': termos_total,
        'empresas_total': empresas_total,
        'col_valor': 'VALOR',
        'col_termo': 'Nº DO TERMO' if 'Nº DO TERMO' in df_dash.columns else 'Termo',
        'col_empresa': 'EMPRESA' if 'EMPRESA' in df_dash.columns else 'Empresa',
        'df_dash': df_dash
    }

# --- Título ---
st.title("📊 Provisionamento Custo Direto")

# --- DASHBOARD SUPERIOR ---
st.markdown("---")
st.subheader("📈 Resumo Provisionamento Custo Direto")

if not st.session_state.df_editavel.empty:
    dashboard_data = criar_dashboard(st.session_state.df_editavel)
    
    if dashboard_data:
        col1, col2, col3, col4 = st.columns(4)
        with col1: 
            st.metric("💰 Total Provisionado", formatar_moeda_br(dashboard_data['total_faturado']))
        with col2: 
            st.metric("💵 Total Pago", formatar_moeda_br(dashboard_data['total_pago']))
        with col3: 
            st.metric("🎯 A Faturar", formatar_moeda_br(dashboard_data['total_a_faturar']))
        with col4: 
            st.metric("📊 Registros", f"{len(st.session_state.df_editavel):,}")
        
        # Mostrar valores em debug
        with st.expander("📊 Valores calculados (debug)"):
            st.write(f"Total Faturado: R$ {dashboard_data['total_faturado']:.2f}")
            st.write(f"Total Pago: R$ {dashboard_data['total_pago']:.2f}")
            st.write(f"Total a Faturar: R$ {dashboard_data['total_a_faturar']:.2f}")
            
            # Mostrar primeiras linhas do DataFrame preparado
            st.write("Primeiras linhas do DataFrame (após conversão):")
            st.dataframe(dashboard_data['df_dash'][['VALOR', 'STATUS']].head() if 'STATUS' in dashboard_data['df_dash'].columns else dashboard_data['df_dash'][['VALOR']].head())
        
        col_l1, col_l2 = st.columns(2)
        
        with col_l1:
            if not dashboard_data['termos_total'].empty:
                st.markdown("**🏆 TOTAL PROVISIONADO POR TERMO**")
                st.caption(f"Coluna: {dashboard_data['col_termo']} | Valor: {dashboard_data['col_valor']}")
                
                df_termos = pd.DataFrame({
                    '🔸 TERMO': dashboard_data['termos_total'].index,
                    'TOTAL (R$)': dashboard_data['termos_total'].values
                })
                
                st.dataframe(  
                    df_termos,
                    use_container_width=True,
                    height=400,  
                    hide_index=True,
                    column_config={
                        'TOTAL (R$)': st.column_config.NumberColumn(format="R$ %.2f")
                    }
                )
            else:
                st.info("Nenhum dado de termo disponível")
        
        with col_l2:
            if not dashboard_data['empresas_total'].empty:
                st.markdown("**🏢 VALOR TOTAL POR EMPRESA**")
                st.caption(f"Coluna: {dashboard_data['col_empresa']} | Valor: {dashboard_data['col_valor']}")
                
                df_empresas = pd.DataFrame({
                    '👨‍💼 EMPRESA': dashboard_data['empresas_total'].index,
                    'TOTAL (R$)': dashboard_data['empresas_total'].values
                })
                
                st.dataframe(  
                    df_empresas,
                    use_container_width=True,
                    height=400,  
                    hide_index=True,
                    column_config={
                        'TOTAL (R$)': st.column_config.NumberColumn(format="R$ %.2f")
                    }
                )
            else:
                st.info("Nenhum dado de empresa disponível")
        
        st.markdown("---")
    else:
        st.warning("Não foi possível criar o dashboard. Verifique se há dados válidos.")
else:
    st.info("👆 Carregue dados para ativar dashboard")

# --- FORMULÁRIO DE CADASTRO ---
st.subheader("📝 Formulário de Provisionamento")

# Carregar lista de instrumentos
lista_instrumentos = obter_lista_instrumentos()
df_instrumentos = carregar_dados_instrumentos()

# Meses do ano
meses = [
    "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# Criar formulário
with st.form("formulario_provisionamento"):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Campo 01 - INSTRUMENTO (lista suspensa)
        instrumento = st.selectbox(
            "INSTRUMENTO *",
            options=[""] + lista_instrumentos if lista_instrumentos else [""],
            key="form_instrumento"
        )
    
    with col2:
        # Campo 02 - Nº DO TERMO (combobox baseado no instrumento)
        opcoes_termo = []
        if instrumento and instrumento in lista_instrumentos:
            df_filtrado = obter_dados_por_instrumento(instrumento)
            if not df_filtrado.empty and len(df_filtrado.columns) > 1:
                opcoes_termo = df_filtrado.iloc[:, 1].dropna().unique().tolist()
        
        numero_termo = st.selectbox(
            "Nº DO TERMO *",
            options=[""] + opcoes_termo if opcoes_termo else [""],
            key="form_termo"
        )
    
    with col3:
        # Campo 03 - TERMO DE COLABORAÇÃO (combobox baseado no instrumento)
        opcoes_termo_nome = []
        if instrumento and instrumento in lista_instrumentos:
            df_filtrado = obter_dados_por_instrumento(instrumento)
            if not df_filtrado.empty and len(df_filtrado.columns) > 2:
                opcoes_termo_nome = df_filtrado.iloc[:, 2].dropna().unique().tolist()
        
        termo_colaboracao = st.selectbox(
            "TERMO DE COLABORAÇÃO *",
            options=[""] + opcoes_termo_nome if opcoes_termo_nome else [""],
            key="form_termo_nome"
        )
    
    col4, col5 = st.columns(2)
    
    with col4:
        # Campo 04 - MÊS COMPETÊNCIA
        mes_competencia = st.selectbox(
            "MÊS COMPETÊNCIA *",
            options=meses,
            key="form_mes"
        )
    
    with col5:
        # Campo 05 - ANO COMPETÊNCIA
        ano_atual = datetime.now().year
        anos = [""] + [str(ano) for ano in range(ano_atual - 2, ano_atual + 3)]
        ano_competencia = st.selectbox(
            "ANO COMPETÊNCIA *",
            options=anos,
            key="form_ano"
        )
    
    col6, col7 = st.columns(2)
    
    with col6:
        # Campo 06 - ITEM (baseado no instrumento)
        itens_disponiveis = []
        if instrumento and instrumento in lista_instrumentos:
            df_filtrado = obter_dados_por_instrumento(instrumento)
            if not df_filtrado.empty and len(df_filtrado.columns) > 3:
                itens_disponiveis = df_filtrado.iloc[:, 3].dropna().unique().tolist()
        
        item = st.selectbox(
            "ITEM *",
            options=[""] + itens_disponiveis,
            key="form_item"
        )
    
    with col7:
        # Campo 07 - DESCRIÇÃO (combobox baseado no instrumento e item)
        opcoes_descricao = []
        if instrumento and item and instrumento in lista_instrumentos:
            df_filtrado = obter_dados_por_item(instrumento, item)
            if not df_filtrado.empty and len(df_filtrado.columns) > 4:
                opcoes_descricao = df_filtrado.iloc[:, 4].dropna().unique().tolist()
        
        descricao = st.selectbox(
            "DESCRIÇÃO *",
            options=[""] + opcoes_descricao if opcoes_descricao else [""],
            key="form_descricao"
        )
    
    col8, col9, col10 = st.columns(3)
    
    with col8:
        # Campo 08 - VALOR (formatado em moeda)
        valor_str = st.text_input(
            "VALOR (R$) *",
            value="R$ 0,00",
            key="form_valor"
        )
        # Converter para float
        valor_float = converter_moeda_para_float(valor_str)
    
    with col9:
        # Campo 09 - NF
        nf = st.text_input(
            "NF",
            value="",
            key="form_nf"
        )
    
    with col10:
        # Campo 10 - EMPRESA
        empresa = st.text_input(
            "EMPRESA *",
            value="",
            key="form_empresa"
        )
    
    col11, col12 = st.columns(2)
    
    with col11:
        # Campo 11 - STATUS
        status = st.selectbox(
            "STATUS",
            options=["", "FATURADA", "PAGA"],
            key="form_status"
        )
    
    with col12:
        # Campo 12 - DATA PGTO
        data_pgto = st.date_input(
            "DATA PGTO",
            value=None,
            min_value=None,
            max_value=None,
            format="DD/MM/YYYY",
            key="form_data_pgto"
        )
    
    # Botões do formulário
    col_btn1, col_btn2, col_btn3 = st.columns(3)
    
    with col_btn1:
        submit_button = st.form_submit_button("📋 PRÉ-VISUALIZAR", use_container_width=True)
    
    with col_btn2:
        if st.session_state.modo_edicao:
            save_button = st.form_submit_button("💾 SALVAR EDIÇÃO", type="primary", use_container_width=True)
        else:
            save_button = st.form_submit_button("✅ ADICIONAR", type="primary", use_container_width=True)
    
    with col_btn3:
        cancel_button = st.form_submit_button("❌ CANCELAR", use_container_width=True)

# --- PRÉ-VISUALIZAÇÃO DOS DADOS ---
if submit_button:
    st.subheader("📋 PRÉ-VISUALIZAÇÃO DOS DADOS")
    
    if not instrumento or not numero_termo or not termo_colaboracao or not mes_competencia or not ano_competencia or not item or not descricao or not empresa:
        st.error("Por favor, preencha todos os campos obrigatórios (*)")
    else:
        dados_preview = {
            "INSTRUMENTO": instrumento,
            "Nº DO TERMO": numero_termo,
            "TERMO DE COLABORAÇÃO": termo_colaboracao,
            "MÊS COMPETÊNCIA": mes_competencia,
            "ANO COMPETÊNCIA": ano_competencia,
            "ITEM": item,
            "DESCRIÇÃO": descricao,
            "VALOR": valor_float,
            "NF": nf,
            "EMPRESA": empresa,
            "STATUS": status,
            "DATA PGTO": data_pgto.strftime("%d/%m/%Y") if data_pgto else ""
        }
        
        df_preview = pd.DataFrame([dados_preview])
        st.dataframe(df_preview, use_container_width=True, hide_index=True)
        
        # Armazenar no session_state
        st.session_state.dados_formulario = dados_preview

# --- SALVAR DADOS DO FORMULÁRIO ---
if save_button:
    if not instrumento or not numero_termo or not termo_colaboracao or not mes_competencia or not ano_competencia or not item or not descricao or not empresa:
        st.error("Por favor, preencha todos os campos obrigatórios (*)")
    else:
        # Criar dicionário com os dados
        novo_registro = {
            "INSTRUMENTO": instrumento,
            "Nº DO TERMO": numero_termo,
            "TERMO DE COLABORAÇÃO": termo_colaboracao,
            "MÊS COMPETÊNCIA": mes_competencia,
            "ANO COMPETÊNCIA": ano_competencia,
            "ITEM": item,
            "DESCRIÇÃO": descricao,
            "VALOR": valor_float,
            "NF": nf,
            "EMPRESA": empresa,
            "STATUS": status,
            "DATA PGTO": data_pgto.strftime("%d/%m/%Y") if data_pgto else ""
        }
        
        if st.session_state.modo_edicao and st.session_state.linha_editando is not None:
            # Modo edição - atualizar linha existente
            df = st.session_state.df_editavel.copy()
            for col, valor in novo_registro.items():
                if col in df.columns:
                    df.at[st.session_state.linha_editando, col] = valor
            st.session_state.df_editavel = df
            st.success("Registro atualizado com sucesso!")
            st.session_state.modo_edicao = False
            st.session_state.linha_editando = None
        else:
            # Modo adição - nova linha
            if st.session_state.df_editavel.empty:
                st.session_state.df_editavel = pd.DataFrame([novo_registro])
            else:
                st.session_state.df_editavel = pd.concat([st.session_state.df_editavel, pd.DataFrame([novo_registro])], ignore_index=True)
            st.success("Registro adicionado com sucesso!")
        
        # Salvar no arquivo provisionamentocd.csv
        salvar_dados_provisionamento(st.session_state.df_editavel)
        st.rerun()

# --- CANCELAR EDIÇÃO ---
if cancel_button:
    st.session_state.modo_edicao = False
    st.session_state.linha_editando = None
    st.rerun()

st.markdown("---")

# --- Sidebar: Upload e Filtros ---
with st.sidebar:
    st.header("📁 CARREGAR DADOS")
    
    file_upload = st.file_uploader("Upload CSV", type=["csv"])
    
    if st.button("📂 provisionamentocd.csv"):
        df_local = carregar_dados_provisionamento()
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
                    ]
                    
                    # Aplicar ordenação
                    if coluna_filtro in df_filtrado.columns:
                        if ordem == "Maior → Menor":
                            df_filtrado = df_filtrado.sort_values(coluna_filtro, ascending=False)
                        elif ordem == "Menor → Maior":
                            df_filtrado = df_filtrado.sort_values(coluna_filtro, ascending=True)
                    
                    st.session_state.df_editavel = df_filtrado.reset_index(drop=True)
                    st.success("✅ Dashboard atualizado!")
                    st.rerun()

# --- Tabela Principal ---
st.subheader("📋 Tabela Completa")

if file_upload is not None:
    df_upload = pd.read_csv(file_upload)
    st.session_state.df_editavel = df_upload
    st.success("✅ Dashboard carregado!")
    st.rerun()

if not st.session_state.df_editavel.empty:
    # Manter tipos originais para edição
    df_display = st.session_state.df_editavel.copy()
    
    # Configuração das colunas
    config_colunas = {}
    for col in df_display.columns:
        if col == 'VALOR':
            config_colunas[col] = st.column_config.TextColumn(
                col,
                help="Valor em R$"
            )
        elif col == 'NF':
            config_colunas[col] = st.column_config.TextColumn(
                col,
                help="Número da Nota Fiscal"
            )
        elif col == 'DATA PGTO':
            config_colunas[col] = st.column_config.TextColumn(
                col,
                help="Data no formato dd/mm/aaaa"
            )
        else:
            config_colunas[col] = st.column_config.TextColumn(col)
    
    # Exibir tabela com data_editor
    edited_df = st.data_editor(
        df_display,
        column_config=config_colunas,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="data_editor_principal"
    )
    
    # Atualizar se houver mudanças
    if not edited_df.equals(st.session_state.df_editavel):
        st.session_state.df_editavel = edited_df
        st.rerun()
    
    # Botões de ação por linha
    st.subheader("🔄 Ações por Linha")
    col_acoes = st.columns(5)
    
    with col_acoes[0]:
        if st.button("✏️ Editar Linha", use_container_width=True):
            st.session_state.modo_edicao = True
            linhas = list(range(len(st.session_state.df_editavel)))
            linha_selecionada = st.selectbox("Selecione a linha para editar:", linhas, key="select_linha_editar")
            st.session_state.linha_editando = linha_selecionada
            
            if linha_selecionada is not None:
                linha_dados = st.session_state.df_editavel.iloc[linha_selecionada]
                st.info("Use o formulário acima para editar os dados desta linha.")
    
    with col_acoes[1]:
        if st.button("🗑️ Excluir Linha", use_container_width=True):
            linha_excluir = st.number_input("Nº da linha para excluir:", min_value=0, max_value=len(st.session_state.df_editavel)-1, step=1, key="input_excluir")
            if st.button("Confirmar Exclusão", key="btn_confirmar_exclusao"):
                st.session_state.df_editavel = st.session_state.df_editavel.drop(linha_excluir).reset_index(drop=True)
                salvar_dados_provisionamento(st.session_state.df_editavel)
                st.success("Linha excluída com sucesso!")
                st.rerun()
    
    with col_acoes[2]:
        if st.button("💾 Salvar no Arquivo", use_container_width=True):
            if salvar_dados_provisionamento(st.session_state.df_editavel):
                st.success("Dados salvos em provisionamentocd.csv!")

# --- EDIÇÃO DO ARQUIVO itens_instrumento.csv ---
st.markdown("---")
st.subheader("✏️ Editar Itens do Instrumento")

with st.expander("Gerenciar itens_instrumento.csv"):
    df_itens = carregar_dados_instrumentos()
    
    if not df_itens.empty:
        # Converter para string para edição
        df_itens_display = df_itens.astype(str)
        
        edited_itens = st.data_editor(
            df_itens_display,
            use_container_width=True,
            num_rows="dynamic",
            key="editor_itens"
        )
        
        col_itens1, col_itens2 = st.columns(2)
        
        with col_itens1:
            if st.button("💾 Salvar Alterações em itens_instrumento.csv", use_container_width=True):
                if salvar_dados_instrumentos(edited_itens):
                    st.success("Arquivo itens_instrumento.csv atualizado!")
                    st.rerun()
        
        with col_itens2:
            if st.button("➕ Adicionar Nova Linha", use_container_width=True):
                nova_linha = pd.DataFrame([["" for _ in df_itens.columns]], columns=df_itens.columns)
                df_atualizado = pd.concat([df_itens, nova_linha], ignore_index=True)
                st.session_state.editor_itens = df_atualizado
                st.rerun()
    else:
        st.warning("Arquivo itens_instrumento.csv não encontrado ou vazio")

# --- AÇÕES RÁPIDAS ---
st.markdown("---")
st.subheader("🛠️ AÇÕES RÁPIDAS")

col_acao1, col_acao2, col_acao3 = st.columns(3)

with col_acao1:
    if st.button("➕ ➕ Nova Linha", type="secondary", use_container_width=True):
        nova_linha = {}
        for col in st.session_state.df_editavel.columns:
            nova_linha[col] = ""
        st.session_state.df_editavel = pd.concat([st.session_state.df_editavel, pd.DataFrame([nova_linha])], ignore_index=True)
        st.rerun()

with col_acao2:
    if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
        if salvar_dados_provisionamento(st.session_state.df_editavel):
            st.success("Alterações salvas com sucesso!")

with col_acao3:
    if st.button("🗑️ Limpar Tudo", type="secondary", use_container_width=True):
        st.session_state.df_editavel = pd.DataFrame()
        st.rerun()

# --- Downloads ---
if not st.session_state.df_editavel.empty:
    st.subheader("💾 DOWNLOADS")
    
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        csv = st.session_state.df_editavel.to_csv(index=False, sep=';', encoding='utf-8-sig')
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
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                          use_container_width=True)