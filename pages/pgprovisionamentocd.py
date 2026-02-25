import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import os
import re
import time

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
    
if 'filtros_aplicados' not in st.session_state:
    st.session_state.filtros_aplicados = {}
    
if 'df_original' not in st.session_state:
    st.session_state.df_original = pd.DataFrame()

if 'instrumento_selecionado' not in st.session_state:
    st.session_state.instrumento_selecionado = ""
    
if 'opcoes_termo' not in st.session_state:
    st.session_state.opcoes_termo = []
if 'opcoes_termo_nome' not in st.session_state:
    st.session_state.opcoes_termo_nome = []
if 'itens_disponiveis' not in st.session_state:
    st.session_state.itens_disponiveis = []
if 'opcoes_descricao' not in st.session_state:
    st.session_state.opcoes_descricao = []
    
if 'preview_trigger' not in st.session_state:
    st.session_state.preview_trigger = False
    
if 'item_selecionado' not in st.session_state:
    st.session_state.item_selecionado = ""

# --- Funções para carregar e salvar dados dos instrumentos ---
@st.cache_data(ttl=60)
def carregar_dados_instrumentos():
    """Carrega os dados completos do arquivo CSV de instrumentos"""
    try:
        # Tentar diferentes caminhos para compatibilidade com Streamlit Cloud
        caminhos_possiveis = ['itens_instrumento.csv', './itens_instrumento.csv', 'data/itens_instrumento.csv']
        
        for caminho in caminhos_possiveis:
            try:
                if os.path.exists(caminho):
                    df_instrumentos = pd.read_csv(caminho, dtype=str, encoding='utf-8-sig')
                    # Garantir nomes de colunas padronizados
                    df_instrumentos.columns = [col.strip().upper() for col in df_instrumentos.columns]
                    return df_instrumentos
            except:
                continue
        
        # Se não encontrar, criar DataFrame vazio
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao carregar itens_instrumento.csv: {e}")
        return pd.DataFrame()

def salvar_dados_instrumentos(df):
    """Salva os dados no arquivo itens_instrumento.csv"""
    try:
        # Tentar diferentes caminhos
        caminhos_possiveis = ['itens_instrumento.csv', './itens_instrumento.csv', 'data/itens_instrumento.csv']
        
        for caminho in caminhos_possiveis:
            try:
                df.to_csv(caminho, index=False, encoding='utf-8-sig')
                st.success(f"Dados salvos com sucesso em {caminho}!")
                # Limpar cache para recarregar dados atualizados
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
            try:
                if os.path.exists(caminho):
                    df = pd.read_csv(caminho, encoding='utf-8-sig')
                    return df
            except:
                continue
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

# --- Função para formatar moeda BR ---
def formatar_moeda_br(valor):
    """Formata valor para Real Brasileiro"""
    try:
        if pd.isna(valor) or valor == "" or valor is None:
            return "R$ 0,00"
        if isinstance(valor, (int, float)):
            return f"R$ {abs(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
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
        valor_limpo = str(valor_str).replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
        return float(valor_limpo)
    except:
        return 0.0

def preparar_df_para_dashboard(df):
    """Prepara o DataFrame para o dashboard convertendo tipos corretamente"""
    if df.empty:
        return df
    
    df_dash = df.copy()
    
    if 'VALOR' in df_dash.columns:
        df_dash['VALOR'] = df_dash['VALOR'].apply(converter_moeda_para_float)
    
    return df_dash

# --- Funções auxiliares ---
@st.cache_data
def carregar_dados_csv(nome_arquivo):
    """Carrega dados do CSV local"""
    try:
        return pd.read_csv(nome_arquivo, low_memory=False, encoding='utf-8-sig')
    except:
        return pd.DataFrame()

def obter_valores_unicos(df, coluna):
    """Obtém valores únicos da coluna para dropdown"""
    if coluna in df.columns:
        valores = df[coluna].dropna().unique()
        valores_str = pd.Series(valores).astype(str).unique()
        return sorted(valores_str)
    return []

def criar_dashboard(df):
    """Dashboard com listas completas"""
    if df.empty:
        return None
    
    df_dash = preparar_df_para_dashboard(df)
    
    if 'VALOR' not in df_dash.columns:
        return None
    
    total_faturado = df_dash['VALOR'].sum()
    
    if 'STATUS' in df_dash.columns:
        total_pago = df_dash[df_dash['STATUS'] == 'PAGA']['VALOR'].sum()
    else:
        total_pago = 0
    
    total_a_faturar = total_faturado - total_pago
    
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

# --- Função para atualizar campos baseado no instrumento selecionado ---
def atualizar_campos_por_instrumento(instrumento):
    """Atualiza os campos Nº DO TERMO, TERMO DE COLABORAÇÃO, ITEM e DESCRIÇÃO"""
    if instrumento and instrumento in obter_lista_instrumentos():
        df_filtrado = obter_dados_por_instrumento(instrumento)
        
        if not df_filtrado.empty:
            if len(df_filtrado.columns) > 1:
                st.session_state.opcoes_termo = df_filtrado.iloc[:, 1].dropna().unique().tolist()
            
            if len(df_filtrado.columns) > 2:
                st.session_state.opcoes_termo_nome = df_filtrado.iloc[:, 2].dropna().unique().tolist()
            
            if len(df_filtrado.columns) > 3:
                st.session_state.itens_disponiveis = df_filtrado.iloc[:, 3].dropna().unique().tolist()
            
            st.session_state.opcoes_descricao = []
    else:
        st.session_state.opcoes_termo = []
        st.session_state.opcoes_termo_nome = []
        st.session_state.itens_disponiveis = []
        st.session_state.opcoes_descricao = []

# --- Função para processar valor monetário (sem modificar widget diretamente) ---
def processar_valor_monetario(valor_input):
    """Processa o valor digitado e retorna o valor formatado e o float"""
    if not valor_input:
        return "R$ 0,00", 0.0
    
    # Extrair apenas números
    numeros = re.sub(r'[^0-9]', '', valor_input)
    if not numeros:
        return "R$ 0,00", 0.0
    
    # Converter para float
    valor_float = float(numeros) / 100
    
    # Formatar
    valor_formatado = f"R$ {valor_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    
    return valor_formatado, valor_float

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

# Verificar se instrumento mudou (fora do form)
if 'form_instrumento' in st.session_state:
    instrumento_atual = st.session_state.form_instrumento
    if instrumento_atual != st.session_state.get('instrumento_anterior', ''):
        atualizar_campos_por_instrumento(instrumento_atual)
        st.session_state.instrumento_anterior = instrumento_atual
        # Simular clique no botão PRÉ-VISUALIZAR
        st.session_state.preview_trigger = True
        st.rerun()

# Verificar se item mudou (fora do form)
if 'form_item' in st.session_state and 'form_instrumento' in st.session_state:
    item_atual = st.session_state.form_item
    instrumento_atual = st.session_state.form_instrumento
    if item_atual != st.session_state.get('item_anterior', '') and instrumento_atual:
        if item_atual:
            df_item = obter_dados_por_item(instrumento_atual, item_atual)
            if not df_item.empty and len(df_item.columns) > 4:
                st.session_state.opcoes_descricao = df_item.iloc[:, 4].dropna().unique().tolist()
            else:
                st.session_state.opcoes_descricao = []
        st.session_state.item_anterior = item_atual
        # Simular clique no botão PRÉ-VISUALIZAR
        st.session_state.preview_trigger = True
        st.rerun()

# Criar formulário
with st.form("formulario_provisionamento"):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Campo 01 - INSTRUMENTO
        instrumento = st.selectbox(
            "INSTRUMENTO *",
            options=[""] + lista_instrumentos if lista_instrumentos else [""],
            key="form_instrumento"
        )
    
    with col2:
        # Campo 02 - Nº DO TERMO
        numero_termo = st.selectbox(
            "Nº DO TERMO *",
            options=[""] + st.session_state.opcoes_termo if st.session_state.opcoes_termo else [""],
            key="form_termo"
        )
    
    with col3:
        # Campo 03 - TERMO DE COLABORAÇÃO
        termo_colaboracao = st.selectbox(
            "TERMO DE COLABORAÇÃO *",
            options=[""] + st.session_state.opcoes_termo_nome if st.session_state.opcoes_termo_nome else [""],
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
        # Campo 06 - ITEM
        item = st.selectbox(
            "ITEM *",
            options=[""] + st.session_state.itens_disponiveis if st.session_state.itens_disponiveis else [""],
            key="form_item"
        )
    
    with col7:
        # Campo 07 - DESCRIÇÃO
        descricao = st.selectbox(
            "DESCRIÇÃO *",
            options=[""] + st.session_state.opcoes_descricao if st.session_state.opcoes_descricao else [""],
            key="form_descricao"
        )
    
    col8, col9, col10 = st.columns(3)
    
    with col8:
        # Campo 08 - VALOR
        valor_input = st.text_input(
            "VALOR (R$) *",
            value="R$ 0,00",
            key="form_valor_input"
        )
        
        # Processar valor
        valor_formatado, valor_float = processar_valor_monetario(valor_input)
    
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

# --- PRÉ-VISUALIZAÇÃO AUTOMÁTICA (triggered por mudanças) ---
if st.session_state.get('preview_trigger', False) and not submit_button and not save_button and not cancel_button:
    # Limpar o trigger
    st.session_state.preview_trigger = False
    
    # Executar pré-visualização automática
    if instrumento and numero_termo and termo_colaboracao and mes_competencia and ano_competencia and item and descricao and empresa:
        st.subheader("📋 PRÉ-VISUALIZAÇÃO DOS DADOS (Automática)")
        
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

# --- PRÉ-VISUALIZAÇÃO MANUAL (botão) ---
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
            df = st.session_state.df_editavel.copy()
            for col, valor in novo_registro.items():
                if col in df.columns:
                    df.at[st.session_state.linha_editando, col] = valor
            st.session_state.df_editavel = df
            st.success("Registro atualizado com sucesso!")
            st.session_state.modo_edicao = False
            st.session_state.linha_editando = None
        else:
            if st.session_state.df_editavel.empty:
                st.session_state.df_editavel = pd.DataFrame([novo_registro])
            else:
                st.session_state.df_editavel = pd.concat([st.session_state.df_editavel, pd.DataFrame([novo_registro])], ignore_index=True)
            st.success("Registro adicionado com sucesso!")
        
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
            st.session_state.df_original = df_local.copy()
            st.success(f"✅ {len(df_local)} linhas!")
            st.rerun()
    
    st.header("🔍 FILTROS")
    colunas = st.session_state.df_editavel.columns.tolist() if not st.session_state.df_editavel.empty else []
    
    if colunas:
        coluna_filtro = st.selectbox("🎯 Coluna para filtrar:", ["-- Selecione --"] + colunas)
        
        if coluna_filtro != "-- Selecione --":
            valores_unicos = obter_valores_unicos(st.session_state.df_editavel, coluna_filtro)
            
            valores_selecionados = st.multiselect(
                f"📋 Valores em {coluna_filtro}:",
                options=valores_unicos[:100],
                default=[]
            )
            
            if st.button("🚀 APLICAR FILTRO", type="primary"):
                if valores_selecionados:
                    if st.session_state.df_original.empty:
                        st.session_state.df_original = st.session_state.df_editavel.copy()
                    
                    df_filtrado = st.session_state.df_editavel[
                        st.session_state.df_editavel[coluna_filtro].astype(str).isin(valores_selecionados)
                    ]
                    
                    st.session_state.df_editavel = df_filtrado.reset_index(drop=True)
                    st.session_state.filtros_aplicados[coluna_filtro] = valores_selecionados
                    st.success(f"✅ Filtro aplicado em {coluna_filtro}!")
                    st.rerun()
        
        st.markdown("---")
        if st.button("🧹 LIMPAR TODOS OS FILTROS", type="secondary", use_container_width=True):
            if not st.session_state.df_original.empty:
                st.session_state.df_editavel = st.session_state.df_original.copy()
                st.session_state.filtros_aplicados = {}
                st.success("✅ Todos os filtros foram removidos!")
                st.rerun()
            else:
                st.warning("Nenhum filtro para limpar")

# --- Tabela Principal ---
st.subheader("📋 Tabela Completa")

if file_upload is not None:
    try:
        df_upload = pd.read_csv(file_upload, encoding='utf-8-sig')
        for col in df_upload.columns:
            df_upload[col] = df_upload[col].astype(str)
        st.session_state.df_editavel = df_upload
        st.session_state.df_original = df_upload.copy()
        st.success("✅ Dashboard carregado!")
        st.rerun()
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {e}")

if not st.session_state.df_editavel.empty:
    df_display = st.session_state.df_editavel.copy()
    
    for col in df_display.columns:
        df_display[col] = df_display[col].astype(str)
    
    config_colunas = {}
    for col in df_display.columns:
        config_colunas[col] = st.column_config.TextColumn(
            col,
            help=f"Campo {col}",
            width="medium"
        )
    
    try:
        coluna_para_filtrar = st.selectbox(
            "🔍 Filtrar tabela por coluna:",
            ["Selecione uma coluna"] + list(df_display.columns),
            key="filtro_tabela_coluna"
        )
        
        if coluna_para_filtrar != "Selecione uma coluna":
            valor_busca = st.text_input(f"Buscar em {coluna_para_filtrar}:", key="filtro_tabela_valor")
            if valor_busca:
                df_display = df_display[df_display[coluna_para_filtrar].str.contains(valor_busca, case=False, na=False)]
        
        edited_df = st.data_editor(
            df_display,
            column_config=config_colunas,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="data_editor_principal"
        )
        
        if not edited_df.equals(df_display):
            st.session_state.df_editavel = edited_df
            st.rerun()
    except Exception as e:
        st.error(f"Erro ao exibir tabela: {e}")
        st.dataframe(df_display, use_container_width=True)
    
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
            if len(st.session_state.df_editavel) > 0:
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
        df_itens_display = df_itens.astype(str)
        
        try:
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
        except Exception as e:
            st.error(f"Erro ao editar itens: {e}")
            st.dataframe(df_itens_display, use_container_width=True)
    else:
        st.warning("Arquivo itens_instrumento.csv não encontrado ou vazio")

# --- AÇÕES RÁPIDAS ---
st.markdown("---")
st.subheader("🛠️ AÇÕES RÁPIDAS")

col_acao1, col_acao2, col_acao3 = st.columns(3)

with col_acao1:
    if st.button("➕ ➕ Nova Linha", type="secondary", use_container_width=True):
        if not st.session_state.df_editavel.empty:
            nova_linha = {}
            for col in st.session_state.df_editavel.columns:
                nova_linha[col] = ""
            st.session_state.df_editavel = pd.concat([st.session_state.df_editavel, pd.DataFrame([nova_linha])], ignore_index=True)
            st.rerun()
        else:
            st.warning("Carregue dados primeiro")

with col_acao2:
    if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
        if not st.session_state.df_editavel.empty:
            if salvar_dados_provisionamento(st.session_state.df_editavel):
                st.success("Alterações salvas com sucesso!")
        else:
            st.warning("Não há dados para salvar")

with col_acao3:
    if st.button("🗑️ Limpar Tudo", type="secondary", use_container_width=True):
        st.session_state.df_editavel = pd.DataFrame()
        st.session_state.df_original = pd.DataFrame()
        st.session_state.filtros_aplicados = {}
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