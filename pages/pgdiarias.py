import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io

# ============================================================================
# CONFIGURAÇÃO DA APLICAÇÃO
# ============================================================================
st.set_page_config(
    page_title="Pagto Diárias", 
    page_icon="📋", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# ============================================================================
# ✅ NOVA FUNÇÃO: SALVAR REGISTRO NA ORDEM EXATA SOLICITADA
# ============================================================================
def salvar_registro_formulario():
    """✅ Salva TODOS os dados do formulário nas 18 colunas EXATAS solicitadas"""
    
    # ✅ COLETA TODOS OS DADOS NA ORDEM EXATA
    dados_registro = {
        'Termo de Colaboração': termo_input,
        'Instrumento': instrumento,
        'Nº Do Termo de Colaboração': numero_termo,
        'Funcionário': funcionario_input,
        'CPF': cpf,
        'Cargo': cargo,
        'Quantidade': qtd,
        'Quantidade por extenso': qtd_extenso,
        'Por extenso': valor_extenso,  # Valor por extenso (campo "Por extenso")
        'Data Recibo': data_input,
        'Data por Extenso': data_extenso_display,
        'Objetivo': objetivo,
        'Localidades': localidades,
        'Período': periodo,
        'Ofício': oficio,
        'Nome Arquivo': nome_arquivo,
        'Nº do Ofício': numero_oficio_input,
        'Nome do Recibo': nome_recibo
    }
    
    # ✅ 18 COLUNAS NA ORDEM EXATA SOLICITADA
    colunas_planilha = [
        'Termo de Colaboração', 'Instrumento', 'Nº Do Termo de Colaboração', 
        'Funcionário', 'CPF', 'Cargo', 'Quantidade', 'Quantidade por extenso', 
        'Por extenso', 'Data Recibo', 'Data por Extenso', 'Objetivo', 
        'Localidades', 'Período', 'Ofício', 'Nome Arquivo', 'Nº do Ofício', 
        'Nome do Recibo'
    ]
    
    novo_registro = pd.DataFrame([dados_registro])[colunas_planilha]
    
    # ✅ CARREGA OU CRIA A PLANILHA PRINCIPAL
    try:
        dados_existentes = pd.read_csv("registros_completos.csv")
        dados_atualizados = pd.concat([dados_existentes, novo_registro], ignore_index=True)
    except:
        dados_atualizados = novo_registro
    
    # ✅ SALVA NA PLANILHA
    dados_atualizados.to_csv("registros_completos.csv", index=False)
    return dados_atualizados, novo_registro

# ============================================================================
# FUNÇÕES EXISTENTES (mantidas)
# ============================================================================
def extrair_numero_oficio(oficio_completo):
    """✅ Extrai apenas caracteres ANTES da primeira '/' do ofício"""
    if pd.isna(oficio_completo) or not isinstance(oficio_completo, str) or '/' not in oficio_completo:
        return str(oficio_completo).strip()
    return oficio_completo.split('/')[0].strip()

def gerar_nome_recibo(nome_arquivo, numero_oficio):
    """✅ Concatena 'NomeArquivo_NºOfício' para Nome do Recibo"""
    if nome_arquivo and numero_oficio:
        return f"{nome_arquivo}_{numero_oficio}".strip()
    return ""

@st.cache_data(ttl=300)
def carregar_termos_colaboracao():
    """✅ Carrega termos únicos do arquivo dados_colaboradores.csv"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        termos = sorted(df_colab['TERMO DE COLABORAÇÃO'].dropna().astype(str).unique())
        st.success(f"✅ {len(termos)} termos carregados de dados_colaboradores.csv")
        return termos
    except FileNotFoundError:
        st.error("❌ Arquivo 'dados_colaboradores.csv' não encontrado!")
        return ['TERMO1', 'TERMO2']
    except KeyError:
        st.error("❌ Coluna 'TERMO DE COLABORAÇÃO' não encontrada!")
        return ['TERMO1', 'TERMO2']
    except Exception as e:
        st.error(f"❌ Erro ao carregar CSV: {e}")
        return ['TERMO1', 'TERMO2']

@st.cache_data(ttl=300)
def buscar_instrumento_por_termo(termo):
    """🔍 Busca instrumento na 3ª coluna do CSV pelo termo selecionado"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 2:
            coluna_termo = df_colab.columns[2]
            mask = df_colab[coluna_termo] == termo
            if mask.any():
                cols_instrumento = ['Instrumento', 'INSTRUMENTO', df_colab.columns[0]]
                for col in cols_instrumento:
                    if col in df_colab.columns:
                        return df_colab.loc[mask, col].iloc[0]
        return ""
    except:
        return ""

@st.cache_data(ttl=300)
def buscar_numero_termo_por_nome(termo):
    """🔍 Busca Nº do termo na 2ª coluna (índice 1) do CSV"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 1:
            coluna_numero = df_colab.columns[1]
            mask = df_colab['TERMO DE COLABORAÇÃO'] == termo
            if mask.any():
                numero_encontrado = df_colab.loc[mask, coluna_numero].iloc[0]
                return str(numero_encontrado).strip()
        return ""
    except:
        return ""

@st.cache_data(ttl=300)
def carregar_funcionarios_por_termo(termo):
    """🔍 Carrega funcionários da OITAVA COLUNA (índice 7) do termo selecionado"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 6:
            coluna_oitava = df_colab.columns[6]
            mask = df_colab['TERMO DE COLABORAÇÃO'] == termo
            if mask.any():
                funcionarios = sorted(df_colab.loc[mask, coluna_oitava].dropna().astype(str).unique())
                return funcionarios
        return []
    except:
        return []

@st.cache_data(ttl=300)
def buscar_cpf_cargo_por_funcionario(termo, funcionario):
    """🔍 Busca CPF e Cargo pelo termo + funcionário da 8ª coluna"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        coluna_oitava = df_colab.columns[6] if len(df_colab.columns) > 6 else None
        
        if coluna_oitava:
            mask = (df_colab['TERMO DE COLABORAÇÃO'] == termo) & (df_colab[coluna_oitava] == funcionario)
            if mask.any():
                linha = df_colab.loc[mask].iloc[0]
                cpf = str(linha.get('CPF', linha.iloc[3] if len(linha) > 3 else '')).strip()
                cargo = str(linha.get('Cargo', linha.iloc[8] if len(linha) > 8 else '')).strip()
                return cpf, cargo
        return "", ""
    except:
        return "", ""

def formatar_data_completa(data_obj):
    """Formato: 01 de janeiro de 2026"""
    if pd.isna(data_obj) or data_obj == '':
        return "16 de fevereiro de 2026"
    try:
        if isinstance(data_obj, date):
            data = data_obj
        else:
            data = pd.to_datetime(data_obj, dayfirst=True).date()
        meses = {1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril', 5: 'maio', 
                6: 'junho', 7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 
                11: 'novembro', 12: 'dezembro'}
        return f"{data.day:02d} de {meses[data.month]} de {data.year}"
    except:
        return "16 de fevereiro de 2026"

def formatar_data_csv(data_obj):
    """✅ PADRÃO DD/MM/YYYY para CSV e exibição"""
    if pd.isna(data_obj) or data_obj == '':
        return datetime.now().strftime("%d/%m/%Y")
    try:
        if isinstance(data_obj, date):
            return data_obj.strftime("%d/%m/%Y")
        return pd.to_datetime(data_obj, dayfirst=True).strftime("%d/%m/%Y")
    except:
        return str(data_obj)

def formatar_moeda(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"

def numero_extenso(n):
    if n == 0: return ''
    unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
    dezenas = ['', 'dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']
    
    if n < 10: return unidades[n]
    elif n < 100:
        d, u = divmod(n, 10)
        if u == 0: return dezenas[d]
        return f"{dezenas[d]} e {unidades[u]}"
    else:
        c, resto = divmod(n, 100)
        if resto == 0:
            return 'cem' if c == 1 else centenas[c]
        return f"{'cem' if c == 1 else centenas[c]} e {numero_extenso(resto)}"

def quantidade_por_extenso(qtd_str):
    try:
        qtd_num = float(qtd_str.replace(',', '.'))
        inteira = int(qtd_num)
        decimal = int((qtd_num - inteira) * 10 + 0.5)
        if decimal == 5:
            if inteira == 0: return "meia"
            elif inteira == 1: return "uma e meia"
            elif inteira == 2: return "duas e meia"
            elif inteira == 3: return "três e meia"
            elif inteira == 4: return "quatro e meia"
        return numero_extenso(inteira)
    except:
        return "quantidade inválida"

def valor_por_extenso(valor_str):
    try:
        valor_limpo = valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
        valor_num = float(valor_limpo)
        reais = int(valor_num)
        centavos = int((valor_num - reais) * 100 + 0.5)
        texto_reais = numero_extenso(reais)
        if reais == 1: texto_reais += " real"
        elif reais == 0: texto_reais = "zero"
        else: texto_reais += " reais"
        if centavos > 0:
            texto_centavos = numero_extenso(centavos)
            sufixo = " centavo" if centavos == 1 else " centavos"
            return f"{texto_reais} e {texto_centavos}{sufixo}"
        return texto_reais
    except:
        return "valor inválido"

# ============================================================================
# DADOS INICIAIS
# ============================================================================
termos_unicos = carregar_termos_colaboracao()
opcoes_quantidade = ['0,0', '0,5', '1,5', '2,5', '3,5', '4,5']

# ============================================================================
# INTERFACE - SIDEBAR
# ============================================================================
with st.sidebar:
    st.header("🔍 Filtros")
    termo_filtro = st.selectbox("Filtrar por Termo:", ['Todos'] + termos_unicos)

# ============================================================================
# FORMULÁRIO PRINCIPAL
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

# ✅ DADOS DO TERMO
st.markdown("**📋 Dados do Termo**")
termo_input = st.selectbox(" **Termo de Colaboração**:", options=[''] + termos_unicos, index=0)

instrumento_auto = buscar_instrumento_por_termo(termo_input) if termo_input else ""
numero_termo_auto = buscar_numero_termo_por_nome(termo_input) if termo_input else ""

col_termo_inst = st.columns([1, 1])
with col_termo_inst[0]:
    instrumento = st.text_input(" **Instrumento**:", value=instrumento_auto)
with col_termo_inst[1]:
    numero_termo = st.text_input(" **Nº Do Termo de Colaboração**:", value=numero_termo_auto)

# ✅ DADOS DO FUNCIONÁRIO
st.markdown("**👤 Dados do Funcionário**")
funcionarios_do_termo = carregar_funcionarios_por_termo(termo_input)
funcionario_input = st.selectbox(
    "👤 **Funcionário**", 
    options=[''] + funcionarios_do_termo, 
    index=0
)

cpf_auto, cargo_auto = "", ""
if termo_input and funcionario_input:
    cpf_auto, cargo_auto = buscar_cpf_cargo_por_funcionario(termo_input, funcionario_input)

col_func_cpf = st.columns([1, 1])
with col_func_cpf[0]:
    cpf = st.text_input("🆔 **CPF**:", value=cpf_auto)
with col_func_cpf[1]:
    cargo = st.text_input("💼 **Cargo**:", value=cargo_auto)

# Valores e Data
st.markdown("**💰 Valores e Data**")
col_qtd_valor = st.columns([1, 1])
with col_qtd_valor[0]:
    st.markdown("**🔢 Quantidade**")
    qtd = st.selectbox("Quantidade:", options=opcoes_quantidade, index=2, key="qtd_select")
    qtd_extenso = quantidade_por_extenso(qtd)
    st.text_input("Quantidade por extenso:", value=qtd_extenso, disabled=True)

with col_qtd_valor[1]:
    st.markdown("**💰 Valor**")
    qtd_num = float(qtd.replace(',', '.'))
    valor = formatar_moeda(qtd_num * 140)
    st.text_input("Valor:", value=valor, disabled=True)
    valor_extenso = valor_por_extenso(valor)
    st.text_input("Valor por extenso:", value=valor_extenso, disabled=True)

col_data_recibo, col_data_extenso = st.columns([1, 1])
with col_data_recibo:
    st.markdown("**📅 Data Recibo**")
    data_recibo = st.date_input("", value=datetime.now().date(), format="DD/MM/YYYY")
    data_input = formatar_data_csv(data_recibo)

with col_data_extenso:
    st.markdown("**📄 Data por Extenso**")
    data_extenso_display = formatar_data_completa(data_recibo)
    st.text_input("", value=data_extenso_display, disabled=True)

# Detalhes da Viagem
st.markdown("**📋 Detalhes da Viagem**")
objetivo = st.text_area("🎯 **Objetivo**:", height=50)
localidades = st.text_area("📍 **Localidades**:", height=50)

# Campos Obrigatórios
st.markdown("**📄 Campos Obrigatórios**")
col_periodo_oficio_arquivo = st.columns([1.5, 2, 1.2, 1.2, 1.1])

with col_periodo_oficio_arquivo[0]:
    periodo = st.text_input("📊 **Período**:", placeholder="01/02 a 03/02")

with col_periodo_oficio_arquivo[1]:
    oficio = st.text_input("📋 **Ofício**:", placeholder="123/2026/PGF")

with col_periodo_oficio_arquivo[2]:
    nome_arquivo_auto = funcionario_input.split()[0] if funcionario_input else ""
    nome_arquivo = st.text_input("📁 **Nome Arquivo**:", value=nome_arquivo_auto)

with col_periodo_oficio_arquivo[3]:
    numero_oficio_auto = extrair_numero_oficio(oficio)
    numero_oficio_input = st.text_input("📄 **Nº do Ofício**:", value=numero_oficio_auto)

with col_periodo_oficio_arquivo[4]:
    nome_recibo_auto = gerar_nome_recibo(nome_arquivo, numero_oficio_auto)
    nome_recibo = st.text_input("📄 **Nome do Recibo**:", value=nome_recibo_auto)

# ============================================================================
# ✅ BOTÃO SALVAR COM NOVA FUNÇÃO
# ============================================================================
st.markdown("---")
st.subheader("⚡ Ações")

col_acoes1, col_acoes2, col_acoes3 = st.columns([3, 2, 3])

with col_acoes1:
    if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
        try:
            dados_planilha, registro_salvo = salvar_registro_formulario()
            
            st.success(f"✅ Registro salvo com sucesso!")
            st.success(f"📊 Nome Recibo: **{registro_salvo['Nome do Recibo'].iloc[0]}**")
            st.balloons()
            st.rerun()
        except Exception as e:
            st.error(f"❌ Erro ao salvar: {str(e)}")

with col_acoes2:
    col2_btn1, col2_btn2 = st.columns(2)
    with col2_btn1:
        if st.button("🔄 Atualizar", use_container_width=True):
            st.rerun()
    with col2_btn2:
        if st.button("🗑️ Limpar", use_container_width=True):
            st.rerun()

# ============================================================================
# TABELA DE REGISTROS
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros")

try:
    dados_completos = pd.read_csv("registros_completos.csv")
    if not dados_completos.empty:
        st.dataframe(dados_completos, use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            csv = io.BytesIO()
            dados_completos.to_csv(csv, index=False)
            csv.seek(0)
            st.download_button(
                "📥 Download CSV", 
                csv, 
                f"registros_completos_{datetime.now().strftime('%d%m%Y_%H%M')}.csv", 
                "text/csv"
            )
        with col2:
            if st.button("🗑️ Limpar Tudo", type="secondary", use_container_width=True):
                pd.DataFrame(columns=dados_completos.columns).to_csv("registros_completos.csv", index=False)
                st.rerun()
    else:
        st.info("👆 Cadastre o primeiro registro!")
except:
    st.info("👆 Cadastre o primeiro registro na planilha!")

# ============================================================================
# ESTATÍSTICAS
# ============================================================================
st.markdown("---")
st.subheader("📊 Resumo")

try:
    dados_completos = pd.read_csv("registros_completos.csv")
    total_registros = len(dados_completos)
    
    st.metric("📋 Total Registros", f"{total_registros:,}")
    
except:
    st.metric("📋 Total Registros", "0")

st.markdown("---")
st.caption("✅ Planilha salva em: **registros_completos.csv** | 18 colunas na ordem exata!")
