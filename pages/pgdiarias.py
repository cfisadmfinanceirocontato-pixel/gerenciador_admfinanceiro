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
# ✅ NOVA FUNÇÃO: EXTRAI CARACTERES ANTES DA "/"
# ============================================================================
def extrair_numero_oficio(oficio_completo):
    """✅ Extrai apenas caracteres ANTES da primeira '/' do ofício"""
    if pd.isna(oficio_completo) or not isinstance(oficio_completo, str) or '/' not in oficio_completo:
        return str(oficio_completo).strip()
    return oficio_completo.split('/')[0].strip()


# ============================================================================
# ✅ NOVA FUNÇÃO: CONCATENA NOME ARQUIVO + Nº OFÍCIO
# ============================================================================
def gerar_nome_recibo(nome_arquivo, numero_oficio):
    """✅ Concatena 'NomeArquivo_NºOfício' para Nome do Recibo"""
    if nome_arquivo and numero_oficio:
        return f"{nome_arquivo}_{numero_oficio}".strip()
    return ""


# ============================================================================
# FUNÇÕES DE CARREGAMENTO DE DADOS
# ============================================================================
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
        return ['TERMO1', 'TERMO2']  # Fallback
    except KeyError:
        st.error("❌ Coluna 'TERMO DE COLABORAÇÃO' não encontrada!")
        return ['TERMO1', 'TERMO2']
    except Exception as e:
        st.error(f"❌ Erro ao carregar CSV: {e}")
        return ['TERMO1', 'TERMO2']


# ✅ FUNÇÃO: Busca instrumento vinculado ao termo (3ª coluna)
@st.cache_data(ttl=300)
def buscar_instrumento_por_termo(termo):
    """🔍 Busca instrumento na 3ª coluna do CSV pelo termo selecionado"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 2:
            coluna_termo = df_colab.columns[2]  # 3ª coluna
            mask = df_colab[coluna_termo] == termo
            if mask.any():
                cols_instrumento = ['Instrumento', 'INSTRUMENTO', df_colab.columns[0]]
                for col in cols_instrumento:
                    if col in df_colab.columns:
                        return df_colab.loc[mask, col].iloc[0]
        return ""
    except:
        return ""


# ✅ FUNÇÃO: Busca NÚMERO do termo na 2ª coluna
@st.cache_data(ttl=300)
def buscar_numero_termo_por_nome(termo):
    """🔍 Busca Nº do termo na 2ª coluna (índice 1) do CSV"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 1:
            coluna_numero = df_colab.columns[1]  # 2ª coluna (índice 1)
            mask = df_colab['TERMO DE COLABORAÇÃO'] == termo
            if mask.any():
                numero_encontrado = df_colab.loc[mask, coluna_numero].iloc[0]
                return str(numero_encontrado).strip()
        return ""
    except:
        return ""


# ============================================================================
# ✅ FUNÇÕES ATUALIZADAS - SÉTIMA COLUNA PARA FUNCIONÁRIOS
# ============================================================================
@st.cache_data(ttl=300)
def carregar_funcionarios_por_termo(termo):
    """🔍 Carrega funcionários da OITAVA COLUNA (índice 7) do termo selecionado"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 6:  # Verifica se existe 8ª coluna
            coluna_oitava = df_colab.columns[6]  # Oitava coluna (índice 7)
            mask = df_colab['TERMO DE COLABORAÇÃO'] == termo
            if mask.any():
                funcionarios = sorted(df_colab.loc[mask, coluna_oitava].dropna().astype(str).unique())
                return funcionarios
        return []
    except:
        return []


@st.cache_data(ttl=300)
def buscar_cpf_cargo_por_funcionario(termo, funcionario):
    """🔍 Busca CPF (coluna específica) e Cargo pelo termo + funcionário da 8ª coluna"""
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        coluna_oitava = df_colab.columns[6] if len(df_colab.columns) > 6 else None
        
        if coluna_oitava:
            mask = (df_colab['TERMO DE COLABORAÇÃO'] == termo) & (df_colab[coluna_oitava] == funcionario)
            if mask.any():
                linha = df_colab.loc[mask].iloc[0]
                # Busca CPF na coluna 'CPF' ou 5ª coluna (índice 4)
                cpf = str(linha.get('CPF', linha.iloc[3] if len(linha) > 3 else '')).strip()
                # Busca Cargo na coluna 'Cargo' ou 6ª coluna (índice 5)
                cargo = str(linha.get('Cargo', linha.iloc[8] if len(linha) > 8 else '')).strip()
                return cpf, cargo
        return "", ""
    except:
        return "", ""


@st.cache_data
def carregar_dados_diarias():
    """Carrega dados das diárias"""
    try:
        dados = pd.read_csv("diarias_data.csv")
        if 'Data Recibo' in dados.columns and 'Data por Extenso' not in dados.columns:
            dados['Data por Extenso'] = dados['Data Recibo'].apply(formatar_data_completa)
        return dados
    except:
        return pd.DataFrame(columns=[
            'Instrumento', 'Termo de Colaboração', 'Funcionário', 'CPF', 'Cargo', 
            'Quantidade', 'Quantidade por Extenso', 'Valor', 'Valor por Extenso', 
            'Objetivo', 'Localidades', 'Período', 'Ofício', 'Data Recibo', 
            'Nome Arquivo', 'Nº do Ofício', 'Nome do Recibo', 'Data por Extenso',
            'Nº Do Termo de Colaboração'
        ])


def salvar_dados(dados_diarias):
    dados_diarias.to_csv("diarias_data.csv", index=False)


# ============================================================================
# FUNÇÕES DE FORMATAÇÃO - DATA NO PADRÃO DD/MM/YYYY
# ============================================================================
def formatar_data_completa(data_obj):
    """Formato: 01 de janeiro de 2026"""
    if pd.isna(data_obj) or data_obj == '':
        return "15 de fevereiro de 2026"
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
        return "15 de fevereiro de 2026"


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


# ============================================================================
# FUNÇÕES DE CONVERSÃO POR EXTENSO
# ============================================================================
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
dados_diarias = carregar_dados_diarias()
termos_unicos = carregar_termos_colaboracao()
opcoes_quantidade = ['0,0', '0,5', '1,5', '2,5', '3,5', '4,5']


# ============================================================================
# INTERFACE - SIDEBAR SIMPLIFICADA
# ============================================================================
with st.sidebar:
    st.header("🔍 Filtros")
    termo_filtro = st.selectbox("Filtrar por Termo:", ['Todos'] + termos_unicos)
    
    # Filtro por funcionário da 8ª coluna (opcional)
    if termo_filtro != 'Todos':
        funcs_filtro = carregar_funcionarios_por_termo(termo_filtro)
        funcionario_filtro = st.selectbox("Filtrar por Funcionário:", ['Todos'] + funcs_filtro)
    else:
        funcionario_filtro = 'Todos'


# ============================================================================
# FORMULÁRIO SIMPLIFICADO - ✅ COM VINCULAÇÃO COMPLETA
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

# ✅ DADOS DO TERMO - COM VINCULAÇÃO AUTOMÁTICA DE INSTRUMENTO E NÚMERO
st.markdown("**📋 Dados do Termo**")

termo_input = st.selectbox(" **Termo de Colaboração**:", options=[''] + termos_unicos, index=0)

# Busca automática do INSTRUMENTO (3ª coluna) e NÚMERO (2ª coluna)
instrumento_auto = buscar_instrumento_por_termo(termo_input) if termo_input else ""
numero_termo_auto = buscar_numero_termo_por_nome(termo_input) if termo_input else ""

col_termo_inst = st.columns([1, 1])

with col_termo_inst[0]:
    instrumento = st.text_input(" **Instrumento**:", 
                               value=instrumento_auto, 
                               help="Auto-preenchido pela 3ª coluna do dados_colaboradores.csv")

with col_termo_inst[1]:
    numero_termo = st.text_input(" **Nº Do Termo de Colaboração**:", 
                                value=numero_termo_auto,
                                help="Auto-preenchido pela 2ª coluna do dados_colaboradores.csv")

# Validação visual dos campos automáticos
if termo_input:
    if not instrumento_auto:
        st.warning("⚠️ Instrumento não encontrado para este termo")
    if not numero_termo_auto:
        st.warning("⚠️ Número do termo não encontrado na 2ª coluna")
    else:
        st.success(f"✅ Nº do Termo: {numero_termo_auto}")

# ✅ DADOS DO FUNCIONÁRIO - OITAVA COLUNA DO CSV
st.markdown("**👤 Dados do Funcionário**")

# Lista suspensa de funcionários da OITAVA COLUNA vinculada ao termo
funcionarios_do_termo = carregar_funcionarios_por_termo(termo_input)
funcionario_input = st.selectbox(
    "👤 **Funcionário**", 
    options=[''] + funcionarios_do_termo, 
    index=0,
    help="Selecione o termo primeiro para carregar funcionários da 8ª coluna do CSV"
)

# Auto-preenchimento CPF e Cargo baseado em termo + funcionário da 8ª coluna
cpf_auto, cargo_auto = "", ""
if termo_input and funcionario_input:
    cpf_auto, cargo_auto = buscar_cpf_cargo_por_funcionario(termo_input, funcionario_input)

col_func_cpf = st.columns([1, 1])
with col_func_cpf[0]:
    cpf = st.text_input("🆔 **CPF**:", value=cpf_auto, help="Auto-preenchido do CSV")

with col_func_cpf[1]:
    cargo = st.text_input("💼 **Cargo**:", value=cargo_auto, help="Auto-preenchido do CSV")

# Validação visual da vinculação da 8ª coluna
if termo_input and funcionario_input:
    if cpf_auto:
        st.success(f"✅ CPF: {cpf_auto}")
    else:
        st.warning("⚠️ CPF não encontrado no CSV")
    if cargo_auto:
        st.success(f"✅ Cargo: {cargo_auto}")
    else:
        st.warning("⚠️ Cargo não encontrado no CSV")
elif termo_input and not funcionarios_do_termo:
    st.warning("⚠️ Nenhum funcionário encontrado na 8ª coluna para este termo")

# Valores e Data
st.markdown("**💰 Valores e Data**")
col_qtd_valor = st.columns([1, 1])
with col_qtd_valor[0]:
    st.markdown("**🔢 Quantidade**")
    qtd = st.selectbox("Quantidade:", options=opcoes_quantidade, index=2, key="qtd_select")
    qtd_extenso = quantidade_por_extenso(qtd)
    st.text_input("Por extenso:", value=qtd_extenso, disabled=True)

with col_qtd_valor[1]:
    st.markdown("**💰 Valor**")
    qtd_num = float(qtd.replace(',', '.'))
    valor = formatar_moeda(qtd_num * 140)
    st.text_input("Valor:", value=valor, disabled=True, help="Quantidade × R$ 140,00")
    valor_extenso = valor_por_extenso(valor)
    st.text_input("Por extenso:", value=valor_extenso, disabled=True)

# Data
col_data_recibo, col_data_extenso = st.columns([1, 1])
with col_data_recibo:
    st.markdown("**📅 Data Recibo**")
    data_recibo = st.date_input("", value=datetime.now().date(), format="DD/MM/YYYY")
    data_input = formatar_data_csv(data_recibo)

with col_data_extenso:
    st.markdown("**📄 Data por Extenso**")
    data_extenso_display = formatar_data_completa(data_recibo)
    st.text_input("", value=data_extenso_display, disabled=True)

# ✅ CAMPOS ORGANIZADOS EM UMA ÚNICA LINHA - SEQUÊNCIA EXATA
st.markdown("**📋 Detalhes da Viagem**")
objetivo = st.text_area("🎯 **Objetivo**:", height=50)
localidades = st.text_area("📍 **Localidades**:", height=50)

# ✅ SEQUÊNCIA: PERÍODO | OFÍCIO | NOME ARQUIVO | Nº DO OFÍCIO | NOME DO RECIBO (AUTO)
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
    # ✅ AUTO-CONCATENA: NomeArquivo_NºOfício
    nome_recibo_auto = gerar_nome_recibo(nome_arquivo, numero_oficio_auto)
    nome_recibo = st.text_input("📄 **Nome do Recibo**:", 
                               value=nome_recibo_auto,
                               help="Auto: NomeArquivo_NºOfício")

# Feedback visual
if oficio and nome_arquivo_auto:
    numero_extraido = extrair_numero_oficio(oficio)
    nome_recibo_gerado = gerar_nome_recibo(nome_arquivo, numero_extraido)
    st.caption(f"✅ Nome Recibo auto: **{nome_recibo_gerado}**")

# ============================================================================
# AÇÕES SIMPLIFICADAS - ✅ COM PROCESSAMENTO AUTOMÁTICO DO NOME RECIBO
# ============================================================================
st.markdown("---")
st.subheader("⚡ Ações")

col_acoes1, col_acoes2, col_acoes3 = st.columns([3, 2, 3])

with col_acoes1:
    if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
        # ✅ GERA NOME DO RECIBO AUTOMATICAMENTE
        nome_recibo_final = gerar_nome_recibo(nome_arquivo, extrair_numero_oficio(numero_oficio_input))
        
        novo_registro = pd.DataFrame([{
            'Instrumento': instrumento,
            'Termo de Colaboração': termo_input,
            'Funcionário': funcionario_input,
            'CPF': cpf,
            'Cargo': cargo,
            'Quantidade': qtd,
            'Quantidade por Extenso': qtd_extenso,
            'Valor': valor,
            'Valor por Extenso': valor_extenso,
            'Objetivo': objetivo,
            'Localidades': localidades,
            'Período': periodo,
            'Ofício': oficio,  # Campo COMPLETO
            'Data Recibo': data_input,
            'Nome Arquivo': nome_arquivo,
            'Nº do Ofício': extrair_numero_oficio(numero_oficio_input),  # ✅ APENAS ANTES DA "/"
            'Nome do Recibo': nome_recibo_final,  # ✅ CONCATENADO AUTOMATICAMENTE
            'Data por Extenso': data_extenso_display,
            'Nº Do Termo de Colaboração': numero_termo
        }])
        
        dados_diarias = pd.concat([dados_diarias, novo_registro], ignore_index=True)
        salvar_dados(dados_diarias)
        st.success(f"✅ Registro salvo! Nome Recibo: **{nome_recibo_final}**")
        st.balloons()
        st.rerun()

with col_acoes2:
    col2_btn1, col2_btn2 = st.columns(2)
    with col2_btn1:
        if st.button("🔄 Atualizar", use_container_width=True):
            st.rerun()
    with col2_btn2:
        if st.button("🗑️ Limpar", use_container_width=True):
            st.rerun()

# ============================================================================
# TABELA COMPLETA - ✅ COM PROCESSAMENTO AUTOMÁTICO
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros")

if not dados_diarias.empty:
    # ✅ APLICA PROCESSAMENTO EM TODOS OS REGISTROS EXISTENTES
    if 'Nº do Ofício' in dados_diarias.columns:
        dados_diarias['Nº do Ofício'] = dados_diarias['Nº do Ofício'].apply(extrair_numero_oficio)
    
    colunas_completas = [
        'Instrumento', 'Termo de Colaboração', 'Nº Do Termo de Colaboração',
        'Funcionário', 'CPF', 'Cargo', 
        'Quantidade', 'Quantidade por Extenso', 'Valor', 'Valor por Extenso', 
        'Objetivo', 'Localidades', 'Período', 'Ofício', 'Data Recibo', 
        'Nome Arquivo', 'Nº do Ofício', 'Nome do Recibo', 'Data por Extenso'
    ]
    
    colunas_display = [col for col in colunas_completas if col in dados_diarias.columns]
    df_display = dados_diarias[colunas_display].copy()
    
    edited_df = st.data_editor(
        df_display, 
        num_rows="dynamic", 
        use_container_width=True,
        column_config={
            "Valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
            "Quantidade": st.column_config.NumberColumn("Quantidade", format="%.1f")
        }
    )
    
    col1, col2, col3 = st.columns(3)
    with col1:
        csv = io.BytesIO()
        edited_df.to_csv(csv, index=False)
        csv.seek(0)
        st.download_button(
            "📥 Excel", 
            csv, 
            f"diarias_{datetime.now().strftime('%d%m%Y_%H%M')}.csv", 
            "text/csv"
        )
    
    with col2:
        if st.button("💾 Salvar Tabela", use_container_width=True):
            if 'Data Recibo' in edited_df.columns:
                edited_df['Data por Extenso'] = edited_df['Data Recibo'].apply(formatar_data_completa)
            # ✅ PROCESSA NÚMERO DO OFÍCIO NA TABELA EDITADA
            if 'Nº do Ofício' in edited_df.columns:
                edited_df['Nº do Ofício'] = edited_df['Nº do Ofício'].apply(extrair_numero_oficio)
            salvar_dados(edited_df)
            st.success("✅ Tabela salva!")
            st.rerun()
    
    with col3:
        if st.button("🗑️ Limpar Tudo", type="secondary", use_container_width=True):
            dados_diarias = pd.DataFrame(columns=colunas_completas)
            salvar_dados(dados_diarias)
            st.rerun()
else:
    st.info("👆 Cadastre o primeiro registro!")

# ============================================================================
# ESTATÍSTICAS
# ============================================================================
st.markdown("---")
st.subheader("📊 Resumo dos Registros")

if not dados_diarias.empty:
    total_registros = len(dados_diarias)
    
    # ✅ CORREÇÃO: Verifica se é string ANTES de usar replace()
    valores = []
    for v in dados_diarias['Valor']:
        if pd.notna(v):
            if isinstance(v, str):
                valor_limpo = v.replace('R$', '').replace('.', '').replace(',', '.').strip()
                try:
                    valores.append(float(valor_limpo))
                except:
                    pass
            else:
                valores.append(float(v))
    
    total_valor = sum(valores) if valores else 0
    
    col_estat1, col_estat2 = st.columns(2)
    with col_estat1:
        st.metric("📋 Total Registros", f"{total_registros:,}")
    with col_estat2:
        st.metric("💰 Valor Total", formatar_moeda(total_valor))
else:
    col_estat1, col_estat2 = st.columns(2)
    with col_estat1:
        st.metric("📋 Total Registros", "0")
    with col_estat2:
        st.metric("💰 Valor Total", "R$ 0,00")

st.markdown("---")
st.caption("✅ Nome Recibo = NomeArquivo_NºOfício | 5 CAMPOS EM 1 LINHA | AUTO-COMPLETO!")
