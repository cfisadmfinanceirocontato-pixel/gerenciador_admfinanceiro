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
# FUNÇÕES DE FORMATAÇÃO
# ============================================================================
def formatar_data_completa(data_obj):
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
# FUNÇÃO PRÉ-VISUALIZAÇÃO - CHAMADA POSICIONAL ✅
# ============================================================================
def exibir_previa_registro(termo, instrumento, numero_termo, funcionario, cpf, cargo, 
                          qtd, qtd_extenso, valor, valor_extenso, objetivo, 
                          localidades, periodo, oficio, data_input):
    """Exibe tabela de pré-visualização do registro"""
    dados_previa = {
        'Campo': ['Termo', 'Instrumento', 'Nº Termo', 'Funcionário', 'CPF', 'Cargo',
                 'Quantidade', 'Qtd. Extenso', 'Valor', 'Valor Extenso', 'Objetivo',
                 'Localidades', 'Período', 'Ofício', 'Data'],
        'Valor': [termo, instrumento, numero_termo, funcionario, cpf, cargo,
                 qtd, qtd_extenso, valor, valor_extenso, objetivo,
                 localidades, periodo, oficio, data_input]
    }
    
    df_previa = pd.DataFrame(dados_previa)
    st.markdown("## ✅ **Pré-visualização do Registro**")
    st.dataframe(df_previa, use_container_width=True, hide_index=True)
    
    col_conf1, col_conf2 = st.columns([3, 1])
    with col_conf1:
        if st.button("✏️ **EDITAR**", key="editar_previa"):
            st.session_state.mostrar_previa = False
            st.rerun()
    with col_conf2:
        if st.button("✅ **CONFIRMAR SALVAR**", type="primary", key="confirmar_previa"):
            return True
    return False

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
# FUNÇÕES DE DADOS
# ============================================================================
@st.cache_data
def carregar_dados():
    try:
        dados = pd.read_csv("diarias_data.csv")
        return dados
    except:
        return pd.DataFrame(columns=[
            'Termo', 'Instrumento', 'Numero_Termo', 'Funcionario', 'CPF', 'Cargo', 
            'Qtd', 'Qtd_Extenso', 'Valor', 'Valor_Extenso', 'Objetivo', 
            'Localidades', 'Periodo', 'Oficio', 'Data'
        ])

def salvar_dados(dados_diarias):
    dados_diarias.to_csv("diarias_data.csv", index=False)

# ============================================================================
# INICIALIZAÇÃO SESSION STATE
# ============================================================================
if 'mostrar_previa' not in st.session_state:
    st.session_state.mostrar_previa = False
if 'dados_previa' not in st.session_state:
    st.session_state.dados_previa = None

dados_diarias = carregar_dados()

# ============================================================================
# DADOS APOIO E CONFIGURAÇÕES
# ============================================================================
dados_apoio = pd.DataFrame({
    'Numero_Termo': ['001/2024', '001/2024', '002/2024', '002/2024'],
    'Termo': ['TERMO1', 'TERMO1', 'TERMO2', 'TERMO2'],
    'Instrumento': ['INST 001/2024', 'INST 001/2024', 'INST 002/2024', 'INST 002/2024'],
    'Funcionario': ['João Silva', 'Maria Santos', 'Pedro Oliveira', 'Ana Costa'],
    'CPF': ['123.456.789-00', '987.654.321-00', '111.222.333-44', '555.666.777-88'],
    'Cargo': ['Analista', 'Técnica', 'Coordenador', 'Assistente']
})

opcoes_quantidade = ['0,0', '0,5', '1,5', '2,5', '3,5', '4,5']
termos_unicos = sorted(dados_apoio['Termo'].dropna().unique())

# ============================================================================
# SIDEBAR
# ============================================================================
with st.sidebar:
    st.header("🔍 Filtros")
    termo_selecionado = st.selectbox("Termo:", [''] + list(termos_unicos))
    funcionarios = []
    if termo_selecionado:
        mask = dados_apoio['Termo'] == termo_selecionado
        funcionarios = sorted(dados_apoio.loc[mask, 'Funcionario'].dropna().unique())
    funcionario_selecionado = st.selectbox("Funcionário:", [''] + funcionarios)

# ============================================================================
# PRÉ-VISUALIZAÇÃO - CHAMADA POSICIONAL ✅
# ============================================================================
if st.session_state.mostrar_previa and st.session_state.dados_previa:
    # ✅ CHAMADA POSICIONAL - NÃO USA **kwargs
    if exibir_previa_registro(
        st.session_state.dados_previa['Termo'],
        st.session_state.dados_previa['Instrumento'],
        st.session_state.dados_previa['Numero_Termo'],
        st.session_state.dados_previa['Funcionario'],
        st.session_state.dados_previa['CPF'],
        st.session_state.dados_previa['Cargo'],
        st.session_state.dados_previa['Qtd'],
        st.session_state.dados_previa['Qtd_Extenso'],
        st.session_state.dados_previa['Valor'],
        st.session_state.dados_previa['Valor_Extenso'],
        st.session_state.dados_previa['Objetivo'],
        st.session_state.dados_previa['Localidades'],
        st.session_state.dados_previa['Periodo'],
        st.session_state.dados_previa['Oficio'],
        st.session_state.dados_previa['Data']
    ):
        novo_registro = pd.DataFrame([st.session_state.dados_previa])
        dados_diarias = pd.concat([dados_diarias, novo_registro], ignore_index=True)
        salvar_dados(dados_diarias)
        st.success("🎉 **REGISTRO SALVO COM SUCESSO!**")
        st.balloons()
        st.session_state.mostrar_previa = False
        st.session_state.dados_previa = None
        st.rerun()

# ============================================================================
# FORMULÁRIO PRINCIPAL
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

# Termo
st.markdown("**📋 Dados do Termo**")
termo_input = st.selectbox("📄 **Termo**:", options=[''] + list(termos_unicos), index=0)
col_inst_num = st.columns([1, 1])
with col_inst_num[0]: instrumento = st.text_input("🎯 **Instrumento**:", key="instrumento")
with col_inst_num[1]: numero_termo = st.text_input("📍 **Nº Termo**:", key="numero_termo")

# Funcionário
st.markdown("**👤 Dados do Funcionário**")
funcionario_input = st.text_input("👤 **Funcionário**:", value=funcionario_selecionado)
col_cpf_cargo = st.columns([1, 1])
with col_cpf_cargo[0]: cpf = st.text_input("🆔 **CPF**:")
with col_cpf_cargo[1]: cargo = st.text_input("💼 **Cargo**:")

# Valores e Data
st.markdown("**💰 Valores e Data**")
col_qtd1, col_valor1, col_data1 = st.columns([1, 1, 1])

with col_qtd1:
    st.markdown("**🔢 Quantidade**")
    qtd = st.selectbox("", options=opcoes_quantidade, index=2, key="qtd_select")
    qtd_extenso = quantidade_por_extenso(qtd)
    st.text_input("Por extenso:", value=qtd_extenso, disabled=True)

with col_valor1:
    st.markdown("**💰 Valor**")
    qtd_num = float(qtd.replace(',', '.'))
    valor = formatar_moeda(qtd_num * 140)
    st.text_input("", value=valor, disabled=True)
    valor_extenso = valor_por_extenso(valor)
    st.text_input("Por extenso:", value=valor_extenso, disabled=True)

with col_data1:
    st.markdown("**📅 Data**")
    data_selecionada = st.date_input("", value=datetime.now().date(), label_visibility="collapsed", key="date_hidden")
    data_display = formatar_data_completa(data_selecionada)
    st.text_input("Por extenso:", value=data_display, disabled=True)
    data_input = formatar_data_csv(data_selecionada)

# Detalhes
st.markdown("**📋 Detalhes da Viagem**")
objetivo = st.text_area("🎯 **Objetivo**:", height=50)
localidades = st.text_area("📍 **Localidades**:", height=50)
periodo = st.text_input("📊 **Período**:")
oficio = st.text_input("📋 **Ofício**:")

# ============================================================================
# AÇÕES
# ============================================================================
st.markdown("---")
st.subheader("⚡ Ações")
col_acoes1, col_acoes2 = st.columns([3, 2])

with col_acoes1:
    if st.button("👁️ **PRÉ-VISUALIZAR REGISTRO**", type="primary", use_container_width=True):
        # ✅ Chaves MAIÚSCULAS para compatibilidade com CSV
        st.session_state.dados_previa = {
            'Termo': termo_input, 'Instrumento': instrumento, 'Numero_Termo': numero_termo,
            'Funcionario': funcionario_input, 'CPF': cpf, 'Cargo': cargo,
            'Qtd': qtd, 'Qtd_Extenso': qtd_extenso, 'Valor': valor, 'Valor_Extenso': valor_extenso,
            'Objetivo': objetivo, 'Localidades': localidades, 'Periodo': periodo, 
            'Oficio': oficio, 'Data': data_input
        }
        st.session_state.mostrar_previa = True
        st.rerun()

with col_acoes2:
    st.markdown("**⚙️ Ferramentas**")
    col2_btn1, col2_btn2 = st.columns(2)
    with col2_btn1:
        if st.button("🔄 Atualizar", use_container_width=True):
            st.rerun()
    with col2_btn2:
        if st.button("🗑️ Limpar", use_container_width=True):
            st.rerun()

# ============================================================================
# REGISTROS E ESTATÍSTICAS (resto igual)
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros")

if not dados_diarias.empty:
    colunas_prioritarias = ['Termo', 'Funcionario', 'Qtd', 'Qtd_Extenso', 'Valor', 'Valor_Extenso', 'Data']
    colunas_display = [col for col in colunas_prioritarias if col in dados_diarias.columns]
    df_display = dados_diarias[colunas_display].copy()
    edited_df = st.data_editor(df_display, num_rows="dynamic", use_container_width=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        csv = io.BytesIO()
        edited_df.to_csv(csv, index=False)
        csv.seek(0)
        st.download_button("📥 Excel", csv, f"diarias_{datetime.now().strftime('%d%m%Y_%H%M')}.csv", "text/csv")
    with col2:
        if st.button("💾 Salvar Tabela"):
            salvar_dados(edited_df)
            st.rerun()
    with col3:
        if st.button("🗑️ Limpar Tudo", type="secondary"):
            dados_diarias = pd.DataFrame(columns=dados_diarias.columns)
            salvar_dados(dados_diarias)
            st.rerun()
else:
    st.info("👆 Cadastre o primeiro registro!")

st.markdown("---")
st.subheader("📊 Resumo")
if not dados_diarias.empty:
    total_registros = len(dados_diarias)
    valores = [float(v.replace('R$', '').replace('.', '').replace(',', '.')) for v in dados_diarias['Valor'] if v]
    total_valor = sum(valores)
    col_estat1, col_estat2 = st.columns(2)
    with col_estat1: st.metric("📋 Total Registros", f"{total_registros:,}")
    with col_estat2: st.metric("💰 Valor Total", formatar_moeda(total_valor))
else:
    col_estat1, col_estat2 = st.columns(2)
    with col_estat1: st.metric("📋 Total Registros", "0")
    with col_estat2: st.metric("💰 Valor Total", "R$ 0,00")

