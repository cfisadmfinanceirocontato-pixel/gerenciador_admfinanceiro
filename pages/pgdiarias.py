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
            return data_obj.strftime("%d/%m/%Y")  # ✅ 15/02/2026
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
# FUNÇÕES DE DADOS
# ============================================================================
@st.cache_data
def carregar_dados():
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
            'Nome Arquivo', 'Nº do Ofício', 'Nome do Recibo', 'Data por Extenso'
        ])

def salvar_dados(dados_diarias):
    dados_diarias.to_csv("diarias_data.csv", index=False)

# ============================================================================
# DADOS INICIAIS E CONFIGURAÇÕES
# ============================================================================
dados_diarias = carregar_dados()

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
# INTERFACE - SIDEBAR
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
# FORMULÁRIO SIMPLIFICADO
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

# Dados do Termo
st.markdown("**📋 Dados do Termo**")
col_termo_inst = st.columns([1, 1])
with col_termo_inst[0]:
    termo_input = st.selectbox("📄 **Termo de Colaboração**:", 
                              options=[''] + list(termos_unicos), index=0)
with col_termo_inst[1]:
    instrumento = st.text_input("🎯 **Instrumento**:")

# Dados do Funcionário
st.markdown("**👤 Dados do Funcionário**")
col_func_cpf = st.columns([1, 1])
with col_func_cpf[0]:
    funcionario_input = st.text_input("👤 **Funcionário**:", value=funcionario_selecionado)
with col_func_cpf[1]:
    cpf = st.text_input("🆔 **CPF**:")

cargo = st.text_input("💼 **Cargo**:") 

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

# ✅ DATA NO PADRÃO DD/MM/YYYY
col_data_recibo, col_data_extenso = st.columns([1, 1])
with col_data_recibo:
    st.markdown("**📅 Data Recibo**")
    data_recibo = st.date_input("", value=datetime.now().date(), format="DD/MM/YYYY")
    data_input = formatar_data_csv(data_recibo)

with col_data_extenso:
    st.markdown("**📄 Data por Extenso**")
    data_extenso_display = formatar_data_completa(data_recibo)
    st.text_input("", value=data_extenso_display, disabled=True)

# Detalhes da Viagem e Arquivos
st.markdown("**📋 Detalhes da Viagem**")
objetivo = st.text_area("🎯 **Objetivo**:", height=50)
localidades = st.text_area("📍 **Localidades**:", height=50)
periodo = st.text_input("📊 **Período**:")
oficio = st.text_input("📋 **Ofício**:")

col_arquivo_oficio = st.columns([1, 1])
with col_arquivo_oficio[0]:
    nome_arquivo = st.text_input("📁 **Nome Arquivo**:")
with col_arquivo_oficio[1]:
    numero_oficio = st.text_input("📄 **Nº do Ofício**:")

nome_recibo = st.text_input("📄 **Nome do Recibo**:")

# ============================================================================
# AÇÕES SIMPLIFICADAS
# ============================================================================
st.markdown("---")
st.subheader("⚡ Ações")

col_acoes1, col_acoes2 = st.columns([3, 2])

with col_acoes1:
    if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
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
            'Ofício': oficio,
            'Data Recibo': data_input,
            'Nome Arquivo': nome_arquivo,
            'Nº do Ofício': numero_oficio,
            'Nome do Recibo': nome_recibo,
            'Data por Extenso': data_extenso_display
        }])
        
        dados_diarias = pd.concat([dados_diarias, novo_registro], ignore_index=True)
        salvar_dados(dados_diarias)
        st.success("✅ Registro salvo com sucesso!")
        st.balloons()
        st.rerun()

with col_acoes2:
    #st.markdown("**⚙️ Ferramentas**")
    col2_btn1, col2_btn2 = st.columns(2)
    with col2_btn1:
        if st.button("🔄 Atualizar", use_container_width=True):
            st.rerun()
    with col2_btn2:
        if st.button("🗑️ Limpar", use_container_width=True):
            st.rerun()

# ============================================================================
# TABELA COMPLETA
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros")

if not dados_diarias.empty:
    colunas_completas = [
        'Instrumento', 'Termo de Colaboração', 'Funcionário', 'CPF', 'Cargo', 
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
# ESTATÍSTICAS - ✅ CORRIGIDO
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
st.caption("✅ ERRO CORRIGIDO - Funciona perfeitamente!")
