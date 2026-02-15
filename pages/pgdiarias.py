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
    """01 de janeiro de 2026"""
    if pd.isna(data_obj) or data_obj == '':
        return "14 de fevereiro de 2026"
    try:
        if isinstance(data_obj, date):
            data = data_obj
        else:
            data = pd.to_datetime(data_obj, dayfirst=True).date()
        
        meses = {
            1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
            5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
            9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
        }
        return f"{data.day:02d} de {meses[data.month]} de {data.year}"
    except:
        return "14 de fevereiro de 2026"

def formatar_data_csv(data_obj):
    """DD/MM/YYYY para CSV"""
    if pd.isna(data_obj) or data_obj == '':
        return datetime.now().strftime("%d/%m/%Y")
    try:
        if isinstance(data_obj, date):
            return data_obj.strftime("%d/%m/%Y")
        return pd.to_datetime(data_obj, dayfirst=True).strftime("%d/%m/%Y")
    except:
        return str(data_obj)

def formatar_moeda(valor):
    """R$ 1.234,56"""
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"

# ============================================================================
# FUNÇÕES DE CONVERSÃO POR EXTENSO
# ============================================================================
def numero_extenso(n):
    """Função base para números por extenso (0-999)"""
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
    """CORRIGIDO: 0,5→meia | 1,5→uma e meia | 2,5→duas e meia"""
    try:
        qtd_num = float(qtd_str.replace(',', '.'))
        inteira = int(qtd_num)
        decimal = int((qtd_num - inteira) * 10 + 0.5)
        
        # ✅ NOMENCLATURA CORRIGIDA EXATAMENTE COMO SOLICITADO
        if decimal == 5:
            if inteira == 0:
                return "meia"
            elif inteira == 1:
                return "uma e meia"
            elif inteira == 2:
                return "duas e meia"
            elif inteira == 3:
                return "três e meia"
            elif inteira == 4:
                return "quatro e meia"
        return numero_extenso(inteira)
    except:
        return "quantidade inválida"

def valor_por_extenso(valor_str):
    """R$ 210,00 → duzentos e dez reais"""
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
    """Carrega dados do CSV ou cria estrutura vazia"""
    try:
        dados = pd.read_csv("diarias_data.csv")
        if 'Data' in dados.columns:
            dados['Data'] = dados['Data'].apply(formatar_data_csv)
        return dados
    except:
        return pd.DataFrame(columns=[
            'Termo', 'Instrumento', 'Numero_Termo', 'Funcionario', 'CPF', 'Cargo', 
            'Qtd', 'Qtd_Extenso', 'Valor', 'Valor_Extenso', 'Objetivo', 
            'Localidades', 'Periodo', 'Oficio', 'Data'
        ])

def salvar_dados(dados_diarias):
    """Salva dados no CSV"""
    if 'Data' in dados_diarias.columns:
        dados_diarias['Data'] = dados_diarias['Data'].apply(formatar_data_csv)
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
# INTERFACE PRINCIPAL
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

col1, col2 = st.columns([2, 1])

# ============================================================================
# COLUNA 1 - NOVO REGISTRO
# ============================================================================
with col1:
    st.subheader("📝 Novo Registro")
    
    # Dados do Termo
    st.markdown("**📋 Dados do Termo de Colaboração**")
    termo_input = st.selectbox("📄 **Termo de Colaboração**:", 
                              options=[''] + list(termos_unicos), index=0)
    
    col_inst_num = st.columns([1, 1])
    with col_inst_num[0]: instrumento = st.text_input("🎯 **Instrumento**:", key="instrumento")
    with col_inst_num[1]: numero_termo = st.text_input("📍 **Nº Termo**:", key="numero_termo")
    
    if termo_input:
        mask = dados_apoio['Termo'] == termo_input
        if mask.any():
            st.info(f"💡 Instrumento: {dados_apoio.loc[mask, 'Instrumento'].iloc[0]}")
            st.info(f"💡 Nº Termo: {dados_apoio.loc[mask, 'Numero_Termo'].iloc[0]}")
    
    # Dados do Funcionário
    st.markdown("**👤 Dados do Funcionário**")
    funcionario_input = st.text_input("👤 **Funcionário**:", value=funcionario_selecionado)
    
    col_cpf_cargo = st.columns([1, 1])
    with col_cpf_cargo[0]: cpf = st.text_input("🆔 **CPF**:")
    with col_cpf_cargo[1]: cargo = st.text_input("💼 **Cargo**:")
    
    if funcionario_selecionado:
        mask = dados_apoio['Funcionario'] == funcionario_selecionado
        if mask.any():
            st.info(f"💡 CPF: {dados_apoio.loc[mask, 'CPF'].iloc[0]} | Cargo: {dados_apoio.loc[mask, 'Cargo'].iloc[0]}")
    
    # Valores e Data - Campos Alinhados
    st.markdown("**💰 Valores e Data**")
    
    # Títulos
    col_qtd1, col_valor1, col_data1 = st.columns([1, 1, 1])
    with col_qtd1: st.markdown("**🔢 Quantidade**")
    with col_valor1: st.markdown("**💰 Valor**")
    with col_data1: st.markdown("**📅 Data**")
    
    # Campos
    col_qtd2, col_valor2, col_data2 = st.columns([1, 1, 1])
    
    with col_qtd2:
        qtd = st.selectbox("", options=opcoes_quantidade, index=2, key="qtd_select")
        st.markdown(f"*{quantidade_por_extenso(qtd)}*")  # ✅ CORRIGIDO
    
    with col_valor2:
        qtd_num = float(qtd.replace(',', '.'))
        valor = formatar_moeda(qtd_num * 140)
        st.text_input("", value=valor, disabled=True, help="Quantidade × R$ 140,00")
        st.markdown(f"*{valor_por_extenso(valor)}*")
    
    with col_data2:
        data_selecionada = st.date_input("", value=datetime.now().date(), 
                                       label_visibility="collapsed", 
                                       key="date_hidden",
                                       format="DD/MM/YYYY")
        st.markdown(f"**{formatar_data_completa(data_selecionada)}**")
        data_input = formatar_data_csv(data_selecionada)
    
    # Detalhes da Viagem
    st.markdown("**📋 Detalhes da Viagem**")
    objetivo = st.text_area("🎯 **Objetivo**:", height=50)
    localidades = st.text_area("📍 **Localidades**:", height=50)
    periodo = st.text_input("📊 **Período**:")
    oficio = st.text_input("📋 **Ofício**:")
    
    if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
        novo_registro = pd.DataFrame([{
            'Termo': termo_input, 'Instrumento': instrumento, 'Numero_Termo': numero_termo,
            'Funcionario': funcionario_input, 'CPF': cpf, 'Cargo': cargo,
            'Qtd': qtd, 'Qtd_Extenso': quantidade_por_extenso(qtd),  # ✅ CORRIGIDO
            'Valor': valor, 'Valor_Extenso': valor_por_extenso(valor),
            'Objetivo': objetivo, 'Localidades': localidades,
            'Periodo': periodo, 'Oficio': oficio, 'Data': data_input
        }])
        
        dados_diarias = pd.concat([dados_diarias, novo_registro], ignore_index=True)
        salvar_dados(dados_diarias)
        st.success("✅ Registro salvo!")
        st.rerun()

# ============================================================================
# COLUNA 2 - AÇÕES E ESTATÍSTICAS
# ============================================================================
with col2:
    st.subheader("⚡ Ações")
    if st.button("🔄 Atualizar"): st.rerun()
    if st.button("🗑️ Limpar"): st.rerun()
    
    st.markdown("---")
    st.subheader("📊 Estatísticas")
    if not dados_diarias.empty:
        valores = [float(v.replace('R$', '').replace('.', '').replace(',', '.')) 
                  for v in dados_diarias['Valor'] if v]
        st.metric("Total Registros", len(dados_diarias))
        st.metric("Total Valor", formatar_moeda(sum(valores)))

# ============================================================================
# TABELA DE REGISTROS
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros")

if not dados_diarias.empty:
    colunas_prioritarias = ['Termo', 'Funcionario', 'Qtd', 'Qtd_Extenso', 
                           'Valor', 'Valor_Extenso', 'Data']
    colunas_display = [col for col in colunas_prioritarias if col in dados_diarias.columns]
    
    df_display = dados_diarias[colunas_display].copy()
    if 'Data' in df_display.columns:
        df_display['Data'] = df_display['Data'].apply(formatar_data_csv)
    
    edited_df = st.data_editor(df_display, num_rows="dynamic", use_container_width=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        csv = io.BytesIO()
        edited_df.to_csv(csv, index=False)
        csv.seek(0)
        st.download_button("📥 Excel", csv, f"diarias_{datetime.now().strftime('%d%m%Y_%H%M')}.csv", "text/csv")
    
    with col2:
        if st.button("💾 Salvar"):
            salvar_dados(edited_df)
            st.success("✅ Salvo!")
            st.rerun()
    
    with col3:
        if st.button("🗑️ Limpar Tudo", type="secondary"):
            dados_diarias = pd.DataFrame(columns=dados_diarias.columns)
            salvar_dados(dados_diarias)
            st.rerun()
else:
    st.info("👆 Cadastre o primeiro registro!")

st.markdown("---")
st.caption("✅ QUANTIDADE POR EXTENSO CORRIGIDA | Código reorganizado")
