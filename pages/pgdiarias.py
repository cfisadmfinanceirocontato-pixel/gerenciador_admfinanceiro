import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import openpyxl
from openpyxl.utils import get_column_letter

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
# ✅ SESSION STATE PARA CONTADOR NUMÉRICO
# ============================================================================
if 'contador_recibo' not in st.session_state:
    st.session_state.contador_recibo = 1

# ============================================================================
# ✅ FUNÇÃO PRINCIPAL: SALVAR REGISTRO NA PLANILHA EXCEL
# ============================================================================
def salvar_registro_formulario():
    """✅ Nome do Recibo: NomeArquivo_NºOfício_Contador"""
    
    # ✅ 19 COLUNAS NA ORDEM EXATA
    colunas_planilha = [
        'Termo de Colaboração', 'Instrumento', 'Nº Do Termo de Colaboração', 
        'Funcionário', 'CPF', 'Cargo', 'Quantidade', 'Quantidade por extenso', 
        'Valor', 'Valor por extenso', 'Data Recibo', 'Data por Extenso', 
        'Objetivo', 'Localidades', 'Período', 'Ofício', 'Nome Arquivo', 
        'Nº do Ofício', 'Nome do Recibo'
    ]
    
    # ✅ NOME DO RECIBO: NomeArquivo_NºOfício_1, NomeArquivo_NºOfício_2, etc.
    nome_recibo_completo = f"{nome_arquivo}_{numero_oficio_input}_{st.session_state.contador_recibo}"
    
    dados_registro = {
        'Termo de Colaboração': termo_input or '',
        'Instrumento': instrumento or '',
        'Nº Do Termo de Colaboração': numero_termo or '',
        'Funcionário': funcionario_input or '',
        'CPF': cpf or '',
        'Cargo': cargo or '',
        'Quantidade': qtd or '',
        'Quantidade por extenso': qtd_extenso or '',
        'Valor': valor or '',
        'Valor por extenso': valor_extenso or '',
        'Data Recibo': data_input or '',
        'Data por Extenso': data_extenso_display or '',
        'Objetivo': objetivo or '',
        'Localidades': localidades or '',
        'Período': periodo or '',
        'Ofício': oficio or '',
        'Nome Arquivo': nome_arquivo or '',
        'Nº do Ofício': numero_oficio_input or '',
        'Nome do Recibo': nome_recibo_completo  # ✅ FORMATO FINAL
    }
    
    # ✅ VALIDAÇÃO
    if not termo_input or not funcionario_input or not cpf:
        st.error("❌ **Preencha obrigatoriamente**: Termo, Funcionário e CPF!")
        return None, None
    
    novo_registro = pd.DataFrame([dados_registro])[colunas_planilha]
    
    # ✅ CARREGA OU CRIA PLANILHA
    try:
        dados_existentes = pd.read_excel("registros_completos.xlsx")
        dados_atualizados = pd.concat([dados_existentes, novo_registro], ignore_index=True)
    except FileNotFoundError:
        dados_atualizados = novo_registro
    except Exception as e:
        st.error(f"❌ Erro ao carregar: {e}")
        return None, None
    
    # ✅ SALVA COM FORMATAÇÃO
    try:
        with pd.ExcelWriter("registros_completos.xlsx", engine='openpyxl') as writer:
            dados_atualizados.to_excel(writer, sheet_name='Registros', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Registros']
            
            # Auto-ajuste colunas
            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Cabeçalho negrito
            for cell in worksheet[1]:
                cell.font = openpyxl.styles.Font(bold=True)
        
        return dados_atualizados, novo_registro, nome_recibo_completo
        
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {e}")
        return None, None, None

# ============================================================================
# FUNÇÃO INCREMENTAR CONTADOR
# ============================================================================
def incrementar_contador():
    """➕ Incrementa contador +1"""
    st.session_state.contador_recibo += 1

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================
def extrair_numero_oficio(oficio_completo):
    if pd.isna(oficio_completo) or not isinstance(oficio_completo, str) or '/' not in oficio_completo:
        return str(oficio_completo).strip()
    return oficio_completo.split('/')[0].strip()

@st.cache_data(ttl=300)
def carregar_termos_colaboracao():
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        termos = sorted(df_colab['TERMO DE COLABORAÇÃO'].dropna().astype(str).unique())
        return termos
    except:
        return ['TERMO1', 'TERMO2']

@st.cache_data(ttl=300)
def buscar_instrumento_por_termo(termo):
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
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 1:
            coluna_numero = df_colab.columns[1]
            mask = df_colab['TERMO DE COLABORAÇÃO'] == termo
            if mask.any():
                return str(df_colab.loc[mask, coluna_numero].iloc[0]).strip()
        return ""
    except:
        return ""

@st.cache_data(ttl=300)
def carregar_funcionarios_por_termo(termo):
    try:
        df_colab = pd.read_csv("dados_colaboradores.csv")
        if len(df_colab.columns) > 6:
            coluna_oitava = df_colab.columns[6]
            mask = df_colab['TERMO DE COLABORAÇÃO'] == termo
            if mask.any():
                return sorted(df_colab.loc[mask, coluna_oitava].dropna().astype(str).unique())
        return []
    except:
        return []

@st.cache_data(ttl=300)
def buscar_cpf_cargo_por_funcionario(termo, funcionario):
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
    
    st.markdown("---")
    st.subheader("📄 **Contador Recibo**")
    st.metric("Próximo Nº", st.session_state.contador_recibo)

# ============================================================================
# FORMULÁRIO PRINCIPAL
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

# Dados do Termo
st.markdown("**📋 Dados do Termo**")
termo_input = st.selectbox(" **Termo de Colaboração**:", options=[''] + termos_unicos, index=0)
instrumento_auto = buscar_instrumento_por_termo(termo_input) if termo_input else ""
numero_termo_auto = buscar_numero_termo_por_nome(termo_input) if termo_input else ""

col_termo_inst = st.columns([1, 1])
with col_termo_inst[0]:
    instrumento = st.text_input(" **Instrumento**:", value=instrumento_auto)
with col_termo_inst[1]:
    numero_termo = st.text_input(" **Nº Do Termo de Colaboração**:", value=numero_termo_auto)

# Dados do Funcionário
st.markdown("**👤 Dados do Funcionário**")
funcionarios_do_termo = carregar_funcionarios_por_termo(termo_input)
funcionario_input = st.selectbox("👤 **Funcionário**", options=[''] + funcionarios_do_termo, index=0)

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

# ✅ NOME DO RECIBO: NomeArquivo_NºOfício_Contador
nome_recibo_auto = f"{nome_arquivo}_{numero_oficio_input}_{st.session_state.contador_recibo}" if nome_arquivo and numero_oficio_input else ""

with col_periodo_oficio_arquivo[4]:
    nome_recibo = st.text_input("📄 **Nome do Recibo**:", value=nome_recibo_auto)

# ============================================================================
# BOTÕES DE AÇÃO
# ============================================================================
st.markdown("---")
st.subheader("⚡ Ações")

col_acoes1, col_acoes2, col_acoes3 = st.columns([3, 2, 3])

with col_acoes1:
    if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
        resultado = salvar_registro_formulario()
        if resultado[0] is not None:
            st.success(f"✅ **REGISTRO SALVO!**")
            st.success(f"📄 **Nome do Recibo**: {resultado[2]}")
            st.success(f"🔢 **Próximo**: {nome_arquivo}_{numero_oficio_input}_{st.session_state.contador_recibo + 1}")
            st.balloons()
            st.rerun()

with col_acoes2:
    col2_btn1, col2_btn2 = st.columns(2)
    with col2_btn1:
        if st.button("🔄 **ATUALIZAR** ➕", use_container_width=True, on_click=incrementar_contador):
            st.success(f"✅ **Contador atualizado**: {st.session_state.contador_recibo}")
            st.rerun()
    with col2_btn2:
        if st.button("🗑️ Limpar", use_container_width=True):
            st.rerun()

# ============================================================================
# TABELA DE REGISTROS
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros Salvos")

try:
    dados_completos = pd.read_excel("registros_completos.xlsx")
    if not dados_completos.empty:
        st.dataframe(dados_completos, use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                dados_completos.to_excel(writer, sheet_name='Registros', index=False)
            buffer_excel.seek(0)
            st.download_button(
                "📥 **Download Excel**", 
                buffer_excel.getvalue(), 
                f"registros_completos_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx", 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col2:
            if st.button("🗑️ **Limpar Tudo**", type="secondary", use_container_width=True):
                colunas_vazias = pd.DataFrame(columns=dados_completos.columns)
                colunas_vazias.to_excel("registros_completos.xlsx", index=False)
                st.rerun()
    else:
        st.info("👆 **Cadastre o primeiro registro!**")
except FileNotFoundError:
    st.info("👆 **Cadastre o primeiro registro!**")
except Exception as e:
    st.error(f"❌ Erro: {e}")

# ============================================================================
# ESTATÍSTICAS
# ============================================================================
st.markdown("---")
st.subheader("📊 Resumo")
try:
    dados_completos = pd.read_excel("registros_completos.xlsx")
    total_registros = len(dados_completos)
    st.metric("📋 Total Registros", f"{total_registros:,}")
    
    if total_registros > 0:
        valores_limpos = dados_completos['Valor'].str.replace('R$', '').str.replace('.', '').str.replace(',', '.').astype(float)
        st.metric("💰 Total Valor", f"R$ {valores_limpos.sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
except:
    st.metric("📋 Total Registros", "0")

st.markdown("---")
st.caption("✅ **Formato Nome Recibo**: `NomeArquivo_NºOfício_1` ➕ `NomeArquivo_NºOfício_2` ➕ etc.")
