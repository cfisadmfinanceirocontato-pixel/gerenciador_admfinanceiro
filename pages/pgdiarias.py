import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import openpyxl
from openpyxl.utils import get_column_letter
from docx import Document
import os
import subprocess
from pathlib import Path
import tempfile
import platform
import shutil

# ============================================================================
# ✅ DETECTAR AMBIENTE (NOVO) - LOCALHOST vs DEPLOY
# ============================================================================
def is_deployed():
    """🔍 Detecta se está em deploy (Streamlit Cloud/HuggingFace/etc.)"""
    try:
        # Verifica se está em ambiente cloud
        return not os.path.exists(str(Path.home() / "Desktop")) or \
               'streamlit-cloud' in st.__version__.lower() or \
               os.environ.get('STREAMLIT_SERVER_HEADLESS', 'false').lower() == 'true'
    except:
        return False

# ============================================================================
# ✅ FUNÇÃO ÁREA DE TRABALHO AUTOMÁTICA (MODIFICADA)
# ============================================================================
def get_output_path():
    """🔄 Detecta pasta correta: Desktop (local) OU Downloads (deploy)"""
    if is_deployed():
        # DEPLOY: Pasta Downloads do sistema
        system = platform.system()
        if system == "Windows":
            return Path.home() / "Downloads" / "Pagto_Diarias"
        elif system == "Darwin":  # macOS
            return Path.home() / "Downloads" / "Pagto_Diarias"
        else:  # Linux
            return Path.home() / "Downloads" / "Pagto_Diarias"
    else:
        # LOCALHOST: Desktop tradicional
        system = platform.system()
        if system == "Windows":
            return Path.home() / "Desktop" / "Pagto_Diarias"
        elif system == "Darwin":  # macOS
            return Path.home() / "Desktop" / "Pagto_Diarias"
        else:  # Linux
            return Path.home() / "Área de Trabalho" / "Pagto_Diarias"

# ============================================================================
# ✅ NOVA FUNÇÃO: PDF NATIVE (SEM LIBREOFFICE) - DEPLOY
# ============================================================================
PDF_NATIVE_AVAILABLE = False
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.pdfgen import canvas
    PDF_NATIVE_AVAILABLE = True
except ImportError:
    PDF_NATIVE_AVAILABLE = False

def docx_to_pdf_native(docx_path, pdf_path):
    """🔥 Converte DOCX → PDF nativamente (sem LibreOffice) - APENAS DEPLOY"""
    if not PDF_NATIVE_AVAILABLE:
        return False
    
    try:
        doc = Document(docx_path)
        c = canvas.Canvas(str(pdf_path), pagesize=letter)
        width, height = letter
        y = height - 50
        
        for para in doc.paragraphs:
            text = para.text
            while text:
                if y < 50:
                    c.showPage()
                    y = height - 50
                line = text[:60]  # Aproximadamente 60 chars por linha
                c.drawString(50, y, line)
                text = text[60:]
                y -= 15
        
        c.save()
        return True
    except:
        return False

# ============================================================================
# ✅ FUNÇÃO GERAR RECIBO INDIVIDUAL (MODIFICADA)
# ============================================================================
def gerar_recibo_individual(dados_registro, template_path, pasta_saida_str):
    """Gera recibo DOCX → PDF para um registro específico - DEPLOY + LOCAL"""
    if not os.path.exists(template_path):
        st.error("❌ **MODELO.docx não encontrado!**")
        return None, None
    
    pasta_saida = Path(pasta_saida_str)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    
    replacements = {
        "(FUNCIONÁRIO)": str(dados_registro.get('Funcionário', '')),
        "(CARGO)": str(dados_registro.get('Cargo', '')),
        "(CPF)": str(dados_registro.get('CPF', '')),
        "(VALOR)": str(dados_registro.get('Valor', '')),
        "((VALOR POR EXTENSO))": str(dados_registro.get('Valor por extenso', '')),
        "(QTD)": str(dados_registro.get('Quantidade', '')),
        "((QTD POR EXTENSO))": str(dados_registro.get('Quantidade por extenso', '')),
        "(NÚMERO DO INSTRUMENTO)": str(dados_registro.get('Instrumento', '')),
        "(TERMO DE COLABORAÇÃO)": str(dados_registro.get('Nº Do Termo de Colaboração', '')),
        "(OBJETIVO)": str(dados_registro.get('Objetivo', '')),
        "(LOCALIDADES)": str(dados_registro.get('Localidades', '')),
        "(PERÍODO)": str(dados_registro.get('Período', '')),
        "(OFÍCIO)": str(dados_registro.get('Ofício', '')),
        "(DATA RECIBO)": str(dados_registro.get('Data por Extenso', ''))
    }
    
    nome_arquivo_recibo = dados_registro.get('Nome do Recibo', 'recibo_sem_nome')
    nome_arquivo_recibo = "".join(c for c in nome_arquivo_recibo if c.isalnum() or c in (' ', '-', '_')).rstrip()
    docx_saida = pasta_saida / f"{nome_arquivo_recibo}.docx"
    pdf_saida = pasta_saida / f"{nome_arquivo_recibo}.pdf"
    
    try:
        doc = Document(template_path)
        for old_text, new_text in replacements.items():
            for paragraph in doc.paragraphs:
                if old_text in paragraph.text:
                    paragraph.text = paragraph.text.replace(old_text, new_text)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if old_text in paragraph.text:
                                paragraph.text = paragraph.text.replace(old_text, new_text)
        
        doc.save(docx_saida)
        
        if not os.path.exists(docx_saida):
            st.error("❌ **Erro: DOCX não foi criado!**")
            return None, None
        
        st.success(f"✅ **DOCX GERADO**: `{docx_saida.name}`")
        
        # ✅ PDF: DEPLOY usa método nativo, LOCAL usa LibreOffice
        pdf_gerado = False
        if is_deployed():
            # DEPLOY: Sempre tenta PDF nativo
            pdf_gerado = docx_to_pdf_native(docx_saida, pdf_saida)
            if pdf_gerado:
                st.success(f"✅ **PDF NATIVE GERADO**: `{pdf_saida.name}`")
            else:
                st.warning("⚠️ **DOCX OK, PDF nativo indisponível**")
        else:
            # LOCALHOST: LibreOffice tradicional
            libreoffice_paths = [
                "C:/Program Files/LibreOffice/program/soffice.exe",
                "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
                shutil.which("libreoffice"),
                shutil.which("soffice")
            ]
            
            libreoffice_cmd = None
            for path in libreoffice_paths:
                if path and os.path.exists(path):
                    libreoffice_cmd = path
                    break
            
            if libreoffice_cmd:
                try:
                    cmd = [
                        libreoffice_cmd, 
                        '--headless', 
                        '--convert-to', 'pdf',
                        '--outdir', str(pasta_saida.absolute()),
                        str(docx_saida.absolute())
                    ]
                    result = subprocess.run(cmd, check=True, capture_output=True, timeout=45)
                    if pdf_saida.exists():
                        st.success(f"✅ **PDF LIBREOFFICE**: `{pdf_saida.name}`")
                        pdf_gerado = True
                    else:
                        st.warning("⚠️ **DOCX OK, mas PDF não encontrado!**")
                except:
                    st.warning("⚠️ **LibreOffice erro** - Apenas DOCX")
            else:
                st.warning("⚠️ **LibreOffice não encontrado!** Apenas DOCX")
        
        return str(pdf_saida) if pdf_gerado else None, str(docx_saida)
        
    except Exception as e:
        st.error(f"❌ **Erro inesperado**: {str(e)[:100]}")
        return None, None

# ============================================================================
# ✅ FUNÇÃO GERAR RECIBOS PARA TODOS OS REGISTROS (MODIFICADA)
# ============================================================================
def gerar_recibos_todos(template_path, pasta_saida_str):
    """🔥 Gera recibos DOCX → PDF para TODOS os registros salvos"""
    if not os.path.exists(template_path):
        st.error("❌ **MODELO.docx não encontrado!**")
        return
    
    caminho_excel = get_output_path() / "registros_completos.xlsx"
    if not os.path.exists(caminho_excel):
        st.error("❌ **Nenhum registro salvo encontrado!**")
        return
    
    try:
        dados_completos = pd.read_excel(caminho_excel)
        if dados_completos.empty:
            st.warning("⚠️ **Nenhum registro para processar**")
            return
        
        pasta_saida = Path(pasta_saida_str)
        pasta_saida.mkdir(parents=True, exist_ok=True)
        
        total_registros = len(dados_completos)
        st.info(f"🚀 **Processando {total_registros} registros...**")
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        pdfs_gerados = 0
        docxs_gerados = 0
        
        for idx, dados_registro in dados_completos.iterrows():
            replacements = {
                "(FUNCIONÁRIO)": str(dados_registro.get('Funcionário', '')),
                "(CARGO)": str(dados_registro.get('Cargo', '')),
                "(CPF)": str(dados_registro.get('CPF', '')),
                "(VALOR)": str(dados_registro.get('Valor', '')),
                "((VALOR POR EXTENSO))": str(dados_registro.get('Valor por extenso', '')),
                "(QTD)": str(dados_registro.get('Quantidade', '')),
                "((QTD POR EXTENSO))": str(dados_registro.get('Quantidade por extenso', '')),
                "(NÚMERO DO INSTRUMENTO)": str(dados_registro.get('Instrumento', '')),
                "(TERMO DE COLABORAÇÃO)": str(dados_registro.get('Nº Do Termo de Colaboração', '')),
                "(OBJETIVO)": str(dados_registro.get('Objetivo', '')),
                "(LOCALIDADES)": str(dados_registro.get('Localidades', '')),
                "(PERÍODO)": str(dados_registro.get('Período', '')),
                "(OFÍCIO)": str(dados_registro.get('Ofício', '')),
                "(DATA RECIBO)": str(dados_registro.get('Data por Extenso', ''))
            }
            
            nome_arquivo_recibo = dados_registro.get('Nome do Recibo', f'recibo_sem_nome_{idx}')
            nome_arquivo_recibo = "".join(c for c in str(nome_arquivo_recibo) if c.isalnum() or c in (' ', '-', '_')).rstrip()
            docx_saida = pasta_saida / f"{nome_arquivo_recibo}.docx"
            pdf_saida = pasta_saida / f"{nome_arquivo_recibo}.pdf"
            
            try:
                doc = Document(template_path)
                for old_text, new_text in replacements.items():
                    for paragraph in doc.paragraphs:
                        if old_text in paragraph.text:
                            paragraph.text = paragraph.text.replace(old_text, new_text)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for paragraph in cell.paragraphs:
                                    if old_text in paragraph.text:
                                        paragraph.text = paragraph.text.replace(old_text, new_text)
                
                doc.save(docx_saida)
                docxs_gerados += 1
                
                # PDF conforme ambiente
                if is_deployed() and PDF_NATIVE_AVAILABLE:
                    if docx_to_pdf_native(docx_saida, pdf_saida):
                        pdfs_gerados += 1
                else:
                    # LibreOffice local
                    libreoffice_paths = [
                        "C:/Program Files/LibreOffice/program/soffice.exe",
                        "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
                        shutil.which("libreoffice"),
                        shutil.which("soffice")
                    ]
                    libreoffice_cmd = next((p for p in libreoffice_paths if p and os.path.exists(p)), None)
                    if libreoffice_cmd:
                        try:
                            cmd = [libreoffice_cmd, '--headless', '--convert-to', 'pdf',
                                   '--outdir', str(pasta_saida.absolute()), str(docx_saida.absolute())]
                            subprocess.run(cmd, check=True, capture_output=True, timeout=30)
                            if pdf_saida.exists():
                                pdfs_gerados += 1
                        except:
                            pass
                
                progress = (idx + 1) / total_registros
                progress_bar.progress(progress)
                status_text.text(f"✅ {idx+1}/{total_registros}: {nome_arquivo_recibo}")
                
            except Exception as e:
                st.error(f"❌ Erro no registro {idx}: {str(e)[:50]}")
                continue
        
        st.success(f"🎉 **PROCESSO CONCLUÍDO!**")
        st.success(f"📄 **DOCX gerados**: {docxs_gerados}")
        st.success(f"📄 **PDFs gerados**: {pdfs_gerados}")
        st.info(f"📂 **Arquivos em**: `{pasta_saida}`")
        
    except Exception as e:
        st.error(f"❌ **Erro geral**: {str(e)}")

# ============================================================================
# ✅ FUNÇÃO PRINCIPAL: SALVAR REGISTRO (MODIFICADA)
# ============================================================================
def salvar_registro_formulario(termo_input, instrumento, numero_termo, funcionario_input, 
                              cpf, cargo, qtd, qtd_extenso, valor, valor_extenso, 
                              data_input, data_extenso_display, objetivo, localidades, 
                              periodo, oficio, nome_arquivo, numero_oficio_input, 
                              nome_recibo_input, pasta_saida_str):
    """✅ Salva registro - Usa pasta correta automaticamente"""
    
    pasta_saida = Path(pasta_saida_str)
    caminho_excel = pasta_saida / "registros_completos.xlsx"
    
    colunas_planilha = [
        'Termo de Colaboração', 'Instrumento', 'Nº Do Termo de Colaboração', 
        'Funcionário', 'CPF', 'Cargo', 'Quantidade', 'Quantidade por extenso', 
        'Valor', 'Valor por extenso', 'Data Recibo', 'Data por Extenso', 
        'Objetivo', 'Localidades', 'Período', 'Ofício', 'Nome Arquivo', 
        'Nº do Ofício', 'Nome do Recibo'
    ]
    
    nome_recibo_completo = nome_recibo_input or f"{nome_arquivo}_{numero_oficio_input}_{st.session_state.contador_recibo}"
    
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
        'Nome do Recibo': nome_recibo_completo
    }
    
    if not termo_input or not funcionario_input or not cpf:
        st.error("❌ **Preencha obrigatoriamente**: Termo, Funcionário e CPF!")
        return None, None, None, None
    
    novo_registro = pd.DataFrame([dados_registro])[colunas_planilha]
    
    try:
        dados_existentes = pd.read_excel(caminho_excel)
        dados_atualizados = pd.concat([dados_existentes, novo_registro], ignore_index=True)
    except FileNotFoundError:
        dados_atualizados = novo_registro
    except Exception as e:
        st.error(f"❌ Erro ao carregar: {e}")
        return None, None, None, None
    
    try:
        with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
            dados_atualizados.to_excel(writer, sheet_name='Registros', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Registros']
            
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
            
            for cell in worksheet[1]:
                cell.font = openpyxl.styles.Font(bold=True)
        
        return dados_atualizados, novo_registro, nome_recibo_completo, dados_registro
        
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {e}")
        return None, None, None, None

# ============================================================================
# FUNÇÕES AUXILIARES (IDENTICAS)
# ============================================================================
def incrementar_contador():
    st.session_state.contador_recibo += 1

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
        return "17 de fevereiro de 2026"
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
        return "17 de fevereiro de 2026"

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
# CONFIGURAÇÃO DA APLICAÇÃO
# ============================================================================
st.set_page_config(
    page_title="Pagto Diárias", 
    page_icon="📋", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

if 'contador_recibo' not in st.session_state:
    st.session_state.contador_recibo = 1

# ============================================================================
# DADOS INICIAIS
# ============================================================================
termos_unicos = carregar_termos_colaboracao()
opcoes_quantidade = ['0,0', '0,5', '1,5', '2,5', '3,5', '4,5']

# ============================================================================
# ✅ PASTA AUTOMÁTICA (GLOBAL)
# ============================================================================
pasta_base = get_output_path()

# ============================================================================
# INTERFACE - SIDEBAR (MODIFICADA)
# ============================================================================
with st.sidebar:
    st.header("🔍 Filtros")
    termo_filtro = st.selectbox("Filtrar por Termo:", ['Todos'] + termos_unicos)
    
    st.markdown("---")
    st.subheader("📄 **Contador Recibo**")
    st.metric("Próximo Nº", st.session_state.contador_recibo)
    
    st.markdown("---")
    st.subheader("📂 **Pasta Automática**")
    ambiente = "🚀 **DEPLOY**" if is_deployed() else "🏠 **LOCALHOST**"
    st.info(f"{ambiente}")
    st.info(f"**Pasta**: `{pasta_base}`")

# ============================================================================
# FORMULÁRIO PRINCIPAL (IDENTICO)
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

st.markdown("**📋 Dados do Termo**")
termo_input = st.selectbox(" **Termo de Colaboração**:", options=[''] + termos_unicos, index=0)
instrumento_auto = buscar_instrumento_por_termo(termo_input) if termo_input else ""
numero_termo_auto = buscar_numero_termo_por_nome(termo_input) if termo_input else ""

col_termo_inst = st.columns([1, 1])
with col_termo_inst[0]:
    instrumento = st.text_input(" **Instrumento**:", value=instrumento_auto)
with col_termo_inst[1]:
    numero_termo = st.text_input(" **Nº Do Termo de Colaboração**:", value=numero_termo_auto)

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

st.markdown("**📋 Detalhes da Viagem**")
objetivo = st.text_area("🎯 **Objetivo**:", height=50)
localidades = st.text_area("📍 **Localidades**:", height=50)

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

nome_recibo_auto = f"{nome_arquivo}_{numero_oficio_input}_{st.session_state.contador_recibo}" if nome_arquivo and numero_oficio_input else ""

with col_periodo_oficio_arquivo[4]:
    nome_recibo = st.text_input("📄 **Nome do Recibo**:", value=nome_recibo_auto)

# ============================================================================
# ✅ 1️⃣ SEÇÃO REGISTROS SALVOS
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros Salvos")

try:
    caminho_excel = pasta_base / "registros_completos.xlsx"
    dados_completos = pd.read_excel(caminho_excel)
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
                colunas_vazias.to_excel(caminho_excel, index=False)
                st.success("✅ **Registros limpos!**")
                st.rerun()
    else:
        st.info("👆 **Cadastre o primeiro registro!**")
except FileNotFoundError:
    st.info("👆 **Cadastre o primeiro registro!**")
except Exception as e:
    st.error(f"❌ Erro: {e}")

# ============================================================================
# ✅ 2️⃣ SEÇÃO GERAR RECIBOS (MODIFICADA)
# ============================================================================
st.markdown("---")
st.subheader("🖨️ **CONFIGURAÇÃO RECIBOS**")

col_template, col_pasta = st.columns([1, 2])
with col_template:
    template_uploaded = st.file_uploader("📄 **Modelo Recibo.docx**", type='docx')

with col_pasta:
    pasta_recibos = st.text_input(
        "📂 **Pasta de Saída**", 
        value=str(pasta_base),
        help="Deploy: Downloads/Pagto_Diarias | Local: Desktop/Pagto_Diarias"
    )
    
    if st.button("✅ **Criar Pasta Automática**", use_container_width=True):
        Path(pasta_recibos).mkdir(parents=True, exist_ok=True)
        st.success(f"✅ **Pasta criada**: `{pasta_recibos}`")
        st.rerun()

# ============================================================================
# ✅ BOTÕES DE AÇÃO (IDENTICOS)
# ============================================================================
st.markdown("---")
st.subheader("⚡ Ações")

col_acoes1, col_acoes2, col_acoes3 = st.columns([3, 2, 3])

with col_acoes1:
    if st.button("💾 **SALVAR & GERAR RECIBO**", type="primary", use_container_width=True):
        resultado = salvar_registro_formulario(
            termo_input, instrumento, numero_termo, funcionario_input, cpf, cargo, 
            qtd, qtd_extenso, valor, valor_extenso, data_input, data_extenso_display, 
            objetivo, localidades, periodo, oficio, nome_arquivo, numero_oficio_input, 
            nome_recibo, pasta_recibos
        )
        
        if resultado[0] is not None:
            dados_registro = resultado[3]
            st.success(f"✅ **REGISTRO SALVO**!")
            st.success(f"📄 **Nome do Recibo**: {resultado[2]}")
            
            if template_uploaded:
                with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_template:
                    tmp_template.write(template_uploaded.read())
                    template_path = tmp_template.name
                
                pdf_gerado, docx_gerado = gerar_recibo_individual(dados_registro, template_path, pasta_recibos)
                if pdf_gerado and docx_gerado:
                    st.success(f"✅ **RECIBO PDF + DOCX GERADOS**!")
                    st.balloons()
                elif docx_gerado:
                    st.success(f"✅ **RECIBO DOCX GERADO**!")
                    st.info("💡 PDF disponível no deploy ou com LibreOffice local")
                os.unlink(template_path)
            else:
                st.info("✅ **Registro salvo!** Carregue modelo DOCX para gerar recibo.")
                st.balloons()
            st.rerun()

with col_acoes2:
    col2_btn1, col2_btn2 = st.columns(2)
    with col2_btn1:
        if st.button("🔄 **ATUALIZAR** ➕", use_container_width=True, on_click=incrementar_contador):
            st.success(f"✅ **Contador**: {st.session_state.contador_recibo}")
            st.rerun()
    with col2_btn2:
        if st.button("🗑️ Limpar", use_container_width=True):
            st.rerun()

with col_acoes3:
    if template_uploaded and st.button("🖨️ **GERAR TODOS OS RECIBOS** 🔥", type="primary", use_container_width=True):
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_template:
            tmp_template.write(template_uploaded.read())
            template_path = tmp_template.name
        
        gerar_recibos_todos(template_path, pasta_recibos)
        os.unlink(template_path)
        st.balloons()
        st.rerun()
    
    if template_uploaded and st.button("🖨️ **GERAR RECIBO ATUAL**", use_container_width=True):
        if all([termo_input, funcionario_input, cpf]):
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
                'Nome do Recibo': nome_recibo
            }
            
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_template:
                tmp_template.write(template_uploaded.read())
                template_path = tmp_template.name
            
            pdf_gerado, docx_gerado = gerar_recibo_individual(dados_registro, template_path, pasta_recibos)
            if pdf_gerado and docx_gerado:
                st.success(f"✅ **PDF + DOCX GERADOS**!")
                st.balloons()
            elif docx_gerado:
                st.success(f"✅ **DOCX GERADO**!")
            os.unlink(template_path)
        else:
            st.error("❌ **Preencha Termo, Funcionário e CPF primeiro!**")

    if st.button("📂 **Abrir Pasta**"):
        if os.path.exists(pasta_recibos):
            system = platform.system()
            try:
                if system == "Windows":
                    os.startfile(pasta_recibos)
                elif system == "Darwin":
                    subprocess.run(["open", pasta_recibos])
                else:
                    subprocess.run(["xdg-open", pasta_recibos])
                st.success("✅ **Pasta aberta!**")
            except:
                st.error("❌ **Erro ao abrir pasta** - Verifique manualmente")
        else:
            st.warning("⚠️ **Crie a pasta primeiro!**")

# ============================================================================
# ESTATÍSTICAS
# ============================================================================
st.markdown("---")
st.subheader("📊 Resumo")
try:
    caminho_excel = pasta_base / "registros_completos.xlsx"
    dados_completos = pd.read_excel(caminho_excel)
    total_registros = len(dados_completos)
    st.metric("📋 Total Registros", f"{total_registros:,}")
    
    if total_registros > 0:
        valores_limpos = dados_completos['Valor'].str.replace('R$', '').str.replace('.', '').str.replace(',', '.').astype(float)
        st.metric("💰 Total Valor", f"R$ {valores_limpos.sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
except:
    st.metric("📋 Total Registros", "0")

st.markdown("---")
st.caption("✅ **CÓDIGO COMPLETO: LOCALHOST + DEPLOY**")
st.caption("🚀 **Deploy**: Downloads/Pagto_Diarias (DOCX+PDF nativo)")
st.caption("🏠 **Local**: Desktop/Pagto_Diarias (DOCX+PDF LibreOffice)")
st.caption("🔥 **ReportLab**: PDF sem LibreOffice no deploy")
