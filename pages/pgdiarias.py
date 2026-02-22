"""
Sistema de Pagamento de Diárias - Versão Híbrida
Funciona em Localhost (com impressão PDF) e Streamlit Cloud (com download)
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import openpyxl
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import subprocess
from pathlib import Path
import tempfile
import platform
import shutil
import requests
import gdown
import zipfile
import time
import base64

# =============================================================================
# 🌐 DETECTAR AMBIENTE (LOCALHOST vs STREAMLIT CLOUD)
# =============================================================================
def is_streamlit_cloud():
    """Detecta se está rodando no Streamlit Cloud"""
    return 'STREAMLIT_SERVER_BASE_URL' in os.environ or 'STREAMLIT_RUNTIME' in os.environ

def is_windows_local():
    """Detecta se é Windows em ambiente local"""
    return platform.system() == "Windows" and not is_streamlit_cloud()

# Importações condicionais para Windows local
WIN32_AVAILABLE = False
if is_windows_local():
    try:
        import win32api
        import win32print
        import win32com.client
        import pythoncom
        WIN32_AVAILABLE = True
    except ImportError:
        WIN32_AVAILABLE = False
        st.warning("⚠️ pywin32 não instalado. A funcionalidade de PDF estará desabilitada. Instale com: pip install pywin32")

# =============================================================================
# 🌐 CONFIGURAÇÕES
# =============================================================================
# ID do template do Google Docs (extraído da URL fornecida)
TEMPLATE_DOC_ID = "1o53p8coWJalAnA6BOt6NJHTboeZoAJwc1L13VO1SN4E"

# =============================================================================
# 📁 CONFIGURAÇÃO DE CAMINHOS
# =============================================================================
def get_base_path():
    """
    Retorna o caminho base para salvar os arquivos
    Adaptado para funcionar em qualquer ambiente
    """
    try:
        if is_streamlit_cloud():
            # No Streamlit Cloud, usa diretório temporário
            pasta_base = Path(tempfile.gettempdir()) / "Pagamento_Diarias"
        else:
            # Local: tenta criar no Desktop
            desktop = Path.home() / "Desktop"
            if desktop.exists():
                pasta_base = desktop / "Pagamento_Diarias"
            else:
                pasta_base = Path(tempfile.gettempdir()) / "Pagamento_Diarias"
        
        pasta_base.mkdir(parents=True, exist_ok=True)
        return pasta_base
    except:
        return Path(tempfile.gettempdir()) / "Pagamento_Diarias"

PASTA_BASE = get_base_path()

def get_csv_path():
    """Retorna o caminho correto para o arquivo CSV"""
    try:
        # Locais comuns para procurar o CSV
        locais_procura = [
            Path.cwd() / "dados_colaboradores.csv",  # Raiz do projeto
            Path(__file__).parent / "dados_colaboradores.csv",  # Mesma pasta do script
            Path(__file__).parent.parent / "dados_colaboradores.csv",  # Pasta pai
        ]
        
        # Adiciona caminhos locais apenas se não estiver no Streamlit Cloud
        if not is_streamlit_cloud():
            locais_procura.extend([
                Path.home() / "Desktop" / "app_streamlit" / "dados_colaboradores.csv",
                Path.home() / "Desktop" / "dados_colaboradores.csv",
            ])
        
        for caminho in locais_procura:
            if caminho.exists():
                return caminho
        
        return locais_procura[0]
    except:
        return Path("dados_colaboradores.csv")

CSV_PATH = get_csv_path()

# =============================================================================
# 📥 FUNÇÃO PARA BAIXAR TEMPLATE
# =============================================================================
def download_template():
    """Baixa o template do recibo usando gdown"""
    try:
        # URL direta para download
        url = f"https://drive.google.com/uc?id={TEMPLATE_DOC_ID}"
        
        # Define caminho para salvar o template
        template_path = PASTA_BASE / "template_recibo.docx"
        
        # Se já existe, usa o existente
        if template_path.exists():
            st.info("📄 Usando template existente")
            return str(template_path)
        
        # Baixa o arquivo
        with st.spinner("📥 Baixando template do recibo..."):
            gdown.download(url, str(template_path), quiet=False)
        
        if template_path.exists() and template_path.stat().st_size > 0:
            st.success("✅ Template baixado com sucesso!")
            return str(template_path)
        else:
            # Tenta método alternativo
            alt_url = f"https://docs.google.com/document/d/{TEMPLATE_DOC_ID}/export?format=docx"
            response = requests.get(alt_url, timeout=30)
            if response.status_code == 200:
                with open(template_path, 'wb') as f:
                    f.write(response.content)
                if template_path.stat().st_size > 0:
                    st.success("✅ Template baixado com sucesso (método alternativo)!")
                    return str(template_path)
            
            st.error("❌ Não foi possível baixar o template")
            return None
            
    except Exception as e:
        st.error(f"❌ Erro ao baixar template: {e}")
        return None

# =============================================================================
# 📁 GERENCIAMENTO DO CSV
# =============================================================================
def carregar_csv_colaboradores():
    """Carrega dados_colaboradores.csv"""
    if not CSV_PATH.exists():
        if is_streamlit_cloud():
            st.warning("📤 **No Streamlit Cloud, faça upload do arquivo CSV:**")
            uploaded_file = st.file_uploader("Carregar dados_colaboradores.csv", type=['csv'], key="csv_uploader")
            if uploaded_file:
                df = pd.read_csv(uploaded_file)
                return df
        else:
            st.error(f"❌ Arquivo dados_colaboradores.csv não encontrado em: {CSV_PATH}")
        
        return pd.DataFrame()

    try:
        # Tentativa com diferentes encodings e separadores
        encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
        separators = [',', ';', '\t']
        
        df = None
        for encoding in encodings:
            for sep in separators:
                try:
                    df = pd.read_csv(CSV_PATH, encoding=encoding, sep=sep)
                    if len(df.columns) > 1:
                        break
                except:
                    continue
            if df is not None and len(df.columns) > 1:
                break
        
        if df is None:
            st.error("❌ Não foi possível ler o arquivo CSV")
            return pd.DataFrame()
        
        # Limpa nomes das colunas
        df.columns = df.columns.str.strip()
        
        return df

    except Exception as e:
        st.error(f"❌ Erro ao carregar CSV: {e}")
        return pd.DataFrame()

# ============================================================================
# ✅ FUNÇÕES DE BUSCA NO CSV
# ============================================================================
@st.cache_data(ttl=300)
def carregar_termos_colaboracao():
    """Carrega lista única de termos de colaboração da terceira coluna do arquivo CSV"""
    df = carregar_csv_colaboradores()
    if df.empty:
        return []

    # Verifica se o DataFrame tem pelo menos 3 colunas
    if len(df.columns) < 3:
        st.warning("⚠️ O arquivo CSV não possui 3 colunas")
        return []
    
    # Pega a terceira coluna (índice 2)
    terceira_coluna = df.columns[2]
    
    # Retorna os valores únicos da terceira coluna
    termos = df[terceira_coluna].dropna().astype(str).unique()
    return sorted(termos)

@st.cache_data(ttl=300)
def buscar_numero_termo_por_termo(termo):
    """
    Busca o número do termo na coluna específica do CSV
    O campo Nº Do Termo de Colaboração recebe a informação da coluna Nº TERMO
    """
    df = carregar_csv_colaboradores()
    if df.empty or not termo:
        return ""

    # Encontra a coluna de termo (TERCEIRA COLUNA)
    if len(df.columns) < 3:
        return ""
    
    col_termo = df.columns[2]  # Terceira coluna

    # Encontra a coluna de número do termo (Nº TERMO)
    col_numero = None
    for col in df.columns:
        if 'Nº' in col.upper() or 'NUMERO' in col.upper() or 'N°' in col.upper():
            col_numero = col
            break
    
    if not col_numero:
        return ""

    # Filtra pelo termo
    mask = df[col_termo].astype(str).str.strip() == str(termo).strip()
    if not mask.any():
        return ""

    # Retorna o número do termo (da coluna Nº TERMO)
    valor = df.loc[mask, col_numero].iloc[0]
    return str(valor).strip() if pd.notna(valor) else ""

@st.cache_data(ttl=300)
def buscar_instrumento_por_termo(termo):
    """Busca o instrumento"""
    df = carregar_csv_colaboradores()
    if df.empty or not termo:
        return ""

    if len(df.columns) < 3:
        return ""
    
    col_termo = df.columns[2]  # Terceira coluna

    mask = df[col_termo].astype(str).str.strip() == str(termo).strip()
    if not mask.any():
        return ""

    if 'INSTRUMENTO' in df.columns:
        valor = df.loc[mask, 'INSTRUMENTO'].iloc[0]
        return str(valor).strip() if pd.notna(valor) else ""
    
    return ""

@st.cache_data(ttl=300)
def carregar_funcionarios_por_termo(termo):
    """Carrega lista de funcionários para um determinado termo"""
    df = carregar_csv_colaboradores()
    if df.empty or not termo:
        return []

    if len(df.columns) < 3:
        return []
    
    col_termo = df.columns[2]  # Terceira coluna

    mask = df[col_termo].astype(str).str.strip() == str(termo).strip()
    if not mask.any():
        return []

    col_func = None
    for col in df.columns:
        if 'FUNCIONÁRIO' in col.upper() or 'NOME' in col.upper():
            col_func = col
            break

    if not col_func:
        return []

    return sorted(df.loc[mask, col_func].dropna().astype(str).unique())

@st.cache_data(ttl=300)
def buscar_cpf_cargo_por_funcionario(termo, funcionario):
    """Busca CPF e cargo do funcionário"""
    df = carregar_csv_colaboradores()
    if df.empty or not termo or not funcionario:
        return "", ""

    if len(df.columns) < 3:
        return "", ""
    
    col_termo = df.columns[2]  # Terceira coluna

    col_func = None
    for col in df.columns:
        if 'FUNCIONÁRIO' in col.upper() or 'NOME' in col.upper():
            col_func = col
            break

    if not col_func:
        return "", ""

    mask = (
        (df[col_termo].astype(str).str.strip() == str(termo).strip()) &
        (df[col_func].astype(str).str.strip() == str(funcionario).strip())
    )

    if not mask.any():
        return "", ""

    linha = df.loc[mask].iloc[0]

    cpf = ""
    for col in df.columns:
        if 'CPF' in col.upper():
            cpf = str(linha[col]).strip() if pd.notna(linha[col]) else ""
            break

    cargo = ""
    for col in df.columns:
        if 'CARGO' in col.upper() or 'FUNÇÃO' in col.upper() or 'FUNCAO' in col.upper():
            cargo = str(linha[col]).strip() if pd.notna(linha[col]) else ""
            break

    return cpf, cargo

# ============================================================================
# ✅ FUNÇÃO PARA IMPRIMIR PDF (APENAS WINDOWS LOCAL)
# ============================================================================
def print_to_pdf(docx_path):
    """
    Converte DOCX para PDF usando o Microsoft Word (ExportAsFixedFormat)
    Baseado no código VBA fornecido
    Funciona apenas em Windows local
    """
    if not is_windows_local():
        return False, "Funcionalidade disponível apenas em Windows local"
    
    if not WIN32_AVAILABLE:
        return False, "pywin32 não instalado. Instale com: pip install pywin32"
    
    try:
        # Inicializa o COM
        pythoncom.CoInitialize()
        
        # Caminho para o PDF de saída
        pdf_path = docx_path.with_suffix('.pdf')
        
        # Abre o Word
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        
        # Abre o documento
        doc = word_app.Documents.Open(str(docx_path))
        
        # Exporta como PDF (equivalente ao ExportAsFixedFormat do VBA)
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_path),
            ExportFormat=17  # wdExportFormatPDF
        )
        
        # Fecha o documento
        doc.Close()
        word_app.Quit()
        
        # Finaliza o COM
        pythoncom.CoUninitialize()
        
        if pdf_path.exists():
            return True, str(pdf_path)
        else:
            return False, "PDF não foi gerado"
            
    except Exception as e:
        # Garante que o COM seja finalizado em caso de erro
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return False, str(e)

# ============================================================================
# ✅ FUNÇÃO PARA GERAR DOCX - VERSÃO SIMPLIFICADA (RECOMENDADA)
# ============================================================================
def gerar_docx_otimizado(dados_registro, template_path):
    """
    Gera DOCX substituindo placeholders no template
    Versão simplificada sem dependência da função Cells
    """
    
    # Mapeamento direto dos placeholders para os valores
    replacements = {
        "(TERMO DE COLABORAÇÃO)": str(dados_registro.get('Termo de Colaboração', '')),
        "(INSTRUMENTO)": str(dados_registro.get('Instrumento', '')),
        "(Nº DO TERMO DE COLABORAÇÃO)": str(dados_registro.get('Nº Do Termo de Colaboração', '')),
        "(FUNCIONÁRIO)": str(dados_registro.get('Funcionário', '')),
        "(CPF)": str(dados_registro.get('CPF', '')),
        "(CARGO)": str(dados_registro.get('Cargo', '')),
        "(QTD)": str(dados_registro.get('Quantidade', '')),
        "((QTD POR EXTENSO))": str(dados_registro.get('Quantidade por extenso', '')),
        "(VALOR)": str(dados_registro.get('Valor', '')),
        "((VALOR POR EXTENSO))": str(dados_registro.get('Valor por extenso', '')),
        "(DATA RECIBO)": str(dados_registro.get('Data Recibo', '')),
        "(DATA POR EXTENSO)": str(dados_registro.get('Data por Extenso', '')),
        "(OBJETIVO)": str(dados_registro.get('Objetivo', '')),
        "(LOCALIDADES)": str(dados_registro.get('Localidades', '')),
        "(PERÍODO)": str(dados_registro.get('Período', '')),
        "(OFÍCIO)": str(dados_registro.get('Ofício', ''))
    }
    
    # Nome do arquivo
    nome_arquivo_recibo = dados_registro.get('Nome do Recibo', 'recibo_sem_nome')
    nome_arquivo_recibo = "".join(c for c in nome_arquivo_recibo if c.isalnum() or c in (' ', '-', '_')).rstrip()
    
    # Caminho do arquivo DOCX
    docx_path = PASTA_BASE / f"{nome_arquivo_recibo}.docx"
    
    try:
        # Carrega o template
        doc = Document(template_path)
        
        # Substitui placeholders em todos os parágrafos
        for paragraph in doc.paragraphs:
            for placeholder, valor in replacements.items():
                if placeholder in paragraph.text:
                    # Substitui mantendo a formatação original
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, valor)
        
        # Substitui placeholders em todas as tabelas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for placeholder, valor in replacements.items():
                            if placeholder in paragraph.text:
                                for run in paragraph.runs:
                                    if placeholder in run.text:
                                        run.text = run.text.replace(placeholder, valor)
        
        # Salva o documento
        doc.save(docx_path)
        
        return {
            'docx': str(docx_path),
            'nome': nome_arquivo_recibo,
            'sucesso': True
        }
        
    except Exception as e:
        return {
            'sucesso': False,
            'erro': str(e)
        }

# ============================================================================
# ✅ FUNÇÃO GERAR RECIBO INDIVIDUAL (VERSÃO HÍBRIDA)
# ============================================================================
def gerar_recibo_individual(dados_registro, template_path, gerar_pdf=False):
    """
    Gera recibo DOCX e opcionalmente PDF
    Versão híbrida que funciona em qualquer ambiente
    """
    
    # Primeiro gera o DOCX
    resultado = gerar_docx_otimizado(dados_registro, template_path)
    
    if not resultado['sucesso']:
        st.error(f"❌ Erro ao gerar DOCX: {resultado.get('erro', 'Erro desconhecido')}")
        return None
    
    docx_path = resultado['docx']
    
    # Prepara o resultado
    resultado_final = {
        'docx': docx_path,
        'nome': resultado['nome'],
        'pdf': None
    }
    
    # Se solicitou PDF e está em Windows local, tenta gerar
    if gerar_pdf and is_windows_local() and WIN32_AVAILABLE:
        with st.spinner("🖨️ Gerando PDF via Microsoft Word..."):
            sucesso, pdf_info = print_to_pdf(Path(docx_path))
            if sucesso:
                resultado_final['pdf'] = pdf_info
                st.success("✅ PDF gerado com sucesso!")
            else:
                st.warning(f"⚠️ Não foi possível gerar PDF: {pdf_info}")
    
    return resultado_final

# ============================================================================
# ✅ FUNÇÃO SALVAR REGISTRO
# ============================================================================
def salvar_registro_formulario(termo_input, instrumento, numero_termo, funcionario_input, 
                              cpf, cargo, qtd, qtd_extenso, valor, valor_extenso, 
                              data_input, data_extenso_display, objetivo, localidades, 
                              periodo, oficio, nome_arquivo, numero_oficio_input, 
                              nome_recibo_input):
    """Salva registro na planilha local"""
    
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
    
    # Caminho do arquivo Excel
    excel_path = PASTA_BASE / "registros_completos.xlsx"
    
    try:
        if excel_path.exists():
            df_existente = pd.read_excel(excel_path)
            df_atualizado = pd.concat([df_existente, novo_registro], ignore_index=True)
        else:
            df_atualizado = novo_registro
        
        # Salva Excel
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_atualizado.to_excel(writer, sheet_name='Registros', index=False)
            
            # Ajusta largura das colunas
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
        
        return df_atualizado, novo_registro, nome_recibo_completo, excel_path
        
    except Exception as e:
        st.error(f"❌ Erro ao salvar Excel: {e}")
        return None, None, None, None

# ============================================================================
# ✅ FUNÇÃO GERAR TODOS OS RECIBOS
# ============================================================================
def gerar_todos_recibos(template_path, gerar_pdf=False):
    """Gera recibos para todos os registros salvos"""
    
    excel_path = PASTA_BASE / "registros_completos.xlsx"
    
    if not excel_path.exists():
        st.error("❌ Nenhum registro encontrado para gerar recibos!")
        return
    
    try:
        df = pd.read_excel(excel_path)
        if df.empty:
            st.warning("⚠️ Nenhum registro para processar")
            return
        
        total_registros = len(df)
        st.info(f"🚀 Gerando {total_registros} recibos...")
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        gerados = 0
        erros = 0
        pdfs_gerados = 0
        
        for idx, row in df.iterrows():
            try:
                dados_registro = {
                    'Termo de Colaboração': row.get('Termo de Colaboração', ''),
                    'Instrumento': row.get('Instrumento', ''),
                    'Nº Do Termo de Colaboração': row.get('Nº Do Termo de Colaboração', ''),
                    'Funcionário': row.get('Funcionário', ''),
                    'CPF': row.get('CPF', ''),
                    'Cargo': row.get('Cargo', ''),
                    'Quantidade': row.get('Quantidade', ''),
                    'Quantidade por extenso': row.get('Quantidade por extenso', ''),
                    'Valor': row.get('Valor', ''),
                    'Valor por extenso': row.get('Valor por extenso', ''),
                    'Data Recibo': row.get('Data Recibo', ''),
                    'Data por Extenso': row.get('Data por Extenso', ''),
                    'Objetivo': row.get('Objetivo', ''),
                    'Localidades': row.get('Localidades', ''),
                    'Período': row.get('Período', ''),
                    'Ofício': row.get('Ofício', ''),
                    'Nome do Recibo': row.get('Nome do Recibo', f'recibo_{idx}')
                }
                
                resultado = gerar_recibo_individual(dados_registro, template_path, gerar_pdf)
                
                if resultado:
                    gerados += 1
                    if resultado.get('pdf'):
                        pdfs_gerados += 1
                else:
                    erros += 1
                
                progress = (idx + 1) / total_registros
                progress_bar.progress(progress)
                status_text.text(f"✅ {idx+1}/{total_registros}: {dados_registro['Nome do Recibo']}")
                
            except Exception as e:
                erros += 1
                st.error(f"Erro no registro {idx}: {e}")
        
        if gerar_pdf and is_windows_local():
            st.success(f"✅ Processo concluído! {gerados} DOCX gerados, {pdfs_gerados} PDF gerados, {erros} erros.")
        else:
            st.success(f"✅ Processo concluído! {gerados} DOCX gerados, {erros} erros.")
        
    except Exception as e:
        st.error(f"❌ Erro ao processar recibos: {e}")

# ============================================================================
# FUNÇÕES AUXILIARES (FORMATAÇÃO)
# ============================================================================
def formatar_data_completa(data_obj):
    if pd.isna(data_obj) or data_obj == '':
        return datetime.now().strftime("%d de %B de %Y").lower()
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
        return datetime.now().strftime("%d de %B de %Y").lower()

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

def incrementar_contador():
    st.session_state.contador_recibo += 1

def resetar_contador_funcionario(funcionario_anterior, funcionario_atual):
    if funcionario_anterior != funcionario_atual and funcionario_atual:
        st.session_state.contador_recibo = 1

def extrair_numero_oficio(oficio_completo):
    if pd.isna(oficio_completo) or not isinstance(oficio_completo, str) or '/' not in oficio_completo:
        return str(oficio_completo).strip()
    return oficio_completo.split('/')[0].strip()

def resetar_formulario():
    st.session_state.contador_recibo = 1
    st.session_state.funcionario_anterior = ""
    st.rerun()

# ============================================================================
# CONFIGURAÇÃO DA APLICAÇÃO
# ============================================================================
st.set_page_config(
    page_title="Gerar Recibos de Diarias", 
    page_icon="📋", 
    layout="wide"
)

# ============================================================================
# SESSION STATE
# ============================================================================
if 'contador_recibo' not in st.session_state:
    st.session_state.contador_recibo = 1
if 'funcionario_anterior' not in st.session_state:
    st.session_state.funcionario_anterior = ""
if 'template_path' not in st.session_state:
    st.session_state.template_path = None

# ============================================================================
# INICIALIZAÇÃO
# ============================================================================
termos_unicos = carregar_termos_colaboracao()
opcoes_quantidade = ['0,0', '0,5', '1,5', '2,5', '3,5', '4,5']

# Baixa template se necessário
if not st.session_state.template_path:
    st.session_state.template_path = download_template()

# ============================================================================
# SIDEBAR
# ============================================================================
with st.sidebar:
    st.header("🔍 Filtros")
    termo_filtro = st.selectbox("Filtrar por Termo:", ['Todos'] + termos_unicos)
    
    st.markdown("---")
    st.subheader("📄 **Contador Recibo**")
    st.metric("Próximo Nº", st.session_state.contador_recibo)
    
    st.markdown("---")
    st.subheader("📂 **Informações**")
    
    # Mostra ambiente atual
    if is_streamlit_cloud():
        st.info("🚀 **Ambiente:** Streamlit Cloud")
        st.info("📥 **Download dos arquivos disponível**")
    else:
        st.info("💻 **Ambiente:** Local")
        if is_windows_local() and WIN32_AVAILABLE:
            st.success("✅ **Impressão PDF disponível**")
        else:
            st.warning("⚠️ **Impressão PDF indisponível**")
    
    st.info(f"📁 **Pasta de trabalho:**\n{PASTA_BASE}")
    
    if CSV_PATH.exists():
        st.success(f"✅ CSV encontrado: {CSV_PATH.name}")
    else:
        if is_streamlit_cloud():
            st.warning("📤 **Faça upload do CSV**")
        else:
            st.error(f"❌ CSV não encontrado")
    
    # Botão para criar ZIP de todos os arquivos
    if st.button("📦 **Criar ZIP de todos os arquivos**", use_container_width=True):
        with st.spinner("Criando arquivo ZIP..."):
            zip_path = PASTA_BASE / "todos_arquivos.zip"
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for file in PASTA_BASE.glob("*"):
                    if file.is_file() and file.name != "todos_arquivos.zip":
                        zipf.write(file, file.name)
            
            with open(zip_path, 'rb') as f:
                zip_data = f.read()
            
            st.download_button(
                "📥 **Clique para baixar ZIP**",
                zip_data,
                file_name="pagamento_diarias.zip",
                mime="application/zip",
                key="download_zip"
            )

# ============================================================================
# VERIFICAÇÕES
# ============================================================================
if not st.session_state.template_path:
    st.error("❌ Template não disponível. Verifique sua conexão.")
    if st.button("🔄 Tentar novamente"):
        st.rerun()
    st.stop()

# ============================================================================
# FORMULÁRIO PRINCIPAL
# ============================================================================
st.title("📋 Gerar Recibos de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

# Dados do Termo
st.markdown("**📋 Dados do Termo**")
# A lista suspensa do campo Termo de Colaboração busca as informações contidas na terceira coluna do arquivo dados_colaboradores.csv
termo_input = st.selectbox(" **Termo de Colaboração**:", options=[''] + termos_unicos, index=0)

# Busca automática do número do termo quando um termo é selecionado
if termo_input:
    numero_termo_auto = buscar_numero_termo_por_termo(termo_input)
    instrumento_auto = buscar_instrumento_por_termo(termo_input)
else:
    numero_termo_auto = ""
    instrumento_auto = ""

col_termo_inst = st.columns([1, 1])
with col_termo_inst[0]:
    instrumento = st.text_input(" **Instrumento**:", value=instrumento_auto)
with col_termo_inst[1]:
    numero_termo = st.text_input(" **Nº Do Termo de Colaboração**:", value=numero_termo_auto)

# Dados do Funcionário
st.markdown("**👤 Dados do Funcionário**")
funcionarios_do_termo = carregar_funcionarios_por_termo(termo_input)
funcionario_input = st.selectbox("👤 **Funcionário**", options=[''] + funcionarios_do_termo, index=0)

if st.session_state.funcionario_anterior != funcionario_input:
    resetar_contador_funcionario(st.session_state.funcionario_anterior, funcionario_input)
    st.session_state.funcionario_anterior = funcionario_input

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
    qtd = st.selectbox("Quantidade:", options=opcoes_quantidade, index=2)
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

nome_recibo_auto = f"{nome_arquivo}_{numero_oficio_input}_{st.session_state.contador_recibo}" if nome_arquivo and numero_oficio_input else ""

with col_periodo_oficio_arquivo[4]:
    nome_recibo = st.text_input("📄 **Nome do Recibo**:", value=nome_recibo_auto)

# ============================================================================
# BOTÕES DE AÇÃO
# ============================================================================
st.markdown("---")

# Layout adaptativo conforme ambiente
if is_windows_local() and WIN32_AVAILABLE:
    # Local com impressão PDF
    col_acoes = st.columns([1, 1, 1, 1])
    
    with col_acoes[0]:
        if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
            if not all([termo_input, funcionario_input, cpf]):
                st.error("❌ **Preencha Termo, Funcionário e CPF!**")
            else:
                resultado = salvar_registro_formulario(
                    termo_input, instrumento, numero_termo, funcionario_input, cpf, cargo,
                    qtd, qtd_extenso, valor, valor_extenso, data_input, data_extenso_display,
                    objetivo, localidades, periodo, oficio, nome_arquivo, numero_oficio_input,
                    nome_recibo
                )
                
                if resultado and resultado[0] is not None:
                    st.success(f"✅ **Registro salvo!**")
                    st.rerun()
    
    with col_acoes[1]:
        if st.button("🖨️ **GERAR DOCX**", use_container_width=True):
            if not all([termo_input, funcionario_input, cpf]):
                st.error("❌ **Preencha Termo, Funcionário e CPF!**")
            else:
                dados_registro = {
                    'Termo de Colaboração': termo_input,
                    'Instrumento': instrumento,
                    'Nº Do Termo de Colaboração': numero_termo,
                    'Funcionário': funcionario_input,
                    'CPF': cpf,
                    'Cargo': cargo,
                    'Quantidade': qtd,
                    'Quantidade por extenso': qtd_extenso,
                    'Valor': valor,
                    'Valor por extenso': valor_extenso,
                    'Data Recibo': data_input,
                    'Data por Extenso': data_extenso_display,
                    'Objetivo': objetivo,
                    'Localidades': localidades,
                    'Período': periodo,
                    'Ofício': oficio,
                    'Nome do Recibo': nome_recibo
                }
                
                resultado = gerar_recibo_individual(dados_registro, st.session_state.template_path, gerar_pdf=False)
                
                if resultado:
                    st.success("✅ **DOCX gerado com sucesso!**")
                    st.info(f"📁 **Salvo em:** {PASTA_BASE}")
    
    with col_acoes[2]:
        if st.button("📄 **GERAR DOCX + PDF**", use_container_width=True):
            if not all([termo_input, funcionario_input, cpf]):
                st.error("❌ **Preencha Termo, Funcionário e CPF!**")
            else:
                dados_registro = {
                    'Termo de Colaboração': termo_input,
                    'Instrumento': instrumento,
                    'Nº Do Termo de Colaboração': numero_termo,
                    'Funcionário': funcionario_input,
                    'CPF': cpf,
                    'Cargo': cargo,
                    'Quantidade': qtd,
                    'Quantidade por extenso': qtd_extenso,
                    'Valor': valor,
                    'Valor por extenso': valor_extenso,
                    'Data Recibo': data_input,
                    'Data por Extenso': data_extenso_display,
                    'Objetivo': objetivo,
                    'Localidades': localidades,
                    'Período': periodo,
                    'Ofício': oficio,
                    'Nome do Recibo': nome_recibo
                }
                
                resultado = gerar_recibo_individual(dados_registro, st.session_state.template_path, gerar_pdf=True)
                
                if resultado:
                    st.success("✅ **DOCX e PDF gerados com sucesso!**")
                    st.info(f"📁 **Arquivos em:** {PASTA_BASE}")
    
    with col_acoes[3]:
        if st.button("🔄 **INCREMENTAR**", use_container_width=True, on_click=incrementar_contador):
            st.success(f"✅ **Contador**: {st.session_state.contador_recibo}")
    
else:
    # Streamlit Cloud ou Windows sem pywin32 - apenas DOCX
    col_acoes = st.columns([1, 1, 1])
    
    with col_acoes[0]:
        if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
            if not all([termo_input, funcionario_input, cpf]):
                st.error("❌ **Preencha Termo, Funcionário e CPF!**")
            else:
                resultado = salvar_registro_formulario(
                    termo_input, instrumento, numero_termo, funcionario_input, cpf, cargo,
                    qtd, qtd_extenso, valor, valor_extenso, data_input, data_extenso_display,
                    objetivo, localidades, periodo, oficio, nome_arquivo, numero_oficio_input,
                    nome_recibo
                )
                
                if resultado and resultado[0] is not None:
                    st.success(f"✅ **Registro salvo!**")
                    st.rerun()
    
    with col_acoes[1]:
        if st.button("🖨️ **GERAR DOCX**", use_container_width=True):
            if not all([termo_input, funcionario_input, cpf]):
                st.error("❌ **Preencha Termo, Funcionário e CPF!**")
            else:
                dados_registro = {
                    'Termo de Colaboração': termo_input,
                    'Instrumento': instrumento,
                    'Nº Do Termo de Colaboração': numero_termo,
                    'Funcionário': funcionario_input,
                    'CPF': cpf,
                    'Cargo': cargo,
                    'Quantidade': qtd,
                    'Quantidade por extenso': qtd_extenso,
                    'Valor': valor,
                    'Valor por extenso': valor_extenso,
                    'Data Recibo': data_input,
                    'Data por Extenso': data_extenso_display,
                    'Objetivo': objetivo,
                    'Localidades': localidades,
                    'Período': periodo,
                    'Ofício': oficio,
                    'Nome do Recibo': nome_recibo
                }
                
                resultado = gerar_recibo_individual(dados_registro, st.session_state.template_path, gerar_pdf=False)
                
                if resultado:
                    st.success("✅ **DOCX gerado com sucesso!**")
                    st.info(f"📁 **Salvo em:** {PASTA_BASE}")
    
    with col_acoes[2]:
        if st.button("🔄 **INCREMENTAR**", use_container_width=True, on_click=incrementar_contador):
            st.success(f"✅ **Contador**: {st.session_state.contador_recibo}")

# ============================================================================
# SEÇÃO PARA GERAR TODOS OS RECIBOS
# ============================================================================
st.markdown("---")
st.subheader("📦 Operações em Lote")

if is_windows_local() and WIN32_AVAILABLE:
    col_lote1, col_lote2, col_lote3 = st.columns(3)
    
    with col_lote1:
        if st.button("📑 **GERAR TODOS DOCX**", use_container_width=True):
            gerar_todos_recibos(st.session_state.template_path, gerar_pdf=False)
    
    with col_lote2:
        if st.button("🖨️ **GERAR TODOS DOCX + PDF**", use_container_width=True):
            gerar_todos_recibos(st.session_state.template_path, gerar_pdf=True)
    
    with col_lote3:
        if st.button("🔄 **RESETAR FORMULÁRIO**", use_container_width=True):
            resetar_formulario()
else:
    col_lote1, col_lote2 = st.columns(2)
    
    with col_lote1:
        if st.button("📑 **GERAR TODOS DOCX**", use_container_width=True):
            gerar_todos_recibos(st.session_state.template_path, gerar_pdf=False)
    
    with col_lote2:
        if st.button("🔄 **RESETAR FORMULÁRIO**", use_container_width=True):
            resetar_formulario()

# ============================================================================
# REGISTROS SALVOS
# ============================================================================
st.markdown("---")
st.subheader("📋 Registros Salvos")

excel_path = PASTA_BASE / "registros_completos.xlsx"
if excel_path.exists():
    try:
        df_existente = pd.read_excel(excel_path)
        if not df_existente.empty:
            st.dataframe(df_existente, use_container_width=True)
            
            col1, col2 = st.columns(2)
            with col1:
                total = len(df_existente)
                st.metric("📋 Total Registros", total)
            
            with col2:
                if 'Valor' in df_existente.columns:
                    try:
                        valores = df_existente['Valor'].str.replace('R$', '').str.replace('.', '').str.replace(',', '.').astype(float)
                        st.metric("💰 Total", f"R$ {valores.sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                    except:
                        pass
            
            # Botão de download do Excel
            with open(excel_path, 'rb') as f:
                st.download_button(
                    "📥 **Download Excel**",
                    f,
                    file_name="registros_completos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except:
        st.info("📂 **Arquivo Excel encontrado mas não foi possível ler**")
else:
    st.info("👆 **Nenhum registro encontrado. Cadastre o primeiro!**")

# ============================================================================
# RODAPÉ
# ============================================================================
st.markdown("---")
st.caption(f"✅ **Arquivos salvos em:** {PASTA_BASE}")
st.caption("📌 **DOCX gerado conforme modelo original**")
if is_windows_local() and WIN32_AVAILABLE:
    st.caption("🖨️ **PDF gerado via Microsoft Word (ExportAsFixedFormat)**")
else:
    st.caption("📥 **Download dos arquivos disponível via botão ZIP**")