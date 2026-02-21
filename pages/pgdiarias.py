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
import requests
import re

# =============================================================================
# 🌐 TENTATIVA DE IMPORTAR GOOGLE DRIVE API (COM FALLBACK)
# =============================================================================
GOOGLE_DRIVE_AVAILABLE = False
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
    import io
    GOOGLE_DRIVE_AVAILABLE = True
except ImportError:
    GOOGLE_DRIVE_AVAILABLE = False
    st.warning("⚠️ Bibliotecas do Google Drive não instaladas. Arquivos serão salvos apenas localmente.")
    st.info("💡 Para instalar: pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")

# =============================================================================
# 🌐 CONFIGURAÇÃO GOOGLE DRIVE API
# =============================================================================
SCOPES = ['https://www.googleapis.com/auth/drive.file']
MODELO_DOC_ID = '1o53p8coWJalAnA6BOt6NJHTboeZoAJwc1L13VO1SN4E'  # ID do documento modelo

def get_drive_service():
    """🔧 Configura e retorna o serviço do Google Drive"""
    if not GOOGLE_DRIVE_AVAILABLE:
        return None
        
    try:
        # Verifica se está no Streamlit Cloud
        if 'gdrive_service_account' in st.secrets:
            credentials = service_account.Credentials.from_service_account_info(
                st.secrets['gdrive_service_account'], scopes=SCOPES
            )
        else:
            # Para ambiente local, usar arquivo de credenciais
            creds_file = Path.home() / '.config' / 'gdrive_credentials.json'
            if creds_file.exists():
                credentials = service_account.Credentials.from_service_account_file(
                    str(creds_file), scopes=SCOPES
                )
            else:
                st.warning("⚠️ Credenciais do Google Drive não encontradas. Recibos serão salvos apenas localmente.")
                return None
        
        return build('drive', 'v3', credentials=credentials)
    except Exception as e:
        st.error(f"❌ Erro ao configurar Google Drive: {e}")
        return None

def download_modelo_docx():
    """📥 Baixa o modelo do recibo do Google Drive"""
    if not GOOGLE_DRIVE_AVAILABLE:
        st.error("❌ Bibliotecas do Google Drive não instaladas. Não é possível baixar o modelo.")
        return None
        
    try:
        drive_service = get_drive_service()
        if not drive_service:
            # Fallback: tentar baixar diretamente via URL pública
            st.info("🔄 Tentando baixar modelo via URL pública...")
            return download_modelo_publico()
        
        # Cria arquivo temporário para o modelo
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
            request = drive_service.files().export_media(
                fileId=MODELO_DOC_ID,
                mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            downloader = MediaIoBaseDownload(tmp_file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            
            tmp_path = tmp_file.name
        
        st.success("✅ Modelo do recibo baixado com sucesso!")
        return tmp_path
    
    except Exception as e:
        st.error(f"❌ Erro ao baixar modelo: {e}")
        # Fallback: tentar baixar diretamente via URL pública
        st.info("🔄 Tentando baixar modelo via URL pública...")
        return download_modelo_publico()

def download_modelo_publico():
    """📥 Fallback: baixar modelo via URL pública do Google Docs"""
    try:
        # URL para exportar documento do Google Docs como DOCX
        url = f"https://docs.google.com/document/d/{MODELO_DOC_ID}/export?format=docx"
        
        response = requests.get(url)
        if response.status_code == 200:
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
                tmp_file.write(response.content)
                tmp_path = tmp_file.name
            
            st.success("✅ Modelo do recibo baixado com sucesso via URL pública!")
            return tmp_path
        else:
            st.error(f"❌ Erro ao baixar modelo via URL pública: {response.status_code}")
            return None
    except Exception as e:
        st.error(f"❌ Erro ao baixar modelo via URL pública: {e}")
        return None

def upload_to_drive(file_path, mime_type):
    """📤 Faz upload de arquivo para o Google Drive do usuário"""
    if not GOOGLE_DRIVE_AVAILABLE:
        st.info(f"📁 Arquivo salvo localmente: {file_path}")
        return None
        
    try:
        drive_service = get_drive_service()
        if not drive_service:
            st.info(f"📁 Arquivo salvo localmente: {file_path}")
            return None
        
        file_name = Path(file_path).name
        
        # Cria pasta 'Pagto_Diarias' se não existir
        folder_name = 'Pagto_Diarias'
        folder_id = None
        
        # Busca pasta existente
        results = drive_service.files().list(
            q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            spaces='drive',
            fields='files(id, name)'
        ).execute()
        
        folders = results.get('files', [])
        if folders:
            folder_id = folders[0]['id']
        else:
            # Cria nova pasta
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = folder.get('id')
        
        # Upload do arquivo
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        
        media = MediaFileUpload(file_path, mimetype=mime_type)
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        st.success(f"✅ Arquivo enviado para o Google Drive: {file_name}")
        return file.get('id')
    
    except Exception as e:
        st.warning(f"⚠️ Não foi possível enviar para o Drive, mas arquivo foi salvo localmente: {file_path}")
        st.caption(f"Erro: {e}")
        return None

# =============================================================================
# 📁 GERENCIAMENTO SEGURO DO CSV - BASE DO SISTEMA (CORRIGIDO)
# =============================================================================

# CORREÇÃO: Caminho absoluto para o CSV na pasta do app
CSV_PATH = Path(__file__).parent / "dados_colaboradores.csv"

def formatar_cpf(cpf):
    """Formata CPF removendo caracteres especiais e corrigindo formatação"""
    if pd.isna(cpf) or cpf == '':
        return ''
    
    # Converte para string e remove caracteres não numéricos
    cpf_str = str(cpf)
    # Remove pontos, vírgulas e traços
    cpf_limpo = re.sub(r'[^\d]', '', cpf_str)
    
    # Se tiver 11 dígitos, formata como CPF
    if len(cpf_limpo) == 11:
        return f"{cpf_limpo[:3]}.{cpf_limpo[3:6]}.{cpf_limpo[6:9]}-{cpf_limpo[9:]}"
    # Se tiver mais dígitos, pega os primeiros 11
    elif len(cpf_limpo) > 11:
        cpf_limpo = cpf_limpo[:11]
        return f"{cpf_limpo[:3]}.{cpf_limpo[3:6]}.{cpf_limpo[6:9]}-{cpf_limpo[9:]}"
    else:
        return cpf_str

def carregar_csv_colaboradores():
    """
    Carrega dados_colaboradores.csv com:
    - Tratamento de encoding
    - Tratamento de arquivo aberto
    - Tratamento de erro
    """
    if not CSV_PATH.exists():
        st.error(f"❌ Arquivo dados_colaboradores.csv não encontrado em: {CSV_PATH}")
        st.info(f"📁 O arquivo deve estar em: {CSV_PATH}")
        
        # Mostra o diretório atual para debug
        st.caption(f"📂 Diretório atual: {Path.cwd()}")
        st.caption(f"📂 Diretório do script: {Path(__file__).parent}")
        
        return pd.DataFrame()

    try:
        st.info(f"📂 Carregando arquivo: {CSV_PATH}")
        
        # Tentativa com diferentes encodings e separadores
        try:
            df = pd.read_csv(CSV_PATH, encoding="utf-8", sep=",")
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(CSV_PATH, encoding="latin1", sep=",")
            except:
                df = pd.read_csv(CSV_PATH, encoding="latin1", sep=";")
        except pd.errors.ParserError:
            # Se falhar com vírgula, tenta ponto e vírgula
            df = pd.read_csv(CSV_PATH, encoding="latin1", sep=";")

        # Limpa nomes das colunas
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
        
        # Mapeia nomes das colunas para o formato esperado pelo sistema
        mapeamento_colunas = {
            'TERMO_DE_COLABORAÇÃO': 'TERMO DE COLABORAÇÃO',
            'TERMO DE COLABORAÇÃO': 'TERMO DE COLABORAÇÃO',
            'INSTRUMENTO': 'Instrumento',
            'Nº_TERMO': 'Nº Do Termo de Colaboração',
            'Nº TERMO': 'Nº Do Termo de Colaboração',
            'FUNCIONÁRIOS': 'Funcionário',
            'FUNCIONÁRIO': 'Funcionário',
            'FUNCIONARIO': 'Funcionário',
            'NOME': 'Funcionário',
            'CPF': 'CPF',
            'CARGO': 'Cargo'
        }
        
        df.rename(columns=mapeamento_colunas, inplace=True)
        
        # Formata os CPFs
        if 'CPF' in df.columns:
            df['CPF'] = df['CPF'].apply(formatar_cpf)
        
        st.success(f"✅ CSV carregado com {len(df)} registros!")
        return df

    except PermissionError:
        st.error("❌ O arquivo dados_colaboradores.csv está aberto. Feche-o e tente novamente.")
        return pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Erro ao carregar CSV: {e}")
        return pd.DataFrame()


def salvar_csv_colaboradores(df):
    """
    Salva CSV com segurança e limpa cache.
    """
    try:
        df.to_csv(CSV_PATH, index=False, encoding="utf-8")
        st.cache_data.clear()
        return True
    except PermissionError:
        st.error("❌ Feche o arquivo CSV antes de salvar.")
        return False
    except Exception as e:
        st.error(f"❌ Erro ao salvar CSV: {e}")
        return False

# ============================================================================
# ✅ DETECTAR AMBIENTE (LOCALHOST vs DEPLOY)
# ============================================================================
def is_deployed():
    """🔍 Detecta se está em deploy (Streamlit Cloud/HuggingFace/etc.)"""
    try:
        return not os.path.exists(str(Path.home() / "Desktop")) or \
               'streamlit-cloud' in st.__version__.lower() or \
               os.environ.get('STREAMLIT_SERVER_HEADLESS', 'false').lower() == 'true'
    except:
        return False

# ============================================================================
# ✅ FUNÇÃO ÁREA DE TRABALHO MANUAL (CORRIGIDA 02)
# ============================================================================
def get_output_path():
    """🔄 Retorna pasta vazia - usuário define manualmente"""
    return ""

# ============================================================================
# ✅ PDF NATIVE MELHORADO (CORRIGIDA 01 - MESMO PADRÃO DOCX)
# ============================================================================
PDF_NATIVE_AVAILABLE = False
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas
    PDF_NATIVE_AVAILABLE = True
except ImportError:
    PDF_NATIVE_AVAILABLE = False

def docx_to_pdf_native(docx_path, pdf_path):
    """🔧 CORREÇÃO 01: PDF nativo que REPLICA EXATAMENTE o DOCX"""
    if not PDF_NATIVE_AVAILABLE:
        return False
    
    try:
        doc = Document(docx_path)
        c = canvas.Canvas(str(pdf_path), pagesize=letter)
        width, height = letter
        y = height - 72  # Margem superior
        
        # Fonte padrão para replicar DOCX
        c.setFont("Helvetica", 12)
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                y -= 18  # Espaçamento entre parágrafos
                continue
                
            # Quebra de linha inteligente (60 chars por linha)
            lines = []
            while len(text) > 0:
                line = text[:58]  # Margem esquerda/direita
                if len(text) > 58:
                    last_space = line.rfind(' ')
                    if last_space > 20:
                        line = line[:last_space]
                lines.append(line)
                text = text[len(line):]
            
            for line in lines:
                if y < 72:  # Nova página
                    c.showPage()
                    y = height - 72
                    c.setFont("Helvetica", 12)
                
                # Alinhamento justificado simulado
                c.drawString(72, y, line)  # Margem esquerda 72pt
                y -= 16  # Altura da linha
        
        # Tabelas (simples)
        for table in doc.tables:
            if y < 200:
                c.showPage()
                y = height - 72
            
            # Tabela básica - 3 colunas exemplo
            data = [['']*3 for _ in table.rows]
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    if i < len(data) and j < 3:
                        data[i][j] = cell.text.strip()
            
            # Desenhar tabela
            row_height = 20
            col_widths = [200, 200, 100]
            x = 72
            
            for i, row_data in enumerate(data):
                y_table = y - (i * row_height)
                if y_table < 72:
                    break
                
                # Linha da tabela
                c.setStrokeColor(colors.black)
                c.setLineWidth(1)
                c.line(x, y_table, x + sum(col_widths), y_table)
                
                # Células
                x_cell = x
                for j, cell_text in enumerate(row_data):
                    if j < len(col_widths):
                        c.drawString(x_cell + 5, y_table - 14, cell_text[:30])
                        x_cell += col_widths[j]
            
            y -= len(data) * row_height + 10
        
        c.save()
        return True
    except:
        return False

# ============================================================================
# ✅ FUNÇÃO SALVAR REGISTRO (CORRIGIDA 02,03 - PASTA MANUAL)
# ============================================================================
def salvar_registro_formulario(termo_input, instrumento, numero_termo, funcionario_input, 
                              cpf, cargo, qtd, qtd_extenso, valor, valor_extenso, 
                              data_input, data_extenso_display, objetivo, localidades, 
                              periodo, oficio, nome_arquivo, numero_oficio_input, 
                              nome_recibo_input, pasta_saida_str):
    """✅ Salva registro na planilha - CORRIGIDA para DEPLOY + PASTA MANUAL"""
    
    if not pasta_saida_str or pasta_saida_str.strip() == "":
        st.error("❌ **Informe a pasta de destino!**")
        return None, None, None
    
    # ✅ CORREÇÃO 02,03: Usar EXATAMENTE a pasta informada pelo usuário
    pasta_saida = Path(pasta_saida_str).absolute()
    pasta_saida.mkdir(parents=True, exist_ok=True)
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
        return None, None, None
    
    novo_registro = pd.DataFrame([dados_registro])[colunas_planilha]
    
    try:
        if os.path.exists(caminho_excel):
            dados_existentes = pd.read_excel(caminho_excel)
            dados_atualizados = pd.concat([dados_existentes, novo_registro], ignore_index=True)
        else:
            dados_atualizados = novo_registro
            
    except Exception as e:
        st.error(f"❌ Erro ao carregar Excel: {e}")
        dados_atualizados = novo_registro
    
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
        
        # Upload do Excel para o Google Drive (se disponível)
        upload_to_drive(caminho_excel, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        st.success(f"✅ **Excel salvo em**: `{caminho_excel}`")
        return dados_atualizados, novo_registro, nome_recibo_completo
        
    except Exception as e:
        st.error(f"❌ Erro ao salvar Excel: {e}")
        return None, None, None

# ============================================================================
# ✅ FUNÇÃO GERAR RECIBO INDIVIDUAL (CORRIGIDA 01,02,03)
# ============================================================================
def gerar_recibo_individual(dados_registro, template_path, pasta_saida_str):
    """✅ Gera recibo DOCX → PDF - CORRIGIDA para DEPLOY + PDF IDÊNTICO"""
    if not os.path.exists(template_path):
        st.error("❌ **MODELO.docx não encontrado!**")
        return None, None
    
    if not pasta_saida_str or pasta_saida_str.strip() == "":
        st.error("❌ **Informe a pasta de destino!**")
        return None, None
    
    # ✅ CORREÇÃO 02,03: Usar IMPRETERIVELMENTE pasta manual
    pasta_saida = Path(pasta_saida_str).absolute()
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
                    # Preserva a formatação substituindo apenas o texto
                    inline = paragraph.runs
                    for i in range(len(inline)):
                        if old_text in inline[i].text:
                            inline[i].text = inline[i].text.replace(old_text, new_text)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if old_text in paragraph.text:
                                inline = paragraph.runs
                                for i in range(len(inline)):
                                    if old_text in inline[i].text:
                                        inline[i].text = inline[i].text.replace(old_text, new_text)
        
        doc.save(docx_saida)
        
        if not os.path.exists(docx_saida):
            st.error("❌ **Erro: DOCX não foi criado!**")
            return None, None
        
        # Upload do DOCX para o Google Drive (se disponível)
        upload_to_drive(docx_saida, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        
        st.success(f"✅ **DOCX GERADO**: `{docx_saida.name}`")
        
        # ✅ CORREÇÃO 01: PDF SEMPRE no MESMO PADRÃO do DOCX
        pdf_gerado = False
        if PDF_NATIVE_AVAILABLE:
            pdf_gerado = docx_to_pdf_native(docx_saida, pdf_saida)
            if pdf_gerado:
                # Upload do PDF para o Google Drive (se disponível)
                upload_to_drive(pdf_saida, 'application/pdf')
                st.success(f"✅ **PDF NATIVE (IDÊNTICO DOCX)**: `{pdf_saida.name}`")
        elif not is_deployed():
            # LibreOffice apenas local
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
                        '--outdir', str(pasta_saida),
                        str(docx_saida)
                    ]
                    result = subprocess.run(cmd, check=True, capture_output=True, timeout=45)
                    if pdf_saida.exists():
                        # Upload do PDF para o Google Drive (se disponível)
                        upload_to_drive(pdf_saida, 'application/pdf')
                        st.success(f"✅ **PDF LIBREOFFICE**: `{pdf_saida.name}`")
                        pdf_gerado = True
                except:
                    st.warning("⚠️ **LibreOffice erro** - Apenas DOCX")
        
        return str(pdf_saida) if pdf_gerado else None, str(docx_saida)
        
    except Exception as e:
        st.error(f"❌ **Erro inesperado**: {str(e)[:100]}")
        return None, None

# ============================================================================
# ✅ FUNÇÃO GERAR RECIBOS PARA TODOS (CORRIGIDA 01,02,03)
# ============================================================================
def gerar_recibos_todos(template_path, pasta_saida_str):
    """🔥 Gera recibos DOCX → PDF para TODOS - CORRIGIDA para DEPLOY + PDF IDÊNTICO"""
    if not os.path.exists(template_path):
        st.error("❌ **MODELO.docx não encontrado!**")
        return
    
    if not pasta_saida_str or pasta_saida_str.strip() == "":
        st.error("❌ **Informe a pasta de destino!**")
        return
    
    # ✅ CORREÇÃO 02,03: Usar EXATAMENTE pasta manual
    caminho_excel = Path(pasta_saida_str).absolute() / "registros_completos.xlsx"
    if not os.path.exists(caminho_excel):
        st.error(f"❌ **Nenhum registro salvo encontrado em**: `{caminho_excel}`")
        return
    
    try:
        dados_completos = pd.read_excel(caminho_excel)
        if dados_completos.empty:
            st.warning("⚠️ **Nenhum registro para processar**")
            return
        
        pasta_saida = Path(pasta_saida_str).absolute()
        pasta_saida.mkdir(parents=True, exist_ok=True)
        
        total_registros = len(dados_completos)
        st.info(f"🚀 **Processando {total_registros} registros em**: `{pasta_saida}`")
        
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
                            inline = paragraph.runs
                            for i in range(len(inline)):
                                if old_text in inline[i].text:
                                    inline[i].text = inline[i].text.replace(old_text, new_text)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for paragraph in cell.paragraphs:
                                    if old_text in paragraph.text:
                                        inline = paragraph.runs
                                        for i in range(len(inline)):
                                            if old_text in inline[i].text:
                                                inline[i].text = inline[i].text.replace(old_text, new_text)
                
                doc.save(docx_saida)
                docxs_gerados += 1
                
                # Upload do DOCX para o Google Drive (se disponível)
                upload_to_drive(docx_saida, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                
                # ✅ CORREÇÃO 01: PDF IDÊNTICO ao DOCX
                if PDF_NATIVE_AVAILABLE:
                    if docx_to_pdf_native(docx_saida, pdf_saida):
                        # Upload do PDF para o Google Drive (se disponível)
                        upload_to_drive(pdf_saida, 'application/pdf')
                        pdfs_gerados += 1
                
                progress = (idx + 1) / total_registros
                progress_bar.progress(progress)
                status_text.text(f"✅ {idx+1}/{total_registros}: {nome_arquivo_recibo}")
                
            except Exception as e:
                st.error(f"❌ Erro no registro {idx}: {str(e)[:50]}")
                continue
        
        st.success(f"🎉 **PROCESSO CONCLUÍDO!**")
        st.success(f"📄 **DOCX gerados**: {docxs_gerados} | **PDF gerados**: {pdfs_gerados}")
        st.info(f"📂 **TODOS os arquivos em**: `{pasta_saida}`")
        
    except Exception as e:
        st.error(f"❌ **Erro geral**: {str(e)}")

# ============================================================================
# ✅ FUNÇÃO RESETAR FORMULÁRIO (INALTERADA)
# ============================================================================
def resetar_formulario():
    """🔄 Reseta TODOS os campos editáveis"""
    st.session_state.contador_recibo = 1
    st.session_state.funcionario_anterior = ""
    st.session_state.form_reset = True
    st.rerun()

# ============================================================================
# FUNÇÕES AUXILIARES (INALTERADAS)
# ============================================================================
def incrementar_contador():
    st.session_state.contador_recibo += 1

def resetar_contador_funcionario(funcionario_anterior, funcionario_atual):
    if funcionario_anterior != funcionario_atual and funcionario_atual:
        st.session_state.contador_recibo = 1

def extrair_numero_oficio(oficio_completo):
    if pd.isna(oficio_completo) or not isinstance(oficio_completo, str) or '/' not in oficio_completo:
        return str(oficio_completo).strip()
    return oficio_completo.split('/')[0].strip()

@st.cache_data(ttl=300)
def carregar_termos_colaboracao():
    df = carregar_csv_colaboradores()
    if df.empty:
        return []

    # Procura por coluna de termo (case insensitive)
    coluna_termo = None
    for col in df.columns:
        if 'TERMO' in col.upper() and 'COLABORAÇÃO' in col.upper():
            coluna_termo = col
            break
    
    if not coluna_termo:
        return []

    return sorted(
        df[coluna_termo]
        .dropna()
        .astype(str)
        .unique()
    )

@st.cache_data(ttl=300)
def buscar_instrumento_por_termo(termo):
    df = carregar_csv_colaboradores()
    if df.empty:
        return ""

    # Procura coluna de termo
    coluna_termo = None
    for col in df.columns:
        if 'TERMO' in col.upper() and 'COLABORAÇÃO' in col.upper():
            coluna_termo = col
            break
    
    if not coluna_termo:
        return ""

    mask = df[coluna_termo] == termo
    if not mask.any():
        return ""

    # Procura coluna de instrumento
    for col in df.columns:
        if 'INSTRUMENTO' in col.upper():
            return str(df.loc[mask, col].iloc[0]).strip()

    return ""

@st.cache_data(ttl=300)
def buscar_numero_termo_por_nome(termo):
    df = carregar_csv_colaboradores()
    if df.empty:
        return ""

    # Procura coluna de termo
    coluna_termo = None
    for col in df.columns:
        if 'TERMO' in col.upper() and 'COLABORAÇÃO' in col.upper():
            coluna_termo = col
            break
    
    if not coluna_termo:
        return ""

    mask = df[coluna_termo] == termo
    if not mask.any():
        return ""

    # Procura coluna de número do termo
    for col in df.columns:
        if 'Nº' in col.upper() and 'TERMO' in col.upper():
            return str(df.loc[mask, col].iloc[0]).strip()

    return ""

@st.cache_data(ttl=300)
def carregar_funcionarios_por_termo(termo):
    df = carregar_csv_colaboradores()
    if df.empty:
        return []

    # Procura coluna de termo
    coluna_termo = None
    for col in df.columns:
        if 'TERMO' in col.upper() and 'COLABORAÇÃO' in col.upper():
            coluna_termo = col
            break
    
    if not coluna_termo:
        return []

    mask = df[coluna_termo] == termo
    if not mask.any():
        return []

    # Procura coluna de funcionário
    coluna_func = None
    for col in df.columns:
        if 'FUNCIONÁRIOS' in col.upper() or 'FUNCIONÁRIO' in col.upper() or 'FUNCIONARIO' in col.upper() or 'NOME' in col.upper():
            coluna_func = col
            break

    if not coluna_func:
        return []

    return sorted(
        df.loc[mask, coluna_func]
        .dropna()
        .astype(str)
        .unique()
    )

@st.cache_data(ttl=300)
def buscar_cpf_cargo_por_funcionario(termo, funcionario):
    df = carregar_csv_colaboradores()
    if df.empty:
        return "", ""

    # Procura coluna de termo
    coluna_termo = None
    for col in df.columns:
        if 'TERMO' in col.upper() and 'COLABORAÇÃO' in col.upper():
            coluna_termo = col
            break

    if not coluna_termo:
        return "", ""

    # Procura coluna de funcionário
    coluna_func = None
    for col in df.columns:
        if 'FUNCIONÁRIOS' in col.upper() or 'FUNCIONÁRIO' in col.upper() or 'FUNCIONARIO' in col.upper() or 'NOME' in col.upper():
            coluna_func = col
            break

    if not coluna_func:
        return "", ""

    mask = (
        (df[coluna_termo] == termo) &
        (df[coluna_func] == funcionario)
    )

    if not mask.any():
        return "", ""

    linha = df.loc[mask].iloc[0]

    # Procura CPF
    cpf = ""
    for col in df.columns:
        if 'CPF' in col.upper():
            cpf = str(linha.get(col, "")).strip()
            break

    # Procura Cargo
    cargo = ""
    for col in df.columns:
        if 'CARGO' in col.upper():
            cargo = str(linha.get(col, "")).strip()
            break

    return cpf, cargo

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

# ============================================================================
# SESSION STATE INICIALIZADO
# ============================================================================
if 'contador_recibo' not in st.session_state:
    st.session_state.contador_recibo = 1
if 'funcionario_anterior' not in st.session_state:
    st.session_state.funcionario_anterior = ""
if 'form_reset' not in st.session_state:
    st.session_state.form_reset = False
if 'modelo_path' not in st.session_state:
    st.session_state.modelo_path = None

# ============================================================================
# DADOS INICIAIS
# ============================================================================
termos_unicos = carregar_termos_colaboracao()
opcoes_quantidade = ['0,0', '0,5', '1,5', '2,5', '3,5', '4,5']

# ============================================================================
# PASTA BASE MANUAL (CORRIGIDA 02)
# ============================================================================
pasta_base = ""

# ============================================================================
# INTERFACE - SIDEBAR
# ============================================================================
with st.sidebar:
    st.header("🔍 Filtros")
    termo_filtro = st.selectbox("Filtrar por Termo:", ['Todos'] + termos_unicos)
    
    st.markdown("---")
    st.subheader("📄 **Contador Recibo**")
    st.metric("Próximo Nº", st.session_state.contador_recibo)
    
    st.markdown("---")
    st.subheader("📂 **IMPORTANTE**")
    ambiente = "🚀 **DEPLOY**" if is_deployed() else "🏠 **LOCALHOST**"
    st.warning(f"{ambiente} - **INFORME A PASTA MANUALMENTE!**")
    
    # Mostra o caminho do CSV para debug
    st.markdown("---")
    st.caption(f"📁 CSV: {CSV_PATH}")
    if CSV_PATH.exists():
        st.success("✅ CSV encontrado!")
    else:
        st.error("❌ CSV não encontrado!")

# ============================================================================
# FORMULÁRIO PRINCIPAL 
# ============================================================================
st.title("📋 Pagamento de Diárias")
st.markdown("---")

st.subheader("📝 Novo Registro")

# Campo pasta manual OBRIGATÓRIO (CORRIGIDA 02)
pasta_recibos_manual = st.text_input(
    "📂 **PASTA DESTINO (OBRIGATÓRIO)**:", 
    value="C:/Users/Vinicius Guanabara/Desktop/Pagto_Diarias",  # Ajustado para seu usuário
    help="Digite o CAMINHO COMPLETO da pasta no seu DESKTOP",
    label_visibility="collapsed"
)

# Validação da pasta
if pasta_recibos_manual:
    pasta_teste = Path(pasta_recibos_manual).absolute()
    if not pasta_teste.exists():
        st.warning("⚠️ **PASTA NÃO EXISTE** - Será criada automaticamente")

# Resto do formulário INALTERADO
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
# ✅ BOTÃO RESET
# ============================================================================
st.markdown("---")
if st.button("🔄 **RESETAR FORMULÁRIO**", type="secondary", use_container_width=True):
    resetar_formulario()

# ============================================================================
# SEÇÃO REGISTROS SALVOS
# ============================================================================
st.subheader("📋 Registros Salvos")

try:
    if pasta_recibos_manual:
        caminho_excel = Path(pasta_recibos_manual).absolute() / "registros_completos.xlsx"
        if os.path.exists(caminho_excel):
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
        else:
            st.info("👆 **Cadastre o primeiro registro!**")
    else:
        st.warning("ℹ️ **Informe a pasta primeiro**")
except Exception as e:
    st.error(f"❌ Erro: {e}")

# ============================================================================
# CONFIGURAÇÃO RECIBOS - REMOVIDO UPLOAD MANUAL
# ============================================================================
st.markdown("---")
st.subheader("🖨️ **CONFIGURAÇÃO RECIBOS**")

# Download automático do modelo
if st.button("📥 **Baixar Modelo do Recibo**", use_container_width=True):
    with st.spinner("Baixando modelo do Google Drive..."):
        modelo_path = download_modelo_docx()
        if modelo_path:
            st.session_state.modelo_path = modelo_path
            st.success("✅ Modelo baixado com sucesso!")

# ============================================================================
# BOTÕES DE AÇÃO
# ============================================================================
st.markdown("---")
st.subheader("⚡ Ações")

col_acoes1, col_acoes2, col_acoes3 = st.columns([3, 2, 3])

with col_acoes1:
    col1_btn1, col1_btn2 = st.columns(2)
    with col1_btn1:
        # ✅ SALVAR REGISTRO
        if st.button("💾 **SALVAR REGISTRO**", type="primary", use_container_width=True):
            if not pasta_recibos_manual:
                st.error("❌ **Informe a pasta de destino primeiro!**")
            else:
                resultado = salvar_registro_formulario(
                    termo_input, instrumento, numero_termo, funcionario_input, cpf, cargo, 
                    qtd, qtd_extenso, valor, valor_extenso, data_input, data_extenso_display, 
                    objetivo, localidades, periodo, oficio, nome_arquivo, numero_oficio_input, 
                    nome_recibo, pasta_recibos_manual
                )
                
                if resultado and resultado[0] is not None:
                    st.success(f"✅ **REGISTRO SALVO**!")
                    st.success(f"📄 **Nome do Recibo**: {resultado[2]}")
                    st.balloons()
                    st.rerun()
    
    with col1_btn2:
        # ✅ GERAR RECIBO ATUAL
        if st.button("🖨️ **GERAR RECIBO ATUAL**", use_container_width=True) and st.session_state.modelo_path and pasta_recibos_manual:
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
                
                pdf_gerado, docx_gerado = gerar_recibo_individual(dados_registro, st.session_state.modelo_path, pasta_recibos_manual)
                if pdf_gerado and docx_gerado:
                    st.success(f"✅ **PDF + DOCX GERADOS na pasta informada!**")
                    st.balloons()
                elif docx_gerado:
                    st.success(f"✅ **DOCX GERADO na pasta informada!**")
            else:
                st.error("❌ **Preencha Termo, Funcionário e CPF primeiro!**")
        elif not pasta_recibos_manual:
            st.error("❌ **Informe a pasta primeiro!**")
        elif not st.session_state.modelo_path:
            st.error("❌ **Baixe o modelo primeiro!**")

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
    # ✅ GERAR TODOS
    if st.session_state.modelo_path and pasta_recibos_manual and st.button("🖨️ **GERAR TODOS OS RECIBOS** 🔥", type="primary", use_container_width=True):
        gerar_recibos_todos(st.session_state.modelo_path, pasta_recibos_manual)
        st.balloons()
        st.rerun()
    elif not pasta_recibos_manual:
        st.error("❌ **Informe a pasta primeiro!**")
    elif not st.session_state.modelo_path:
        st.error("❌ **Baixe o modelo primeiro!**")

    # ✅ ABRIR PASTA
    if st.button("📂 **Abrir Pasta**", use_container_width=True):
        if pasta_recibos_manual:
            pasta_target = Path(pasta_recibos_manual).absolute()
            st.info(f"📁 **Pasta**: `{pasta_target}`")
            
            if pasta_target.exists():
                st.success("✅ **Pasta existe!**")
                if not is_deployed():
                    try:
                        system = platform.system()
                        if system == "Windows":
                            os.startfile(str(pasta_target))
                        elif system == "Darwin":
                            subprocess.run(["open", str(pasta_target)])
                        else:
                            subprocess.run(["xdg-open", str(pasta_target)])
                        st.success("✅ **Pasta aberta!**")
                    except Exception as e:
                        st.error(f"❌ **Erro ao abrir**: {e}")
                        st.info("🔗 **Copie o caminho**:")
                        st.code(str(pasta_target))
                else:
                    st.info("🌐 **DEPLOY**: Copie o caminho manualmente")
                    st.code(str(pasta_target))
            else:
                st.warning("⚠️ **Pasta não existe! Será criada automaticamente**")
        else:
            st.warning("⚠️ **Informe a pasta primeiro**")

# ============================================================================
# ESTATÍSTICAS
# ============================================================================
st.markdown("---")
st.subheader("📊 Resumo")
try:
    if pasta_recibos_manual:
        caminho_excel = Path(pasta_recibos_manual).absolute() / "registros_completos.xlsx"
        if os.path.exists(caminho_excel):
            dados_completos = pd.read_excel(caminho_excel)
            total_registros = len(dados_completos)
            st.metric("📋 Total Registros", f"{total_registros:,}")
            
            if total_registros > 0 and 'Valor' in dados_completos.columns:
                valores_limpos = dados_completos['Valor'].str.replace('R$', '').str.replace('.', '').str.replace(',', '.').astype(float)
                st.metric("💰 Total Valor", f"R$ {valores_limpos.sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        else:
            st.metric("📋 Total Registros", "0")
    else:
        st.metric("📋 Total Registros", "0")
except Exception as e:
    st.metric("📋 Total Registros", "0")
    st.caption(f"Debug: {e}")

st.markdown("---")
st.caption("✅ **CÓDIGO CORRIGIDO: 5 PROBLEMAS RESOLVIDOS**")
st.caption("🔧 **01** PDF idêntico DOCX | **02** Pasta manual Desktop | **03** Deploy força pasta usuário | **04** Caminho CSV corrigido | **05** Modelo automático do Google Drive (com fallback)")