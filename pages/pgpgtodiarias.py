import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO, StringIO
import openpyxl
from datetime import datetime
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os

# Configuração da página
st.set_page_config(
    page_title="Importador de Diárias",
    page_icon="💰",
    layout="wide"
)

class ProcessadorDiarias:
    def __init__(self):
        self.df_importado = None
        self.aba_importada_nome = None
        self.mapeamento_colunas = {
            'CPF': None,
            'Funcionário': None,
            'Quantidade': None,
            'Período': None,
            'Ofício': None,
            'Nº do Ofício': None,
            'Data Recibo': None,
            'Nome do Recibo': None,
            'Termo de Colaboração': None,
            'Instrumento': None
        }
        
    def identificar_colunas(self, df):
        """Identifica automaticamente as colunas necessárias no DataFrame"""
        colunas_encontradas = {}
        
        # Mapeamento de possíveis nomes de colunas
        mapeamento_possivel = {
            'CPF': ['CPF', 'Cpf', 'cpf', 'DOC', 'Doc', 'documento', 'Documento', 'CPF/CGC', 'CPF/CNPJ'],
            'Funcionário': ['Funcionário', 'Funcionario', 'funcionario', 'NOME', 'Nome', 'nome', 'SERVIDOR', 'Servidor', 'EMPREGADO', 'empregado'],
            'Quantidade': ['Quantidade', 'quantidade', 'QTD', 'Qtd', 'qtd', 'DIAS', 'Dias', 'dias', 'QTDE', 'qtde'],
            'Período': ['Período', 'Periodo', 'periodo', 'MÊS', 'Mês', 'mes', 'COMPETÊNCIA', 'Competencia', 'REFERÊNCIA', 'Referencia'],
            'Ofício': ['Ofício', 'Oficio', 'oficio', 'OFÍCIO', 'NUMERO', 'Número', 'Nº', 'nº', 'NUM'],
            'Nº do Ofício': ['Nº do Ofício', 'No do Oficio', 'NUMERO OFICIO', 'Número Ofício', 'Nº OFICIO', 'NUM OFICIO'],
            'Data Recibo': ['Data Recibo', 'Data do Recibo', 'DATA', 'Data', 'DT', 'DT RECIBO', 'DATA RECIBO'],
            'Nome do Recibo': ['Nome do Recibo', 'NOME RECIBO', 'Nome Recibo', 'RECIBO', 'Recibo', 'NOME DO RECIBO'],
            'Termo de Colaboração': ['Termo de Colaboração', 'Termo Colaboracao', 'TERMO', 'Termo', 'Colaboração', 'Colaboracao'],
            'Instrumento': ['Instrumento', 'INSTRUMENTO', 'instrumento', 'Tipo Instrumento', 'TIPO INSTRUMENTO']
        }
        
        # Primeiro, verifica se o DataFrame tem cabeçalho (colunas com nomes)
        if len(df.columns) > 0:
            # Se as colunas são strings, assume que é cabeçalho
            if all(isinstance(col, str) for col in df.columns):
                # Procura cada coluna necessária
                for coluna_alvo, possibilidades in mapeamento_possivel.items():
                    for col in df.columns:
                        col_lower = str(col).lower().strip()
                        for possibilidade in possibilidades:
                            if possibilidade.lower() in col_lower:
                                colunas_encontradas[coluna_alvo] = col
                                break
        
        return colunas_encontradas
    
    def extrair_competencia_mmaa(self, data):
        """Extrai competência no formato MM/AA de uma data"""
        try:
            if pd.isna(data):
                return None
            
            if isinstance(data, (pd.Timestamp, datetime)):
                return data.strftime('%m/%y')
            
            elif isinstance(data, str):
                data = data.strip()
                
                for fmt in ['%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y', '%d%m%Y']:
                    try:
                        data_dt = datetime.strptime(data, fmt)
                        return data_dt.strftime('%m/%y')
                    except:
                        continue
                
                match = re.search(r'(\d{1,2})[/\-.](\d{4})', data)
                if match:
                    mes, ano = match.groups()
                    ano_curto = ano[2:]
                    return f"{int(mes):02d}/{ano_curto}"
                
                match = re.search(r'(\d{2})[/\-.]?(\d{4})', data)
                if match:
                    mes, ano = match.groups()
                    ano_curto = ano[2:]
                    return f"{mes}/{ano_curto}"
                
                match = re.search(r'(\d{2})[/\-.]?(\d{2})', data)
                if match:
                    mes, ano = match.groups()
                    return f"{mes}/{ano}"
            
            return None
        except:
            return None
    
    def extrair_competencia_numeros(self, data):
        """Extrai competência no formato MMAA (apenas números, sem barra)"""
        try:
            if pd.isna(data):
                return None
            
            if isinstance(data, (pd.Timestamp, datetime)):
                return data.strftime('%m%y')
            
            elif isinstance(data, str):
                data = data.strip()
                
                for fmt in ['%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y', '%d%m%Y']:
                    try:
                        data_dt = datetime.strptime(data, fmt)
                        return data_dt.strftime('%m%y')
                    except:
                        continue
                
                match = re.search(r'(\d{1,2})[/\-.](\d{4})', data)
                if match:
                    mes, ano = match.groups()
                    ano_curto = ano[2:]
                    return f"{int(mes):02d}{ano_curto}"
                
                match = re.search(r'(\d{2})[/\-.]?(\d{4})', data)
                if match:
                    mes, ano = match.groups()
                    ano_curto = ano[2:]
                    return f"{mes}{ano_curto}"
                
                match = re.search(r'(\d{2})[/\-.]?(\d{2})', data)
                if match:
                    mes, ano = match.groups()
                    return f"{mes}{ano}"
            
            return None
        except:
            return None
    
    def limpar_caracteres_especiais(self, texto):
        """Remove todos os caracteres especiais, mantendo apenas números"""
        if pd.isna(texto) or texto is None:
            return None
        try:
            texto_str = str(texto).strip()
            numeros = re.sub(r'[^0-9.,]', '', texto_str)  # Mantém vírgula e ponto
            # Substitui vírgula por ponto para conversão float
            numeros = numeros.replace(',', '.')
            return numeros if numeros else None
        except:
            return None
    
    def gerar_contrato(self, oficio, data):
        """Gera o número do contrato: ofício (só números) + competência (só números)"""
        try:
            if pd.isna(oficio) or pd.isna(data):
                return None
            
            oficio_limpo = self.limpar_caracteres_especiais(oficio)
            competencia = self.extrair_competencia_numeros(data)
            
            if oficio_limpo and competencia:
                return f"{oficio_limpo}{competencia}"
            elif oficio_limpo:
                return oficio_limpo
            else:
                return None
        except:
            return None
    
    def gerar_dl_unico(self, df, coluna_oficio):
        """Gera a coluna DL com numeração única para cada ofício (apenas números)"""
        try:
            dl_dict = {}
            dl_lista = []
            
            for idx, oficio in enumerate(df[coluna_oficio]):
                if pd.isna(oficio):
                    dl_lista.append(None)
                    continue
                
                oficio_limpo = self.limpar_caracteres_especiais(oficio)
                
                if not oficio_limpo:
                    dl_lista.append(None)
                    continue
                
                if oficio_limpo in dl_dict:
                    dl_dict[oficio_limpo] += 1
                    dl_lista.append(f"{oficio_limpo}{dl_dict[oficio_limpo]}")
                else:
                    dl_dict[oficio_limpo] = 1
                    dl_lista.append(f"{oficio_limpo}1")
            
            return dl_lista
        except Exception as e:
            st.error(f"Erro ao gerar DL: {str(e)}")
            return [None] * len(df)
    
    def carregar_arquivo(self, arquivo, aba_selecionada):
        """Carrega o arquivo Excel com a aba especificada"""
        try:
            self.df_importado = pd.read_excel(arquivo, sheet_name=aba_selecionada, header=0)
            
            if len(self.df_importado) > 0:
                primeira_linha = self.df_importado.iloc[0]
                valores_numericos = sum([isinstance(v, (int, float)) for v in primeira_linha])
                if valores_numericos > len(primeira_linha) / 2:
                    self.df_importado = pd.read_excel(arquivo, sheet_name=aba_selecionada, header=None)
            
            self.aba_importada_nome = aba_selecionada
            colunas_encontradas = self.identificar_colunas(self.df_importado)
            
            return True, colunas_encontradas
            
        except Exception as e:
            st.error(f"❌ Erro ao carregar arquivo: {str(e)}")
            return False, {}
    
    def processar_dados(self, mapeamento):
        """Processa os dados conforme o mapeamento especificado"""
        if self.df_importado is None:
            st.error("Arquivo não carregado!")
            return None
        
        try:
            # Cria DataFrame de resultado com as colunas desejadas
            colunas_resultado = [
                'CPF',
                'FUNCIONÁRIO',
                'QTD',
                'PERÍODO',
                'OFÍCIO',
                'CONTRATO',
                'DL',
                'NOME RECIBO',
                'TERMO_COLABORACAO',
                'INSTRUMENTO'
            ]
            
            # Listas para armazenar os dados
            dados = {col: [] for col in colunas_resultado}
            
            # Processa cada linha do DataFrame importado
            for idx in range(len(self.df_importado)):
                # Inicializa a linha com None
                for col in colunas_resultado:
                    dados[col].append(None)
                
                # Mapeia cada campo
                for coluna_destino, coluna_origem in mapeamento.items():
                    if coluna_origem and coluna_origem != "Selecione...":
                        try:
                            if coluna_origem in self.df_importado.columns:
                                valor = self.df_importado.iloc[idx][coluna_origem]
                                
                                if coluna_destino == 'CPF':
                                    dados['CPF'][idx] = self.limpar_caracteres_especiais(valor) if pd.notna(valor) else None
                                elif coluna_destino == 'Funcionário':
                                    dados['FUNCIONÁRIO'][idx] = valor
                                elif coluna_destino == 'Quantidade':
                                    # CORREÇÃO: Preserva as casas decimais
                                    if pd.notna(valor):
                                        try:
                                            # Se já for numérico, mantém como está
                                            if isinstance(valor, (int, float)):
                                                dados['QTD'][idx] = float(valor)
                                            else:
                                                # Tenta converter string mantendo decimais
                                                valor_str = str(valor).strip()
                                                # Substitui vírgula por ponto para conversão
                                                valor_str = valor_str.replace(',', '.')
                                                # Remove caracteres não numéricos exceto ponto
                                                valor_limpo = re.sub(r'[^0-9.]', '', valor_str)
                                                if valor_limpo:
                                                    dados['QTD'][idx] = float(valor_limpo)
                                                else:
                                                    dados['QTD'][idx] = 0.0
                                        except:
                                            dados['QTD'][idx] = 0.0
                                    else:
                                        dados['QTD'][idx] = 0.0
                                elif coluna_destino == 'Período':
                                    dados['PERÍODO'][idx] = valor
                                elif coluna_destino == 'Ofício':
                                    dados['OFÍCIO'][idx] = valor
                                elif coluna_destino == 'Nome do Recibo':
                                    dados['NOME RECIBO'][idx] = valor
                                elif coluna_destino == 'Termo de Colaboração':
                                    dados['TERMO_COLABORACAO'][idx] = valor
                                elif coluna_destino == 'Instrumento':
                                    dados['INSTRUMENTO'][idx] = valor
                        except Exception as e:
                            continue
            
            # Cria DataFrame a partir dos dicionários
            df_resultado = pd.DataFrame(dados)
            
            # Converte QTD para numérico mantendo as casas decimais
            df_resultado['QTD'] = pd.to_numeric(df_resultado['QTD'], errors='coerce').fillna(0)
            
            # Processa CONTRATO
            if (mapeamento.get('Nº do Ofício') and mapeamento.get('Nº do Ofício') != "Selecione..." and 
                mapeamento.get('Data Recibo') and mapeamento.get('Data Recibo') != "Selecione..."):
                
                col_oficio = mapeamento['Nº do Ofício']
                col_data = mapeamento['Data Recibo']
                
                if col_oficio in self.df_importado.columns and col_data in self.df_importado.columns:
                    contratos = []
                    for idx in range(len(self.df_importado)):
                        oficio = self.df_importado.iloc[idx][col_oficio]
                        data = self.df_importado.iloc[idx][col_data]
                        contrato = self.gerar_contrato(oficio, data)
                        contratos.append(contrato)
                    
                    df_resultado['CONTRATO'] = contratos
                    
                    if len(self.df_importado) > 0:
                        exemplo_data = self.df_importado.iloc[0][col_data] if col_data in self.df_importado.columns else None
                        if exemplo_data is not None:
                            competencia_exemplo = self.extrair_competencia_mmaa(exemplo_data)
                            st.info(f"📅 Exemplo de extração: Data '{exemplo_data}' → Competência '{competencia_exemplo}' → CONTRATO: {contratos[0] if contratos else 'N/A'}")
            
            # Processa DL
            if mapeamento.get('Nº do Ofício') and mapeamento.get('Nº do Ofício') != "Selecione...":
                col_oficio = mapeamento['Nº do Ofício']
                if col_oficio in self.df_importado.columns:
                    df_resultado['DL'] = self.gerar_dl_unico(self.df_importado, col_oficio)
            
            # Remove linhas completamente vazias
            df_resultado = df_resultado.dropna(how='all').reset_index(drop=True)
            
            return df_resultado
            
        except Exception as e:
            st.error(f"❌ Erro durante o processamento: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
            return None

def formatar_moeda_br(valor):
    """Formata um valor numérico para moeda brasileira (R$)"""
    if valor is None or pd.isna(valor) or valor == 0:
        return "R$ 0,00"
    try:
        # Formata com 2 casas decimais
        valor_formatado = f"R$ {float(valor):,.2f}"
        # Ajusta para padrão brasileiro
        return valor_formatado.replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

def inicializar_session_state():
    """Inicializa as variáveis de sessão"""
    if 'processamento_concluido' not in st.session_state:
        st.session_state.processamento_concluido = False
    if 'df_resultado' not in st.session_state:
        st.session_state.df_resultado = None
    if 'mapeamento_salvo' not in st.session_state:
        st.session_state.mapeamento_salvo = None
    if 'colunas_disponiveis' not in st.session_state:
        st.session_state.colunas_disponiveis = []
    if 'processador' not in st.session_state:
        st.session_state.processador = None
    if 'dados_pa' not in st.session_state:
        st.session_state.dados_pa = {
            'tipo_pa': 'COMPRA DIRETA',
            'termo_colaboracao': '',
            'instrumento': '',
            'numero_pa': '',
            'data_pa': datetime.now().date(),
            'valor_diaria': 0.0,
            'item_aquisicao': '',
            'valor_pa': 0.0
        }
    if 'dados_pa_salvos' not in st.session_state:
        st.session_state.dados_pa_salvos = False
    if 'des_item_pa' not in st.session_state:
        st.session_state.des_item_pa = "5"
    if 'ultimo_calculo_pa' not in st.session_state:
        st.session_state.ultimo_calculo_pa = 0.0

class GerarPAPessoal:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def iniciar_driver(self):
        """Inicializa o driver do Chrome"""
        chrome_options = Options()
        # chrome_options.add_argument("--headless")
        chrome_options.add_argument("--start-maximized")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 20)
        
    def fazer_login(self, usuario, senha):
        """Faz login no sistema e-parcerias"""
        try:
            self.driver.get("https://e-parcerias.cge.ce.gov.br/e-parcerias-web/padrao-web/paginas/seguranca/login.seam")
            time.sleep(5)
            
            login_field = self.wait.until(EC.element_to_be_clickable((By.ID, "loginForm:username")))
            login_field.click()
            login_field.send_keys(usuario)
            
            senha_field = self.driver.find_element(By.ID, "loginForm:password")
            senha_field.click()
            senha_field.send_keys(senha)
            time.sleep(2)
            
            self.driver.find_element(By.ID, "loginForm:login").click()
            time.sleep(2)
            
            self.driver.find_element(By.ID, "j_id237:listPessoaJuridicaDecorate:pagedDataTable:0:uncheckRadio").click()
            self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.TAB)
            time.sleep(2)
            
            self.driver.get("https://e-parcerias.cge.ce.gov.br/e-parcerias-web/paginas/home/home.seam?actionMethod=paginas%2Fhome%2Fhome.xhtml%3AhomeController.escolherParceiro%28%29")
            time.sleep(5)
            
            return True
        except Exception as e:
            st.error(f"Erro no login: {str(e)}")
            return False
    
    def navegar_para_processo_aquisicao(self):
        """Navega até a tela de processo de aquisição"""
        try:
            menu_exec = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu:form_menu_corporativo:j_id73']/div[1]/table/tbody/tr/td")))
            menu_exec.click()
            time.sleep(2)
            
            proc_aquisicao = self.driver.find_element(By.XPATH, "//*[@id='menu:form_menu_corporativo:subMenuProcessoAquisicao']")
            proc_aquisicao.click()
            time.sleep(2)
            
            incluir_proc = self.driver.find_element(By.XPATH, "//*[@id='tableInstrumento']/div[1]/div[2]/button")
            incluir_proc.click()
            time.sleep(2)
            
            compra_direta = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div/div/div/table/tbody/tr[2]/td[1]/input")
            compra_direta.click()
            time.sleep(2)
            
            confirmar = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[3]/a/button")
            confirmar.click()
            time.sleep(2)
            
            return True
        except Exception as e:
            st.error(f"Erro na navegação: {str(e)}")
            return False
    
    def selecionar_instrumento(self, num_instrumento):
        """Seleciona o número do instrumento"""
        try:
            select_instrumento = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[2]/div[1]/form/div[1]/div[1]/div/i")
            select_instrumento.click()
            time.sleep(2)
            
            body = self.driver.find_element(By.TAG_NAME, 'body')
            body.send_keys(Keys.TAB)
            body.send_keys(Keys.TAB)
            body.send_keys(num_instrumento)
            
            pesquisar = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div[1]/form/button")
            pesquisar.click()
            
            selecionar = self.wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/div[2]/div[2]/div/div/div[1]/table/tbody/tr/td[1]/input")))
            selecionar.click()
            
            confirmar = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[3]/button[1]")
            confirmar.click()
            
            return True
        except Exception as e:
            st.error(f"Erro ao selecionar instrumento: {str(e)}")
            return False
    
    def preencher_dados_pa(self, num_pa, data_pa, valor_pa):
        """Preenche os dados principais da PA"""
        try:
            num_pa_field = self.driver.find_element(By.XPATH, "//*[@id='numeroProcessoAquisicao']")
            num_pa_field.click()
            num_pa_field.send_keys(num_pa)
            
            data_field = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[2]/div[1]/form/div[4]/div[4]/div/div/div[1]/div/input")
            data_field.click()
            data_field.send_keys(data_pa.strftime('%d/%m/%Y') if isinstance(data_pa, datetime) else data_pa)
            
            valor_field = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[2]/div[1]/form/div[4]/div[5]/div/div/input")
            valor_field.click()
            valor_field.send_keys(str(valor_pa))
            
            return True
        except Exception as e:
            st.error(f"Erro ao preencher dados da PA: {str(e)}")
            return False
    
    def incluir_itens(self, df_itens, valor_item_diarias, des_item_pa):
        """Inclui os itens na PA baseado na tabela do Excel"""
        try:
            incluir_item_btn = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[1]/div[2]/a")
            incluir_item_btn.click()
            time.sleep(5)
            
            body = self.driver.find_element(By.TAG_NAME, 'body')
            for _ in range(18):
                body.send_keys(Keys.TAB)
            body.send_keys(Keys.ENTER)
            body.send_keys(Keys.END)
            body.send_keys(Keys.ENTER)
            time.sleep(5)
            
            item_checkbox = self.driver.find_element(By.XPATH, f"//*[@id='main-table']/div/div/div[1]/table/tbody/tr[{des_item_pa}]/td[1]/input")
            item_checkbox.click()
            time.sleep(3)
            
            confirmar_item = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[3]/button[1]")
            confirmar_item.click()
            time.sleep(2)
            
            incluir_item = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/button[1]")
            incluir_item.click()
            time.sleep(2)
            
            for idx, row in df_itens.iterrows():
                cpf_fornecedor = str(row.iloc[0]) if len(row) > 0 else ""
                nome_fornecedor = str(row.iloc[1]) if len(row) > 1 else ""
                quant_item = str(row.iloc[2]) if len(row) > 2 else "0"
                periodo_oficio = str(row.iloc[3]) if len(row) > 3 else ""
                num_oficio = str(row.iloc[4]) if len(row) > 4 else ""
                
                time.sleep(2)
                
                select_fornecedor = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[2]/div[1]/div/i")
                select_fornecedor.click()
                
                cpf_field = self.driver.find_element(By.ID, "cpfCnpj")
                cpf_field.click()
                cpf_field.send_keys(cpf_fornecedor)
                
                body.send_keys(Keys.TAB)
                body.send_keys(Keys.TAB)
                body.send_keys(Keys.ENTER)
                time.sleep(5)
                
                select_fornecedor_result = self.wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div/div/form/div/div/div[2]/div/div/div[1]/table/tbody/tr/td[1]/input")))
                select_fornecedor_result.click()
                
                confirmar_fornecedor = self.driver.find_element(By.XPATH, "/html/body/div[5]/div/div/form/div/button[1]")
                confirmar_fornecedor.click()
                
                descricao_field = self.driver.find_element(By.XPATH, "//*[@id='descricao']")
                descricao_field.click()
                descricao_field.send_keys(f"PAGAMENTO DE DIÁRIAS CONFORME OFÍCIO - {num_oficio} - PERÍODO {periodo_oficio} - {nome_fornecedor}")
                body.send_keys(Keys.TAB)
                
                quantidade_field = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div/form/div[2]/div[3]/div[2]/div/div/input")
                quantidade_field.click()
                quantidade_field.send_keys(quant_item)
                body.send_keys(Keys.TAB)
                
                valor_field = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div/form/div[2]/div[3]/div[3]/div/div/input")
                valor_field.click()
                valor_field.send_keys(str(valor_item_diarias))
                body.send_keys(Keys.TAB)
                
                select_unidade = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[2]/div/form/div[2]/div[3]/div[4]/div/i")
                select_unidade.click()
                time.sleep(5)
                
                body.send_keys(Keys.TAB)
                body.send_keys(Keys.TAB)
                body.send_keys(Keys.BACKSPACE)
                body.send_keys("unidade")
                body.send_keys(Keys.TAB)
                body.send_keys(Keys.ENTER)
                time.sleep(5)
                
                select_unidade_result = self.wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div/div/div[2]/div[2]/div/div/div[1]/table/tbody/tr[1]/td[1]/input")))
                select_unidade_result.click()
                time.sleep(5)
                
                confirmar_unidade = self.driver.find_element(By.XPATH, "/html/body/div[5]/div/div/div[3]/button[1]")
                confirmar_unidade.click()
                time.sleep(5)
                
                confirmar_item = self.driver.find_element(By.XPATH, "/html/body/div[3]/div/div/div[3]/button[1]")
                confirmar_item.click()
                
                if idx < len(df_itens) - 1:
                    incluir_item = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/button[1]")
                    incluir_item.click()
            
            body.send_keys(Keys.TAB)
            body.send_keys(Keys.ENTER)
            
            return True
        except Exception as e:
            st.error(f"Erro ao incluir itens: {str(e)}")
            return False
    
    def anexar_arquivo(self, caminho_arquivo):
        """Anexa arquivo à PA"""
        try:
            anexos_link = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[1]/div[1]/a")
            anexos_link.click()
            time.sleep(2)
            
            incluir_anexo = self.driver.find_element(By.XPATH, "//*[@id='content-wrap']/div/div[3]/div/div[2]/div[2]/div[1]/form/div[5]/div/div/button")
            incluir_anexo.click()
            time.sleep(2)
            
            nome_anexo = self.driver.find_element(By.XPATH, "//*[@id='nomeAnexo']")
            nome_anexo.click()
            nome_anexo.send_keys("PLANEJAMENTO")
            time.sleep(5)
            
            body = self.driver.find_element(By.TAG_NAME, 'body')
            body.send_keys(Keys.TAB)
            body.send_keys(Keys.TAB)
            body.send_keys(Keys.ENTER)
            time.sleep(5)
            
            return True
        except Exception as e:
            st.error(f"Erro ao anexar arquivo: {str(e)}")
            return False
    
    def fechar(self):
        """Fecha o driver"""
        if self.driver:
            self.driver.quit()

def main():
    st.title("💰 Importação e Pagamento de Diárias")
    st.markdown("---")
    
    # Inicializa session state
    inicializar_session_state()
    
    # Sidebar com informações
    with st.sidebar:
        st.header("ℹ️ Sobre")
        st.info(
            """
            Esta aplicação importa dados de diárias e mapeia para uma tabela específica.
            
            **Colunas de destino:**
            - CPF (apenas números)
            - FUNCIONÁRIO
            - QTD (preserva casas decimais)
            - PERÍODO
            - OFÍCIO
            - CONTRATO (Nº Ofício + Competência MM/AA) - apenas números
            - DL (Nº Ofício + numeração única) - apenas números
            - NOME RECIBO
            - TERMO DE COLABORAÇÃO
            - INSTRUMENTO
            """
        )
        
        st.header("📋 Instruções")
        st.markdown(
            """
            1. Faça upload do arquivo Excel
            2. Selecione a aba com os dados
            3. Mapeie as colunas de origem
            4. Clique em 'Processar Dados'
            5. Preencha os dados da PA
            6. Clique em 'SALVAR DADOS DA PA'
            7. Configure e execute a geração automática da PA
            """
        )
        
        # Botão para reset
        if st.button("🔄 Novo Processamento", use_container_width=True):
            for key in ['processamento_concluido', 'df_resultado', 'mapeamento_salvo', 
                       'colunas_disponiveis', 'processador', 'dados_pa', 'dados_pa_salvos']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
    
    # Área principal
    col1, col2 = st.columns([2, 1])
    
    with col1:
        arquivo = st.file_uploader(
            "📁 Selecione o arquivo Excel",
            type=['xlsx', 'xls', 'xlsm'],
            key="uploader_arquivo",
            help="Arquivo contendo os dados de diárias"
        )
    
    if arquivo is not None and not st.session_state.processamento_concluido:
        # Mostra as abas disponíveis
        try:
            excel_file = pd.ExcelFile(arquivo)
            abas = excel_file.sheet_names
            
            st.markdown("---")
            st.markdown("### 📑 Abas disponíveis no arquivo:")
            
            cols_abas = st.columns(3)
            for i, aba in enumerate(abas):
                with cols_abas[i % 3]:
                    st.write(f"📄 {aba}")
            
            st.markdown("---")
            st.markdown("### 🔄 Selecione a aba com os dados:")
            
            aba_selecionada = st.selectbox(
                "**Aba de dados:**",
                options=abas,
                index=0 if abas else None,
                key="select_aba_dados"
            )
            
            col_carregar1, col_carregar2, col_carregar3 = st.columns([1, 2, 1])
            with col_carregar2:
                if st.button("📊 Carregar e Mapear Colunas", type="secondary", use_container_width=True):
                    if aba_selecionada:
                        with st.spinner("Carregando arquivo..."):
                            processador = ProcessadorDiarias()
                            sucesso, colunas_encontradas = processador.carregar_arquivo(arquivo, aba_selecionada)
                            
                            if sucesso:
                                st.session_state.processador = processador
                                st.session_state.colunas_disponiveis = list(processador.df_importado.columns)
                                st.session_state.mapeamento_auto = colunas_encontradas
                                st.success("✅ Arquivo carregado! Faça o mapeamento das colunas abaixo.")
                                st.rerun()
            
            if st.session_state.processador and st.session_state.colunas_disponiveis:
                st.markdown("---")
                st.markdown("### 🔄 Mapeamento de Colunas")
                st.info("Selecione a coluna de origem para cada campo de destino:")
                
                processador = st.session_state.processador
                
                with st.expander("🔍 Prévia dos dados carregados"):
                    st.dataframe(processador.df_importado.head(10), use_container_width=True)
                
                colunas_origem = ["Selecione..."] + list(st.session_state.colunas_disponiveis)
                
                mapeamento = {}
                
                col_map1, col_map2, col_map3 = st.columns(3)
                
                with col_map1:
                    st.markdown("**Campos de destino (Grupo 1):**")
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('CPF', 'Selecione...')
                    mapeamento['CPF'] = st.selectbox(
                        "CPF (será limpo - só números):",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_cpf"
                    )
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Funcionário', 'Selecione...')
                    mapeamento['Funcionário'] = st.selectbox(
                        "FUNCIONÁRIO:",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_funcionario"
                    )
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Quantidade', 'Selecione...')
                    mapeamento['Quantidade'] = st.selectbox(
                        "QTD:",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_qtd"
                    )
                
                with col_map2:
                    st.markdown("**Campos de destino (Grupo 2):**")
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Período', 'Selecione...')
                    mapeamento['Período'] = st.selectbox(
                        "PERÍODO:",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_periodo"
                    )
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Ofício', 'Selecione...')
                    mapeamento['Ofício'] = st.selectbox(
                        "OFÍCIO:",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_oficio"
                    )
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Nº do Ofício', 'Selecione...')
                    mapeamento['Nº do Ofício'] = st.selectbox(
                        "Nº do Ofício (para CONTRATO e DL - será limpo):",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_num_oficio"
                    )
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Data Recibo', 'Selecione...')
                    mapeamento['Data Recibo'] = st.selectbox(
                        "Data Recibo (para CONTRATO - será convertida para MM/AA):",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_data"
                    )
                
                with col_map3:
                    st.markdown("**Campos de destino (Grupo 3):**")
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Nome do Recibo', 'Selecione...')
                    mapeamento['Nome do Recibo'] = st.selectbox(
                        "NOME RECIBO:",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_nome_recibo"
                    )
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Termo de Colaboração', 'Selecione...')
                    mapeamento['Termo de Colaboração'] = st.selectbox(
                        "TERMO DE COLABORAÇÃO:",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_termo_colaboracao"
                    )
                    
                    valor_sugerido = st.session_state.mapeamento_auto.get('Instrumento', 'Selecione...')
                    mapeamento['Instrumento'] = st.selectbox(
                        "INSTRUMENTO:",
                        options=colunas_origem,
                        index=colunas_origem.index(valor_sugerido) if valor_sugerido in colunas_origem else 0,
                        key="map_instrumento"
                    )
                
                st.markdown("---")
                col_botao1, col_botao2, col_botao3 = st.columns([1, 2, 1])
                
                with col_botao2:
                    if st.button("🚀 PROCESSAR DADOS", type="primary", use_container_width=True):
                        obrigatorias = ['CPF', 'Funcionário', 'Nº do Ofício', 'Data Recibo']
                        faltantes = [col for col in obrigatorias if mapeamento.get(col) == "Selecione..."]
                        
                        if faltantes:
                            st.error(f"❌ Colunas obrigatórias não mapeadas: {', '.join(faltantes)}")
                        else:
                            with st.spinner("🔄 Processando dados..."):
                                df_resultado = processador.processar_dados(mapeamento)
                                
                                if df_resultado is not None and len(df_resultado) > 0:
                                    st.session_state.df_resultado = df_resultado
                                    st.session_state.mapeamento_salvo = mapeamento
                                    st.session_state.processamento_concluido = True
                                    
                                    st.success(f"✅ {len(df_resultado)} registros processados com sucesso!")
                                    st.balloons()
                                    st.rerun()
                                else:
                                    st.error("❌ Nenhum dado foi processado. Verifique o mapeamento.")
                        
        except Exception as e:
            st.error(f"❌ Erro ao processar arquivo: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
    
    # Mostra resultado se existir
    elif st.session_state.processamento_concluido and st.session_state.df_resultado is not None:
        st.markdown("---")
        st.markdown("### 📊 Resultado do Processamento")
        
        df_resultado = st.session_state.df_resultado
        
        # Garante que QTD seja numérico com casas decimais
        df_resultado['QTD'] = pd.to_numeric(df_resultado['QTD'], errors='coerce').fillna(0)
        
        # Mostra o DataFrame com formatação
        with st.expander("🔍 Visualizar dados processados", expanded=True):
            # Formata a coluna QTD para mostrar com 2 casas decimais
            df_display = df_resultado.copy()
            if 'QTD' in df_display.columns:
                df_display['QTD'] = df_display['QTD'].apply(lambda x: f"{x:.2f}".replace('.', ','))
            st.dataframe(
                df_display,
                use_container_width=True,
                height=400
            )
            
            # Mostra estatísticas da coluna QTD
            st.info(f"📊 **Estatísticas da coluna QTD:**")
            st.write(f"- Total de registros: {len(df_resultado)}")
            st.write(f"- Soma total: {df_resultado['QTD'].sum():.2f}".replace('.', ','))
            st.write(f"- Média: {df_resultado['QTD'].mean():.2f}".replace('.', ','))
            st.write(f"- Mínimo: {df_resultado['QTD'].min():.2f}".replace('.', ','))
            st.write(f"- Máximo: {df_resultado['QTD'].max():.2f}".replace('.', ','))
        
        # Métricas
        col_metric1, col_metric2, col_metric3, col_metric4 = st.columns(4)
        
        with col_metric1:
            st.metric("Total de Registros", len(df_resultado))
        
        with col_metric2:
            st.metric("Total de Colunas", len(df_resultado.columns))
        
        with col_metric3:
            cpf_validos = df_resultado['CPF'].notna().sum()
            st.metric("Registros com CPF", cpf_validos)
        
        with col_metric4:
            contratos_validos = df_resultado['CONTRATO'].notna().sum()
            st.metric("Contratos Gerados", contratos_validos)
        
        # =====================================================================
        # SEÇÃO 1: DADOS DA PRESTAÇÃO DE CONTAS (PA) - CORRIGIDA
        # =====================================================================
        st.markdown("---")
        st.markdown("### 📝 Dados da Prestação de Contas (PA)")
        st.markdown("Preencha os dados da PA abaixo:")
        
        col_form1, col_form2 = st.columns(2)
        
        with col_form1:
            st.text_input(
                "**TIPO DA PA**",
                value="COMPRA DIRETA",
                disabled=True,
                key="tipo_pa_display"
            )
            
            termo_valor = df_resultado['TERMO_COLABORACAO'].iloc[0] if 'TERMO_COLABORACAO' in df_resultado.columns and len(df_resultado) > 0 else ''
            termo_colaboracao = st.text_input(
                "**TERMO DE COLABORAÇÃO**",
                value=str(termo_valor) if termo_valor else '',
                placeholder="Valor importado da planilha",
                key="termo_colaboracao_input",
                disabled=True
            )
            
            instrumento_valor = df_resultado['INSTRUMENTO'].iloc[0] if 'INSTRUMENTO' in df_resultado.columns and len(df_resultado) > 0 else ''
            instrumento = st.text_input(
                "**INSTRUMENTO**",
                value=str(instrumento_valor) if instrumento_valor else '',
                placeholder="Valor importado da planilha",
                key="instrumento_input",
                disabled=True
            )
            
            numero_pa = st.text_input(
                "**NÚMERO DA PA**",
                value=st.session_state.dados_pa.get('numero_pa', ''),
                placeholder="Digite o número da PA",
                key="numero_pa_input"
            )
            
            data_pa = st.date_input(
                "**DATA DA PA**",
                value=st.session_state.dados_pa.get('data_pa', datetime.now().date()),
                format="DD/MM/YYYY",
                key="data_pa_input"
            )
        
        with col_form2:
            valor_diaria = st.number_input(
                "**VALOR DA DIÁRIA (R$)**",
                min_value=0.0,
                value=float(st.session_state.dados_pa.get('valor_diaria', 0.0)),
                step=0.01,
                format="%.2f",
                key="valor_diaria_input"
            )
            
            # CORREÇÃO: Calcula o valor da PA corretamente
            qtd_total = float(df_resultado['QTD'].sum())
            valor_pa_calculado = qtd_total * valor_diaria
            
            # Atualiza o session state com o valor calculado
            st.session_state.ultimo_calculo_pa = valor_pa_calculado
            
            # Mostra o cálculo detalhado
            st.markdown(f"""
            **🔢 Detalhamento do Cálculo:**
            - Total de diárias: {qtd_total:.2f}
            - Valor unitário: R$ {valor_diaria:.2f}
            - **VALOR TOTAL: R$ {valor_pa_calculado:,.2f}**
            """.replace('.', ','))
            
            # Campo de exibição do valor da PA
            valor_pa_display = st.text_input(
                "**VALOR DA PA (R$)**",
                value=formatar_moeda_br(valor_pa_calculado),
                disabled=True,
                key="valor_pa_display"
            )
            
            item_aquisicao = st.text_input(
                "**ITEM DE AQUISIÇÃO**",
                value=st.session_state.dados_pa.get('item_aquisicao', ''),
                placeholder="Digite o item de aquisição",
                key="item_aquisicao_input"
            )
            
            mes_competencia = data_pa.strftime("%m") if data_pa else ""
            ano_competencia = data_pa.strftime("%Y") if data_pa else ""
            
            st.text_input(
                "**MÊS COMPETÊNCIA**",
                value=mes_competencia,
                disabled=True,
                key="mes_competencia_display"
            )
            
            st.text_input(
                "**ANO COMPETÊNCIA**",
                value=ano_competencia,
                disabled=True,
                key="ano_competencia_display"
            )
        
        col_salvar1, col_salvar2, col_salvar3 = st.columns([1, 2, 1])
        with col_salvar2:
            if st.button("💾 SALVAR DADOS DA PA", type="primary", use_container_width=True):
                st.session_state.dados_pa.update({
                    'termo_colaboracao': termo_valor,
                    'instrumento': instrumento_valor,
                    'numero_pa': numero_pa,
                    'data_pa': data_pa,
                    'valor_diaria': valor_diaria,
                    'item_aquisicao': item_aquisicao,
                    'valor_pa': valor_pa_calculado
                })
                st.session_state.dados_pa_salvos = True
                st.success("✅ Dados da PA salvos com sucesso!")
                st.rerun()
        
        if st.session_state.dados_pa_salvos:
            with st.expander("📋 Resumo dos Dados da PA", expanded=True):
                col_resumo1, col_resumo2 = st.columns(2)
                
                with col_resumo1:
                    st.markdown(f"**TIPO DA PA:** COMPRA DIRETA")
                    st.markdown(f"**TERMO DE COLABORAÇÃO:** {st.session_state.dados_pa.get('termo_colaboracao', '')}")
                    st.markdown(f"**INSTRUMENTO:** {st.session_state.dados_pa.get('instrumento', '')}")
                    st.markdown(f"**NÚMERO DA PA:** {st.session_state.dados_pa.get('numero_pa', '')}")
                
                with col_resumo2:
                    data_pa_formatada = st.session_state.dados_pa.get('data_pa', '').strftime('%d/%m/%Y') if st.session_state.dados_pa.get('data_pa') else ''
                    st.markdown(f"**DATA DA PA:** {data_pa_formatada}")
                    st.markdown(f"**VALOR DA DIÁRIA:** {formatar_moeda_br(st.session_state.dados_pa.get('valor_diaria', 0))}")
                    st.markdown(f"**VALOR DA PA:** {formatar_moeda_br(st.session_state.dados_pa.get('valor_pa', 0))}")
                    st.markdown(f"**ITEM DE AQUISIÇÃO:** {st.session_state.dados_pa.get('item_aquisicao', '')}")
                    st.markdown(f"**MÊS/ANO COMPETÊNCIA:** {mes_competencia}/{ano_competencia}")
        
        # =====================================================================
        # SEÇÃO 2: GERAR PA DE DIÁRIAS
        # =====================================================================
        st.markdown("---")
        st.markdown("### 🤖 Gerar PA de Diárias")
        st.markdown("Configure e execute a geração automática da PA no sistema e-Parcerias:")
        
        # Criar abas para organizar as configurações
        tab1, tab2, tab3, tab4 = st.tabs(["🔐 Login", "📋 Dados PA", "📊 Itens", "📎 Anexos"])
        
        with tab1:
            st.subheader("Credenciais de Acesso")
            col_login1, col_login2 = st.columns(2)
            
            with col_login1:
                usuario = st.text_input(
                    "**Usuário (CPF)**",
                    value="613.324.373-20",
                    key="usuario_login"
                )
            
            with col_login2:
                senha = st.text_input(
                    "**Senha**",
                    value="silviarac@cfis2025",
                    type="password",
                    key="senha_login"
                )
        
        with tab2:
            st.subheader("Dados da PA no Sistema")
            
            # Pré-carregar com dados salvos
            num_instrumento_sugerido = st.session_state.dados_pa.get('instrumento', '1')
            num_pa_sugerido = st.session_state.dados_pa.get('numero_pa', '2024.0001')
            data_pa_sugerida = st.session_state.dados_pa.get('data_pa', datetime.now())
            valor_pa_sugerido = st.session_state.dados_pa.get('valor_pa', 1500.00)
            
            col_pa1, col_pa2 = st.columns(2)
            
            with col_pa1:
                num_instrumento = st.text_input(
                    "**Número do Instrumento**",
                    value=num_instrumento_sugerido,
                    key="num_instrumento"
                )
                
                num_pa = st.text_input(
                    "**Número da PA**",
                    value=num_pa_sugerido,
                    key="num_pa"
                )
            
            with col_pa2:
                data_pa_sistema = st.date_input(
                    "**Data da PA**",
                    value=data_pa_sugerida,
                    format="DD/MM/YYYY",
                    key="data_pa_sistema"
                )
                
                valor_pa_sistema = st.number_input(
                    "**Valor da PA (R$)**",
                    min_value=0.0,
                    value=float(valor_pa_sugerido),
                    step=0.01,
                    format="%.2f",
                    key="valor_pa_sistema"
                )
        
        with tab3:
            st.subheader("Configuração dos Itens")
            
            # Mostrar prévia dos itens que serão incluídos
            st.info("📊 **Itens a serem incluídos na PA:**")
            
            # Preparar DataFrame para visualização com casas decimais
            df_itens_preview = df_resultado[['CPF', 'FUNCIONÁRIO', 'QTD', 'PERÍODO', 'OFÍCIO']].copy()
            df_itens_preview['QTD'] = df_itens_preview['QTD'].apply(lambda x: f"{x:.2f}".replace('.', ','))
            df_itens_preview.columns = ['CPF/CNPJ', 'Nome', 'Quantidade', 'Período', 'Nº Ofício']
            st.dataframe(df_itens_preview, use_container_width=True)
            
            st.markdown("---")
            
            col_item1, col_item2 = st.columns(2)
            
            with col_item1:
                des_item_pa = st.text_input(
                    "**Índice do Item de Aquisição**",
                    value=st.session_state.des_item_pa,
                    key="des_item_pa_input",
                    help="Número do item na lista de aquisição do sistema"
                )
                st.session_state.des_item_pa = des_item_pa
                
                mes_diarias = st.text_input(
                    "**Mês de Competência**",
                    value=mes_competencia if mes_competencia else "01",
                    key="mes_diarias"
                )
            
            with col_item2:
                ano_diarias = st.text_input(
                    "**Ano de Competência**",
                    value=ano_competencia if ano_competencia else "2024",
                    key="ano_diarias"
                )
                
                valor_item = st.number_input(
                    "**Valor do Item (Diárias) R$**",
                    min_value=0.0,
                    value=float(st.session_state.dados_pa.get('valor_diaria', 500.00)),
                    step=0.01,
                    format="%.2f",
                    key="valor_item",
                    help="Valor unitário de cada diária"
                )
        
        with tab4:
            st.subheader("Configuração de Anexos")
            
            caminho_anexo = st.text_input(
                "**Caminho do arquivo PDF para anexar**",
                value="C:/Users/VINICIUS GUANABARA/Desktop/GESTOR_FINANCEIRO/ANEXOS",
                key="caminho_anexo",
                help="Caminho completo do arquivo PDF a ser anexado à PA"
            )
            
            # Verificar se o arquivo existe
            if caminho_anexo and os.path.exists(caminho_anexo):
                st.success(f"✅ Arquivo encontrado: {os.path.basename(caminho_anexo)}")
            else:
                st.warning("⚠️ Arquivo não encontrado no caminho especificado")
        
        st.markdown("---")
        
        # Resumo das configurações
        with st.expander("📋 Resumo das Configurações para Geração da PA"):
            col_resumo_pa1, col_resumo_pa2, col_resumo_pa3 = st.columns(3)
            
            with col_resumo_pa1:
                st.markdown("**🔐 Login:**")
                st.markdown(f"- Usuário: {usuario}")
                st.markdown(f"- Senha: {'*' * len(senha)}")
            
            with col_resumo_pa2:
                st.markdown("**📋 Dados PA:**")
                st.markdown(f"- Instrumento: {num_instrumento}")
                st.markdown(f"- Nº PA: {num_pa}")
                st.markdown(f"- Data: {data_pa_sistema.strftime('%d/%m/%Y')}")
                st.markdown(f"- Valor: R$ {valor_pa_sistema:.2f}")
            
            with col_resumo_pa3:
                st.markdown("**📊 Itens:**")
                st.markdown(f"- Total de itens: {len(df_resultado)}")
                st.markdown(f"- Valor unitário: R$ {valor_item:.2f}")
                st.markdown(f"- Item índice: {st.session_state.des_item_pa}")
                st.markdown(f"- Competência: {mes_diarias}/{ano_diarias}")
        
        # Botão para executar a geração automática
        col_exec1, col_exec2, col_exec3 = st.columns([1, 2, 1])
        with col_exec2:
            if st.button("🚀 GERAR PA AUTOMATICAMENTE", type="primary", use_container_width=True):
                with st.spinner("Processando... Isso pode levar alguns minutos..."):
                    gerador = GerarPAPessoal()
                    
                    try:
                        # Iniciar driver
                        st.info("1️⃣ Iniciando navegador...")
                        gerador.iniciar_driver()
                        
                        # Login
                        st.info("2️⃣ Fazendo login no sistema...")
                        if not gerador.fazer_login(usuario, senha):
                            raise Exception("Falha no login")
                        
                        # Navegar
                        st.info("3️⃣ Navegando para processo de aquisição...")
                        if not gerador.navegar_para_processo_aquisicao():
                            raise Exception("Falha na navegação")
                        
                        # Selecionar instrumento
                        st.info("4️⃣ Selecionando instrumento...")
                        if not gerador.selecionar_instrumento(num_instrumento):
                            raise Exception("Falha ao selecionar instrumento")
                        
                        # Preencher dados PA
                        st.info("5️⃣ Preenchendo dados da PA...")
                        if not gerador.preencher_dados_pa(num_pa, data_pa_sistema, valor_pa_sistema):
                            raise Exception("Falha ao preencher dados da PA")
                        
                        # Incluir itens
                        st.info("6️⃣ Incluindo itens na PA...")
                        if not gerador.incluir_itens(df_resultado, valor_item, st.session_state.des_item_pa):
                            raise Exception("Falha ao incluir itens")
                        
                        # Anexar arquivo
                        st.info("7️⃣ Anexando arquivo...")
                        if not gerador.anexar_arquivo(caminho_anexo):
                            raise Exception("Falha ao anexar arquivo")
                        
                        st.success("✅ PA gerada com sucesso!")
                        st.balloons()
                        
                    except Exception as e:
                        st.error(f"❌ Erro durante o processo: {str(e)}")
                    
                    finally:
                        gerador.fechar()
        
        # =====================================================================
        # ESTATÍSTICAS E DOWNLOAD
        # =====================================================================
        
        with st.expander("📈 Estatísticas Detalhadas"):
            st.write("**Informações do DataFrame:**")
            string_buffer = StringIO()
            df_resultado.info(buf=string_buffer)
            st.text(string_buffer.getvalue())
            
            st.write("**Primeiras 10 linhas:**")
            df_display_stats = df_resultado.head(10).copy()
            if 'QTD' in df_display_stats.columns:
                df_display_stats['QTD'] = df_display_stats['QTD'].apply(lambda x: f"{x:.2f}".replace('.', ','))
            st.dataframe(df_display_stats)
            
            st.write("**Exemplos de CONTRATO e DL (apenas números):**")
            exemplos = df_resultado[['CONTRATO', 'DL']].head(10)
            st.dataframe(exemplos)
            
            st.write("**Valores nulos por coluna:**")
            null_counts = df_resultado.isnull().sum()
            null_counts = null_counts[null_counts > 0]
            if len(null_counts) > 0:
                st.dataframe(null_counts)
            else:
                st.write("✅ Sem valores nulos")
        
        st.markdown("---")
        col_download1, col_download2, col_download3 = st.columns([1, 2, 1])
        
        with col_download2:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Aba principal com os dados processados
                df_resultado.to_excel(
                    writer,
                    index=False,
                    sheet_name="DIARIAS_PROCESSADAS"
                )
                
                # Aba com os dados da PA (se foram salvos)
                if st.session_state.dados_pa_salvos:
                    dados_pa_df = pd.DataFrame([{
                        'TIPO_PA': 'COMPRA DIRETA',
                        'TERMO_COLABORACAO': st.session_state.dados_pa.get('termo_colaboracao', ''),
                        'INSTRUMENTO': st.session_state.dados_pa.get('instrumento', ''),
                        'NUMERO_PA': st.session_state.dados_pa.get('numero_pa', ''),
                        'DATA_PA': st.session_state.dados_pa.get('data_pa', ''),
                        'VALOR_DIARIA': st.session_state.dados_pa.get('valor_diaria', 0),
                        'VALOR_PA': st.session_state.dados_pa.get('valor_pa', 0),
                        'ITEM_AQUISICAO': st.session_state.dados_pa.get('item_aquisicao', ''),
                        'MES_COMPETENCIA': mes_competencia,
                        'ANO_COMPETENCIA': ano_competencia,
                        'QTD_TOTAL_DIARIAS': qtd_total
                    }])
                    dados_pa_df.to_excel(
                        writer,
                        index=False,
                        sheet_name="DADOS_PA"
                    )
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            st.download_button(
                label="📥 DOWNLOAD ARQUIVO PROCESSADO",
                data=output.getvalue(),
                file_name=f"diarias_pa_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

if __name__ == "__main__":
    main()