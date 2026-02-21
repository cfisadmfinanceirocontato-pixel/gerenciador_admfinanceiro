import streamlit as st
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple

# ============================================================================
# CONFIGURAÇÕES E CONSTANTES
# ============================================================================

# Configuração da página
PAGE_CONFIG = {
    "page_title": "Cfis App Financeiro",
    "page_icon": "💰",
    "layout": "wide",
    "initial_sidebar_state": "expanded"
}

# Definição da estrutura de navegação
NAVIGATION_STRUCTURE: Dict[str, List[Tuple[str, str]]] = {
    "Página Inicial": [
        ("Home", "pages/home.py")
    ],
    "Cadastros": [
        ("Instrumentos", "pages/cadastrosinstrumentos.py"),
        ("Fornecedores", "pages/pgfornecedores.py"),
        ("Funcionários", "pages/pgfuncionarios.py")
    ],
    "Dashboard Financeiro": [
        ("Dashboard financeiro", "pages/dashboard.py"),
        ("Demonstrativo financeiro", "pages/pginstrumentos.py"),
        ("Cronograma de Repasse", "pages/pgrepasses.py")
    ],
    "Provisionamento Mensal": [
        ("Provisionamento mensal", "pages/pgprovisionamento.py"),
        ("Provisionamento mensal custo direto", "pages/pgprovisionamentocd.py"),
        ("Provisionamento mensal veiculos custo direto", "pages/pgveiculos.py"),
        ("Provisionamento mensal enventos custo direto", "pages/pgeventoscd.py"),
        ("Provisionamento mensal manutenções", "pages/pgmanutencoes.py"),
        ("Provisionamento mensal serviços eventuais", "pages/pgservieventuais.py"),
        ("Provisionamento mensal veiculos custo indireto", "pages/pgvprovisionamentoci.py"),
        ("Pagamento de consumo", "pages/pgconsumo.py"),
        ("Pagamento de serviços", "pages/pgservicos.py")
    ],
    "Diárias": [
        ("Controle de Diárias", "pages/pghomediarias.py"),
        ("Relatórios de Diárias", "pages/pgrelatoriosdiarias.py"),
        ("Recibos de Diárias", "pages/pgdiarias.py"),
        ("Pagamentos de Diárias", "pages/pgpgtodiarias.py")
    ],
    "Pessoal e Encargos": [
        ("Custos de pessoal", "pages/pgcustodepessoal.py"),
        ("Provisionamento de pessoal", "pages/pgprovisionamentopessoal.py"),
        ("Pagamento de pessoal", "pages/pgpessoal.py")
    ]
}

# ============================================================================
# FUNÇÕES DE UTILIDADE
# ============================================================================

def create_navigation_pages(nav_structure: Dict[str, List[Tuple[str, str]]]) -> Dict:
    """
    Cria as páginas de navegação baseado na estrutura fornecida.
    
    Args:
        nav_structure: Dicionário com a estrutura de navegação
        
    Returns:
        Dicionário com as páginas configuradas para o st.navigation
    """
    navigation_dict = {}
    
    for section, pages in nav_structure.items():
        section_pages = []
        for title, path in pages:
            # Corrige possíveis erros de digitação nos nomes dos arquivos
            corrected_path = correct_file_path(path)
            page = st.Page(corrected_path, title=title)
            section_pages.append(page)
        navigation_dict[section] = section_pages
    
    return navigation_dict

def correct_file_path(path: str) -> str:
    """
    Corrige possíveis erros de digitação nos caminhos dos arquivos.
    
    Args:
        path: Caminho do arquivo original
        
    Returns:
        Caminho corrigido do arquivo
    """
    # Mapeamento de correções comuns
    corrections = {
        "pgvprovisionamentoci.py": "pgprovisionamentoci.py",  # Remove 'v' extra
        # Adicione outras correções conforme necessário
    }
    
    filename = path.split("/")[-1]
    if filename in corrections:
        return path.replace(filename, corrections[filename])
    
    return path

def initialize_session_state():
    """
    Inicializa variáveis de estado da sessão se não existirem.
    """
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if 'user' not in st.session_state:
        st.session_state.user = None
    
    if 'current_page' not in st.session_state:
        st.session_state.current_page = None

def setup_custom_css():
    """
    Aplica CSS customizado para melhorar a aparência.
    """
    st.markdown("""
        <style>
        .main-header {
            font-size: 2rem;
            color: #1E88E5;
            margin-bottom: 1rem;
        }
        .stButton button {
            background-color: #1E88E5;
            color: white;
        }
        </style>
    """, unsafe_allow_html=True)

# ============================================================================
# COMPONENTES DA INTERFACE
# ============================================================================

def render_sidebar_info():
    """
    Renderiza informações adicionais na barra lateral.
    """
    with st.sidebar:
        st.markdown("---")
        st.markdown("### ℹ️ Informações")
        st.info("Sistema de gestão financeira CFIS")
        
        # Exibe usuário logado (se implementar autenticação)
        if st.session_state.get('user'):
            st.markdown(f"**Usuário:** {st.session_state.user}")
        
        # Versão do sistema
        st.caption("Versão 1.0.0")

# ============================================================================
# FUNÇÃO PRINCIPAL
# ============================================================================

def main():
    """
    Função principal que inicializa e executa o aplicativo.
    """
    # Configuração inicial da página
    st.set_page_config(**PAGE_CONFIG)
    
    # Inicializa o estado da sessão
    initialize_session_state()
    
    # Aplica CSS customizado
    setup_custom_css()
    
    # Título principal (opcional, comentado para não interferir com o conteúdo das páginas)
    # st.markdown('<h1 class="main-header">CFIS App Financeiro</h1>', unsafe_allow_html=True)
    
    # Cria as páginas de navegação
    pages = create_navigation_pages(NAVIGATION_STRUCTURE)
    
    # Configura e executa a navegação
    pg = st.navigation(pages)
    
    # Renderiza informações adicionais na sidebar
    render_sidebar_info()
    
    # Executa a página selecionada
    pg.run()

# ============================================================================
# PONTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
    main()