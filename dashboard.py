import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# --- Configurações da Página ---
st.set_page_config(
    page_title="Dashboard de Projetos - PSO",
    page_icon="📊",
    layout="wide"
)

# --- Carregamento e Preparação dos Dados ---
@st.cache_data
def carregar_dados():
    caminho_do_arquivo = "SISTEMA PSO.xlsx"
    colunas_responsaveis = [
        'Engineering', 'Information Technology', 'Qtech', 
        'Maintenance', 'Safety', 'Finance', 'Environment'
    ]
    
    try:
        df = pd.read_excel(caminho_do_arquivo, header=3)
        
        cols_to_drop = [col for col in df.columns if 'Unnamed:' in str(col)]
        df = df.drop(columns=cols_to_drop)
            
    except FileNotFoundError:
        st.error(f"ERRO: Arquivo não encontrado no caminho: '{caminho_do_arquivo}'.")
        return None, []
    except Exception as e:
        st.error(f"Ocorreu um erro ao ler o arquivo Excel: {e}")
        return None, []
        
    if 'LEGENDA CORES' in df.columns:
        df['LEGENDA CORES'] = df['LEGENDA CORES'].astype(str)
        
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].fillna('')
            
    if 'CADASTRADOS' not in df.columns:
        st.error("ERRO: A coluna 'CADASTRADOS' não foi encontrada no arquivo Excel.")
        df['CADASTRADOS'] = ''
    
    custo_col = 'CUSTO'
    if custo_col in df.columns:
        df['CUSTO_NUMERICO'] = pd.to_numeric(df[custo_col], errors='coerce').fillna(0)
    else:
        st.warning(f"Atenção: A coluna '{custo_col}' não foi encontrada.")
        df['CUSTO_NUMERICO'] = 0

    return df, colunas_responsaveis

df, colunas_responsaveis = carregar_dados()
if df is None:
    st.stop()


# --- Barra Lateral (Sidebar) ---
st.sidebar.header("Filtros")

if 'SOLICITANTE' in df.columns:
    lista_solicitantes = df['SOLICITANTE'][df['SOLICITANTE'] != ''].unique()
    lista_ordenada = sorted(list(lista_solicitantes))
else:
    st.sidebar.error("Coluna 'SOLICITANTE' não encontrada no arquivo.")
    lista_ordenada = []

lista_ordenada.insert(0, "Visão Geral")

selecao = st.sidebar.selectbox(
    "Selecione um Engenheiro para Análise:",
    options=lista_ordenada
)


# --- Painel Principal ---
if selecao == "Visão Geral":
    st.title("Visão Geral de Todos os Projetos")
    
    # --- INÍCIO DA ALTERAÇÃO ---
    
    # Cálculos das métricas
    total_projetos = len(df)
    custo_total = df['CUSTO_NUMERICO'].sum()
    numero_aprovados = df[df['CADASTRADOS'].str.upper() == 'OK'].shape[0]
    numero_recusados = df[df['CADASTRADOS'].str.upper() == 'NOK'].shape[0] # Novo cálculo
    
    # Cálculo das porcentagens (com segurança para evitar divisão por zero)
    if total_projetos > 0:
        porcentagem_aprovados = (numero_aprovados / total_projetos) * 100
        porcentagem_recusados = (numero_recusados / total_projetos) * 100 # Novo cálculo
    else:
        porcentagem_aprovados = 0
        porcentagem_recusados = 0
    
    # Layout 3x2 para as métricas
    st.subheader("Indicadores Chave de Projetos (KPIs)")
    col1, col2, col3 = st.columns(3)
    col4, col5, col6 = st.columns(3)

    with col1:
        st.metric("Número Total de Projetos", f"{total_projetos}")
    with col2:
        st.metric("Projetos Aprovados (OK)", f"{numero_aprovados}")
    with col3:
        st.metric("Projetos não Lançados (NOK)", f"{numero_recusados}")

    with col4:
        st.metric("Custo Total", f"R$ {custo_total:,.2f}")
    with col5:
        st.metric("Percentual de Aprovados", f"{porcentagem_aprovados:.2f}%")
    with col6:
        st.metric("Percentual de Recusados", f"{porcentagem_recusados:.2f}%")
        
    # --- FIM DA ALTERAÇÃO ---
    
    st.markdown("---")
    
    st.header("Análises Gráficas")
    
    st.subheader("Distribuição de Custos por Categoria")
    custo_por_categoria = df.groupby('CATEGORIA')['CUSTO_NUMERICO'].sum().reset_index()
    fig_pie = px.pie(custo_por_categoria, 
                     names='CATEGORIA', 
                     values='CUSTO_NUMERICO',
                     title='Percentual de Custo por Categoria de Projeto',
                     hole=.3)
    st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("---")

    st.subheader("Projetos por Engenheiro Principal")
    projetos_por_engenheiro = df[df['Engineering'] != ''].groupby('Engineering').size().reset_index(name='Contagem')
    projetos_por_engenheiro = projetos_por_engenheiro.sort_values(by='Contagem', ascending=False)
    fig_bar = px.bar(projetos_por_engenheiro, 
                     x='Engineering', 
                     y='Contagem',
                     title='Nº de Projetos por Engenheiro Principal',
                     text_auto=True)
    st.plotly_chart(fig_bar, use_container_width=True)
    
    st.markdown("---")
    st.header("Tabela Completa de Projetos")
    st.dataframe(df.drop(columns=['CUSTO_NUMERICO'], errors='ignore'))

else:
    # O código da visão por solicitante continua o mesmo
    st.title(f"Análise de Projetos de: {selecao}")
    df_filtrado = df[df['SOLICITANTE'] == selecao]

    total_projetos_eng = len(df_filtrado)
    custo_total_eng = df_filtrado['CUSTO_NUMERICO'].sum()
    
    col1, col2 = st.columns(2)
    col1.metric("Total de Projetos Solicitados", f"{total_projetos_eng}")
    col2.metric("Custo Total dos Projetos Solicitados", f"R$ {custo_total_eng:,.2f}")
    
    st.markdown("---")
    st.header("Detalhes dos Projetos")
    
    for index, row in df_filtrado.iterrows():
        projeto_titulo = row.get('PROJETO', 'Projeto sem nome')
        wbs_titulo = row.get('WBS (LCP)', 'N/A')
        with st.expander(f"**{projeto_titulo}** (WBS: {wbs_titulo})"):
            
            status_cadastro = row.get('CADASTRADOS', '').upper()
            if status_cadastro == 'OK':
                st.success("✅ Projeto Cadastrado no PSO")
            elif status_cadastro == 'NOK':
                st.error("❌ Não Cadastrado no PSO")
            else:
                st.warning("Status de Cadastro Pendente ou Inválido")
            st.markdown("---")
            
            def exibir_info(label, valor):
                if pd.isna(valor) or (isinstance(valor, str) and not valor.strip()):
                    st.markdown(f"**{label}:** <span style='color: orange; font-weight: bold;'>⚠️ PENDENTE</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"**{label}:** {valor}")
            
            c1, c2 = st.columns(2)
            with c1:
                exibir_info("Classificação", row.get('CLASSIFICAÇÃO'))
                exibir_info("Categoria", row.get('CATEGORIA'))
                exibir_info("Custo", f"R$ {row.get('CUSTO', 0):,.2f}")
            with c2:
                exibir_info("Duração", row.get('DURAÇÃO'))
                link = row.get('LINK DA PASTA')
                if link:
                    st.markdown(f"**Link da Pasta:**({link})")
                else:
                    exibir_info("Link da Pasta", link)
            
            st.markdown("---")
            st.markdown("**Equipe Envolvida no Projeto:**")
            for resp_col in colunas_responsaveis:
                if resp_col in row:
                    responsavel = row.get(resp_col)
                    if pd.isna(responsavel) or not str(responsavel).strip():
                        st.markdown(f"- {resp_col}: <span style='color: orange; font-weight: bold;'>⚠️ PENDENTE</span>", unsafe_allow_html=True)
                    else:
                        if responsavel == selecao:
                            st.markdown(f"- **{resp_col}:** {responsavel}")
                        else:
                            st.markdown(f"- {resp_col}: {responsavel}")