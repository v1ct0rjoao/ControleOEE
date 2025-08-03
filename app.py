import streamlit as st
import pandas as pd
import os
from datetime import datetime
from modulos.processamento import limpar_dados_brutos, gerar_dashboard_oee
import config

st.set_page_config(page_title="Dashboard OEE", layout="wide")

# --- Inicialização do Estado da Sessão ---
if 'processamento_concluido' not in st.session_state:
    st.session_state.processamento_concluido = False
if 'nomes_arquivos_processados' not in st.session_state:
    st.session_state.nomes_arquivos_processados = []
if 'lista_de_circuitos' not in st.session_state:
    st.session_state.lista_de_circuitos = []
if 'resultados_gerados' not in st.session_state:
    st.session_state.resultados_gerados = None

# --- Função para colorir a tabela de pré-visualização ---
def colorir_status(val):
    color = 'white'
    if val == 'UP': color = '#92D050'
    elif val == 'PQ': color = '#FF0000'
    elif val == 'PP': color = '#9BC2E6'
    elif val == 'SD': color = '#FFEB9C'
    return f'background-color: {color}'

# --- Barra Lateral (Sidebar) ---
with st.sidebar:
    st.title("⚙️ Controles")
    st.header("Passo 1: Enviar Dados")
    uploaded_files = st.file_uploader(
        "Selecione os arquivos .txt", type="txt", accept_multiple_files=True, key="file_uploader"
    )

    formato_data_selecionado = st.radio(
        "Formato da data nos arquivos .txt:",
        options=["dd/mm/aaaa", "mm/dd/aaaa"],
        horizontal=True,
        help="Escolha o formato de data correspondente aos seus arquivos de entrada."
    )

    if st.session_state.processamento_concluido:
        st.divider()
        st.header("Passo 2: Gerar Relatório")
        
        agora = datetime.now()
        ano_desejado = st.number_input("Ano", min_value=2020, max_value=agora.year + 5, value=agora.year)
        
        meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        mes_selecionado = st.selectbox("Mês", options=meses_pt, index=agora.month - 1)
        mes_desejado = meses_pt.index(mes_selecionado) + 1

        capacidade_total_input = st.number_input(
            "Capacidade Total de Circuitos:", min_value=1, value=375, step=1,
            help="Informe o número total de circuitos disponíveis neste mês."
        )

        st.divider()
        st.subheader("Entradas de OEE")
        
        ensaios_solicitados_input = st.number_input("Ensaios Solicitados (C):", min_value=0, value=0, help="Número total de ensaios solicitados.")
        ensaios_executados_input = st.number_input("Ensaios Executados (D):", min_value=0, value=0, help="Número de ensaios executados.")
        relatorios_emitidos_input = st.number_input("Relatórios Emitidos (E):", min_value=0, value=0, help="Número total de relatórios emitidos.")
        relatorios_no_prazo_input = st.number_input("Relatórios no Prazo (F):", min_value=0, value=0, help="Número de relatórios emitidos no prazo.")

        st.divider()
        st.subheader("Opções Avançadas")
        
      
        aplicar_min_dias_up = st.checkbox(
            "Aplicar regra de mínimo de dias 'UP'?", value=False,
            help="Se desmarcado, todos os circuitos com alguma atividade serão considerados."
        )
        
        min_dias_up_input = 1
        if aplicar_min_dias_up:
            min_dias_up_input = st.number_input(
                "Mínimo de dias 'UP' para uso:", min_value=1, value=1, step=1,
                help="Circuitos com menos dias 'UP' do que este valor serão considerados não utilizados."
            )
        
        st.write("**Forçar Circuitos como Produtivos (UP):**")
        circuitos_force_up = st.multiselect("Selecione os circuitos:", options=st.session_state.lista_de_circuitos, key='force_up_select')
        tipo_force_up = st.radio("Tipo de regra 'UP':", ["Forçar 100% UP", "Forçar Semana Padrão (Seg-Sex UP)"], key="tipo_force_up")

        st.write("**Forçar Circuitos como Quebrados (PQ):**")
        circuitos_force_pq = st.multiselect("Selecione os circuitos:", options=st.session_state.lista_de_circuitos, key='force_pq_select')
        
        st.write("**Forçar Circuitos como Vazios:**")
        circuitos_force_vazio = st.multiselect("Selecione os circuitos para remover do relatório:", options=st.session_state.lista_de_circuitos, key='force_vazio_select')
        
        if st.button("Gerar Relatório e Visualização", use_container_width=True, type="primary"):
            regras_de_force = {
                'circuitos_up': circuitos_force_up,
                'tipo_up': tipo_force_up,
                'circuitos_pq': circuitos_force_pq,
                'circuitos_vazio': circuitos_force_vazio
            }
            with st.spinner('Criando seu relatório e visualização... 📊'):
                resultados = gerar_dashboard_oee(
                    arquivo_entrada_path=config.PROCESSED_CSV_PATH,
                    arquivo_saida_folder=config.OUTPUT_FOLDER,
                    ano=ano_desejado,
                    mes=mes_desejado,
                    capacidade_total=capacidade_total_input,
                    regras_de_force=regras_de_force,
                    min_dias_up=min_dias_up_input,
                    aplicar_min_dias_up=aplicar_min_dias_up,
                    ensaios_executados=ensaios_executados_input,
                    ensaios_solicitados=ensaios_solicitados_input,
                    relatorios_no_prazo=relatorios_no_prazo_input,
                    relatorios_emitidos=relatorios_emitidos_input
                )
            st.session_state.resultados_gerados = resultados
            st.rerun()
    else:
        st.info("Aguardando o envio de arquivos para habilitar as opções.")

# --- Página Principal ---
st.title("🤖 Dashboard de Automação de OEE")
st.markdown("Use os controles na barra lateral à esquerda para gerar seus relatórios.")
st.divider()

nomes_arquivos_atuais = sorted([f.name for f in uploaded_files])
if uploaded_files and nomes_arquivos_atuais != st.session_state.nomes_arquivos_processados:
    with st.spinner('Salvando e processando arquivos... ⏳'):
        lista_paths = [os.path.join(config.UPLOAD_FOLDER, f.name) for f in uploaded_files]
        for uploaded_file, path in zip(uploaded_files, lista_paths):
            with open(path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        
        sucesso, lista_circuitos = limpar_dados_brutos(lista_paths, config.PROCESSED_CSV_PATH, formato_data=formato_data_selecionado)
    
    if sucesso:
        st.session_state.processamento_concluido = True
        st.session_state.nomes_arquivos_processados = nomes_arquivos_atuais
        st.session_state.lista_de_circuitos = lista_circuitos
        st.session_state.resultados_gerados = None
        st.rerun()
    else:
        st.error("Falha ao processar os arquivos. Verifique o formato dos dados.")
        st.session_state.processamento_concluido = False
        st.session_state.nomes_arquivos_processados = []
        st.session_state.lista_de_circuitos = []

# --- Seção de Resultados e Visualização ---
if st.session_state.processamento_concluido:
    st.subheader("Downloads")
    col1, col2 = st.columns(2)
    with col1:
        try:
            with open(config.PROCESSED_CSV_PATH, 'r', encoding='utf-8') as f: csv_data = f.read()
            st.download_button("⬇️ Baixar Dados Processados (.csv)", csv_data, config.PROCESSED_CSV_FILENAME, 'text/csv', use_container_width=True)
        except FileNotFoundError: st.error("Arquivo CSV não encontrado.")

    if st.session_state.resultados_gerados and st.session_state.resultados_gerados.get('caminho_excel'):
        with col2:
            caminho_excel = st.session_state.resultados_gerados['caminho_excel']
            with open(caminho_excel, "rb") as file:
                st.download_button("⬇️ Baixar Relatório Completo (.xlsx)", file, os.path.basename(caminho_excel), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    
    st.divider()

    if st.session_state.resultados_gerados and st.session_state.resultados_gerados.get('sumario'):
        st.subheader("Resumo do OEE")
        sumario = st.session_state.resultados_gerados['sumario']
        
        col_oee1, col_oee2, col_oee3, col_oee4 = st.columns(4)
        
        col_oee1.metric(label="Disponibilidade", value=f"{sumario.get('Disponibilidade', 0):.2f}%")
        col_oee2.metric(label="Performance", value=f"{sumario.get('Performance', 0):.2f}%")
        col_oee3.metric(label="Qualidade", value=f"{sumario.get('Qualidade', 0):.2f}%")
        col_oee4.metric(label="OEE Final", value=f"{sumario.get('OEE', 0):.2f}%")

    if st.session_state.resultados_gerados and st.session_state.resultados_gerados.get('df_preview') is not None:
        st.subheader("Pré-visualização do Relatório")
        df_preview = st.session_state.resultados_gerados.get('df_preview')
        with st.expander("Clique para ver a pré-visualização do preenchimento"):
            styled_df = df_preview.style.map(colorir_status)
            st.dataframe(styled_df, use_container_width=True)

elif not uploaded_files and st.session_state.processamento_concluido:
    st.session_state.processamento_concluido = False
    st.session_state.nomes_arquivos_processados = []
    st.session_state.resultados_gerados = None
    st.session_state.lista_de_circuitos = []
    st.rerun()