import streamlit as st
import pandas as pd
import os
from datetime import datetime
from modulos.processamento import limpar_dados_brutos, gerar_dashboard_oee
import config

st.set_page_config(page_title="Dashboard OEE", layout="wide")

if 'processamento_concluido' not in st.session_state:
    st.session_state.processamento_concluido = False
if 'nomes_arquivos_processados' not in st.session_state:
    st.session_state.nomes_arquivos_processados = []
if 'caminho_relatorio_gerado' not in st.session_state:
    st.session_state.caminho_relatorio_gerado = None
if 'lista_de_circuitos' not in st.session_state:
    st.session_state.lista_de_circuitos = []

with st.sidebar:
    st.title("⚙️ Controles")

    st.header("Passo 1: Enviar Dados")
    uploaded_files = st.file_uploader(
        "Selecione os arquivos .txt", type="txt", accept_multiple_files=True, key="file_uploader"
    )

    if st.session_state.processamento_concluido:
        st.divider()
        st.header("Passo 2: Gerar Relatório")
        
        agora = datetime.now()
        ano_desejado = st.number_input("Ano", min_value=2020, max_value=agora.year + 1, value=agora.year)
        
        meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        mes_selecionado = st.selectbox("Mês", options=meses_pt, index=agora.month - 1)
        mes_desejado = meses_pt.index(mes_selecionado) + 1

        st.divider()
        st.subheader("Opções Avançadas")
        
        circuitos_selecionados_force = st.multiselect(
            "Forçar status para circuitos:",
            options=st.session_state.lista_de_circuitos,
            help="Selecione circuitos para aplicar uma regra de status manual."
        )
        
        tipo_force = st.radio(
            "Tipo de regra manual:",
            ["Nenhuma", "Forçar 100% UP", "Forçar Semana Padrão (Seg-Sex UP)"],
            horizontal=True, key="tipo_force",
            help="A regra será aplicada apenas nos circuitos selecionados acima."
        )
        
        if st.button("Gerar Relatório Excel OEE", use_container_width=True, type="primary"):
            tipo_force_final = tipo_force if tipo_force != "Nenhuma" else None
            with st.spinner('Criando seu relatório Excel... 📊'):
                caminho_relatorio = gerar_dashboard_oee(
                    arquivo_entrada_path=config.PROCESSED_CSV_PATH,
                    arquivo_saida_folder=config.OUTPUT_FOLDER,
                    ano=ano_desejado,
                    mes=mes_desejado,
                    circuitos_para_forcar=circuitos_selecionados_force,
                    tipo_de_force=tipo_force_final
                )
            st.session_state.caminho_relatorio_gerado = caminho_relatorio
    else:
        st.info("Aguardando o envio de arquivos para habilitar as opções.")

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
        
        sucesso, lista_circuitos = limpar_dados_brutos(lista_paths, config.PROCESSED_CSV_PATH)
    
    if sucesso:
        st.session_state.processamento_concluido = True
        st.session_state.nomes_arquivos_processados = nomes_arquivos_atuais
        st.session_state.lista_de_circuitos = lista_circuitos
        st.session_state.caminho_relatorio_gerado = None
        st.rerun()
    else:
        st.error("Falha ao processar os arquivos. Verifique o formato dos dados.")
        st.session_state.processamento_concluido = False
        st.session_state.nomes_arquivos_processados = []
        st.session_state.lista_de_circuitos = []

if st.session_state.processamento_concluido:
    st.success("Dados brutos processados com sucesso! ✔️")
    
    try:
        with open(config.PROCESSED_CSV_PATH, 'r', encoding='utf-8') as f:
            csv_data = f.read()
        st.download_button(
           label="⬇️ Baixar arquivo de dados processado (.csv)",
           data=csv_data,
           file_name=config.PROCESSED_CSV_FILENAME,
           mime='text/csv',
        )
    except FileNotFoundError:
        st.error("Não foi possível encontrar o arquivo CSV para download.")

if st.session_state.caminho_relatorio_gerado:
    st.success("Relatório Excel gerado! ✔️")
    with open(st.session_state.caminho_relatorio_gerado, "rb") as file:
        st.download_button(
            label="Clique para baixar o Relatório (.xlsx)",
            data=file,
            file_name=os.path.basename(st.session_state.caminho_relatorio_gerado),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif not uploaded_files and st.session_state.processamento_concluido:
    st.session_state.processamento_concluido = False
    st.session_state.nomes_arquivos_processados = []
    st.session_state.caminho_relatorio_gerado = None
    st.session_state.lista_de_circuitos = []
    st.rerun()