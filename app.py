import streamlit as st
import os
from datetime import datetime
from modulos.processamento import limpar_dados_brutos, gerar_dashboard_oee
import config
import pandas as pd

st.set_page_config(
    page_title="Página Principal - OEE Moura",
    page_icon="🤖",
    layout="wide"
)

# Inicialização do session_state
if 'processamento_concluido' not in st.session_state:
    st.session_state.processamento_concluido = False
if 'lista_de_circuitos' not in st.session_state:
    st.session_state.lista_de_circuitos = []
if 'resultados_gerados' not in st.session_state:
    st.session_state.resultados_gerados = None

st.markdown("""
<style>
    /* Estilo para o container do cartão */
    .card {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
        transition: 0.3s;
        text-align: center;
    }
    .card:hover {
        box-shadow: 0 8px 16px 0 rgba(0,0,0,0.2);
    }
    /* Remove a decoração do link no título */
    .card h3 a {
        text-decoration: none;
        color: #262730 !important;
    }
</style>
""", unsafe_allow_html=True)


# --- Corpo Principal com Navegação ---
if 'resultados_gerados' in st.session_state and st.session_state.resultados_gerados is not None:
    st.title("✅ Relatório Gerado com Sucesso!")
    st.markdown("### Selecione uma das opções abaixo para visualizar os resultados:")
else:
    st.title("🤖 Dashboard de Automação de OEE")
    st.markdown("### Bem-vindo! Comece gerando um relatório na barra lateral.")

st.divider()

col1, col2, col3 = st.columns(3, gap="large")

with col1:
    st.markdown('<div class="card"><h3>📊 Dashboard do Mês</h3></div>', unsafe_allow_html=True)
    if st.button("Acessar Dashboard", key="btn_dash", use_container_width=True):
        st.switch_page("pages/1_Dashboard.py")

with col2:
    st.markdown('<div class="card"><h3>📈 Análise Histórica</h3></div>', unsafe_allow_html=True)
    if st.button("Analisar Histórico", key="btn_hist", use_container_width=True):
        st.switch_page("pages/2_Historico.py")

with col3:
    st.markdown('<div class="card"><h3>📄 Dados e Downloads</h3></div>', unsafe_allow_html=True)
    if st.button("Ver Dados Detalhados", key="btn_dados", use_container_width=True):
        st.switch_page("pages/3_Dados.py")


# --- Barra Lateral (Sidebar) com a Lógica de Processamento ---
with st.sidebar:
    st.title("⚙️ Controles")
    st.header("Passo 1: Enviar e Processar Dados")
    uploaded_files = st.file_uploader(
        "Selecione os arquivos .txt", type="txt", accept_multiple_files=True, key="file_uploader"
    )

    formato_data_selecionado = st.radio(
        "Formato da data nos arquivos:",
        options=["dd/mm/aaaa", "mm/dd/aaaa"],
        horizontal=True
    )
    
 
    if st.button("Processar Arquivos", use_container_width=True):
        if uploaded_files:
            with st.spinner('Salvando e processando arquivos... ⏳'):
                lista_paths = [os.path.join(config.UPLOAD_FOLDER, f.name) for f in uploaded_files]
                for uploaded_file, path in zip(uploaded_files, lista_paths):
                    with open(path, "wb") as f: f.write(uploaded_file.getbuffer())
                
                sucesso, lista_circuitos = limpar_dados_brutos(lista_paths, config.PROCESSED_CSV_PATH, formato_data=formato_data_selecionado)
            
            if sucesso:
                st.session_state.processamento_concluido = True
                st.session_state.lista_de_circuitos = lista_circuitos
                st.session_state.resultados_gerados = None
                st.success("Arquivos processados! Pronto para gerar o relatório.")
            else:
                st.error("Falha ao processar os arquivos.")
                st.session_state.processamento_concluido = False
                st.session_state.lista_de_circuitos = []
        else:
            st.warning("Por favor, envie pelo menos um arquivo .txt para processar.")


    if st.session_state.get('processamento_concluido'):
        st.divider()
        st.header("Passo 2: Gerar Relatório Final")
        
        agora = datetime.now()
        ano_desejado = st.number_input("Ano", min_value=2020, max_value=agora.year + 5, value=agora.year)
        
        meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        mes_selecionado = st.selectbox("Mês", options=meses_pt, index=agora.month - 1)
        mes_desejado = meses_pt.index(mes_selecionado) + 1

        capacidade_total_input = st.number_input(
            "Capacidade Total de Circuitos:", min_value=1, value=375, step=1
        )
        st.divider()
        st.subheader("Entradas de OEE")
        ensaios_solicitados_input = st.number_input("Ensaios Solicitados (C):", min_value=0, value=0)
        ensaios_executados_input = st.number_input("Ensaios Executados (D):", min_value=0, value=0)
        relatorios_emitidos_input = st.number_input("Relatórios Emitidos (E):", min_value=0, value=0)
        relatorios_no_prazo_input = st.number_input("Relatórios no Prazo (F):", min_value=0, value=0)

        st.divider()
        with st.expander("Opções Avançadas"):
            aplicar_min_dias_up = st.checkbox("Aplicar regra de mínimo de dias 'UP'?", value=False)
            min_dias_up_input = 1
            if aplicar_min_dias_up:
                min_dias_up_input = st.number_input("Mínimo de dias 'UP' para uso:", min_value=1, value=1, step=1)
            
            st.write("**Forçar Circuitos como Produtivos (UP):**")
            circuitos_force_up = st.multiselect("Selecione:", options=st.session_state.get('lista_de_circuitos', []), key='force_up_select')
            tipo_force_up = st.radio("Tipo de regra 'UP':", ["Forçar 100% UP", "Forçar Semana Padrão (Seg-Sex UP)"], key="tipo_force_up")

            st.write("**Forçar Circuitos como Quebrados (PQ):**")
            circuitos_force_pq = st.multiselect("Selecione:", options=st.session_state.get('lista_de_circuitos', []), key='force_pq_select')
            
            st.write("**Remover Circuitos do Relatório:**")
            circuitos_force_vazio = st.multiselect("Selecione:", options=st.session_state.get('lista_de_circuitos', []), key='force_vazio_select')
        
        if st.button("Gerar Relatório e Dashboard", use_container_width=True, type="primary"):
            regras_de_force = {'circuitos_up': circuitos_force_up, 'tipo_up': tipo_force_up, 'circuitos_pq': circuitos_force_pq, 'circuitos_vazio': circuitos_force_vazio}
            with st.spinner('Criando seu relatório e visualização... 📊'):
                resultados = gerar_dashboard_oee(
                    arquivo_entrada_path=config.PROCESSED_CSV_PATH, arquivo_saida_folder=config.OUTPUT_FOLDER,
                    ano=ano_desejado, mes=mes_desejado, capacidade_total=capacidade_total_input,
                    regras_de_force=regras_de_force, min_dias_up=min_dias_up_input, aplicar_min_dias_up=aplicar_min_dias_up,
                    ensaios_executados=ensaios_executados_input, ensaios_solicitados=ensaios_solicitados_input,
                    relatorios_no_prazo=relatorios_no_prazo_input, relatorios_emitidos=relatorios_emitidos_input
                )
            st.session_state.resultados_gerados = resultados
            st.rerun()

    st.divider()
    st.header("Visualizar Mês Anterior")
    caminho_csv = os.path.join(config.OUTPUT_FOLDER, 'historico_oee.csv')
    if os.path.exists(caminho_csv):
        df_historico = pd.read_csv(caminho_csv)
        meses_pt_map = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
        df_historico['display'] = df_historico.apply(lambda row: f"{meses_pt_map.get(row['mes'], '')}/{row['ano']}", axis=1)
        
        mes_ano_selecionado = st.selectbox(
            "Selecione um mês do histórico:",
            options=df_historico['display'].unique()
        )

        if st.button("Carregar Mês", use_container_width=True):
            if mes_ano_selecionado:
                display_selecionado = df_historico[df_historico['display'] == mes_ano_selecionado].iloc[0]
                ano, mes = int(display_selecionado['ano']), int(display_selecionado['mes'])
                
                with st.spinner(f"Recalculando dados para {mes_ano_selecionado}..."):
                    resultados = gerar_dashboard_oee(
                        arquivo_entrada_path=config.PROCESSED_CSV_PATH, arquivo_saida_folder=config.OUTPUT_FOLDER,
                        ano=ano, mes=mes
                    )
                st.session_state.resultados_gerados = resultados
                st.success(f"Dados de {mes_ano_selecionado} carregados!")
                st.switch_page("pages/1_Dashboard.py")
    else:
        st.info("Nenhum histórico encontrado. Gere um relatório para começar.")