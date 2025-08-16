import streamlit as st
import os
import config

st.set_page_config(page_title="Dados Detalhados", page_icon="📄", layout="wide")
st.title("📄 Dados Detalhados e Downloads")
def colorir_status(val):
    color = 'white'
    if val == 'UP': color = '#92D050'
    elif val == 'PQ': color = '#FF0000'
    elif val == 'PP': color = '#9BC2E6'
    elif val == 'SD': color = '#FFEB9C'
    return f'background-color: {color}'

if 'resultados_gerados' not in st.session_state or st.session_state.resultados_gerados is None:
    st.warning("Por favor, gere um relatório na Página Principal primeiro para visualizar os dados.")
else:
    st.header("Tabela de Atividades do Mês")
    
    df_preview = st.session_state.resultados_gerados.get('df_preview')
    if df_preview is not None:
        styled_df = df_preview.style.map(colorir_status)
        st.dataframe(styled_df, use_container_width=True)
    
    st.divider()
    st.subheader("Downloads")
    col_down1, col_down2 = st.columns(2)
    
    with col_down1:
        try:
            with open(config.PROCESSED_CSV_PATH, 'r', encoding='utf-8') as f: csv_data = f.read()
            st.download_button("⬇️ Baixar Dados Processados (.csv)", csv_data, config.PROCESSED_CSV_FILENAME, 'text/csv', use_container_width=True)
        except FileNotFoundError:
            st.warning("Arquivo CSV processado não encontrado.")
    
    with col_down2:
        caminho_excel = st.session_state.resultados_gerados['caminho_excel']
        with open(caminho_excel, "rb") as file:
            st.download_button("⬇️ Baixar Relatório Excel Completo", file, os.path.basename(caminho_excel), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.divider()
if st.button("⬅️ Voltar ao Menu Principal"):
    st.switch_page("app.py")