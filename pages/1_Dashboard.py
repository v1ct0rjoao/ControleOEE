import streamlit as st
import plotly.graph_objects as go

st.set_page_config(page_title="Dashboard do Mês", page_icon="📊", layout="wide")
st.title("📊 Dashboard do Mês")

if 'resultados_gerados' not in st.session_state or st.session_state.resultados_gerados is None:
    st.warning("Por favor, gere um relatório na Página Principal primeiro para visualizar o dashboard.")
else:
    sumario = st.session_state.resultados_gerados['sumario']
    oee_value = sumario.get('OEE', 0)
    
    st.header(f"OEE Final do Mês")

    fig_gauge = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = oee_value,
        number = {'suffix': '%', 'font': {'size': 60}},
        domain = {'x': [0, 1], 'y': [0, 1]},
        gauge = {
            'shape': "angular",
            'axis': {'range': [None, 100]},
            'bar': {'color': "#006B3D", 'thickness': 0.75}
        }))
    
    fig_gauge.update_layout(height=400, margin=dict(l=20, r=20, t=50, b=20))
    st.plotly_chart(fig_gauge, use_container_width=True)

st.divider()
if st.button("⬅️ Voltar ao Menu Principal"):
    st.switch_page("app.py")