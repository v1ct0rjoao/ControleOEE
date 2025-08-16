import streamlit as st
import pandas as pd
import os
import config
import plotly.express as px

st.set_page_config(page_title="Análise Histórica", page_icon="📈", layout="wide")
st.title("📈 Análise Histórica")

caminho_csv = os.path.join(config.OUTPUT_FOLDER, 'historico_oee.csv')

if os.path.exists(caminho_csv):
    try:
        df_historico = pd.read_csv(caminho_csv)
        df_historico['periodo'] = pd.to_datetime(df_historico['ano'].astype(str) + '-' + df_historico['mes'].astype(str) + '-01')
        df_historico = df_historico.sort_values('periodo')
        
        opcoes_grafico = {
            'OEE Final': 'oee_final',
            'Disponibilidade': 'disponibilidade',
            'Performance': 'performance',
            'Qualidade': 'qualidade'
        }
        
        metricas_selecionadas_nomes = st.multiselect(
            "Selecione os indicadores para visualizar:",
            options=list(opcoes_grafico.keys()),
            default=['OEE Final']
        )
        
        metricas_selecionadas_cols = [opcoes_grafico[nome] for nome in metricas_selecionadas_nomes]

        if metricas_selecionadas_cols:
            df_melted = df_historico.melt(id_vars='periodo', value_vars=metricas_selecionadas_cols, var_name='Indicador', value_name='Valor')
            
            fig_historico = px.bar(
                df_melted, x='periodo', y='Valor', color='Indicador',
                barmode='group', text_auto='.2f', title="Histórico de Indicadores OEE"
            )
            fig_historico.update_traces(textposition='outside')
            fig_historico.update_layout(yaxis_title="Percentual (%)", xaxis_title="Período", yaxis=dict(range=[0, 105]))
            st.plotly_chart(fig_historico, use_container_width=True)
        else:
            st.warning("Por favor, selecione pelo menos um indicador para visualizar o gráfico.")

        st.divider()
        with st.expander("Gerenciar Histórico de Dados"):
            st.subheader("Apagar meses específicos do histórico")
            
            df_historico['mes_ano_str'] = df_historico['periodo'].dt.strftime('%m/%Y')
            meses_para_apagar = st.multiselect(
                "Selecione os meses que deseja remover:",
                options=df_historico['mes_ano_str'].unique()
            )

            if st.button("Apagar Meses Selecionados", type="primary"):
                if meses_para_apagar:
                    df_filtrado = df_historico[~df_historico['mes_ano_str'].isin(meses_para_apagar)]
                    df_filtrado.drop(columns=['periodo', 'mes_ano_str']).to_csv(caminho_csv, index=False)
                    st.success(f"Meses {', '.join(meses_para_apagar)} foram removidos! A página será recarregada.")
                    st.rerun()
                else:
                    st.warning("Nenhum mês foi selecionado para remoção.")

    except Exception as e:
        st.error(f"Não foi possível ler o arquivo de histórico CSV: {e}")
else:
    st.info("O arquivo de histórico (historico_oee.csv) será criado na pasta 'relatorios' assim que você gerar o primeiro relatório.")

st.divider()
if st.button("⬅️ Voltar ao Menu Principal"):
    st.switch_page("app.py")