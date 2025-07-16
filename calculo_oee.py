import pandas as pd
import re
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

def gerar_relatorio_oee_excel(arquivo_entrada, ano, mes):
    """
    Lê um arquivo CSV, gera um relatório de OEE diário, e trata corretamente
    atividades em andamento vs. atividades antigas sem data de fim.
    """
    print("\n--- INICIANDO SCRIPT DE RELATÓRIO OEE (LÓGICA DEFINITIVA) ---")
    
    agora = pd.Timestamp.now().tz_localize(None).normalize()

    try:
        print(f"Lendo o arquivo de dados processados: '{arquivo_entrada}'...")
        df_atividades = pd.read_csv(
            arquivo_entrada, 
            sep=';',
            dtype={'circuito': str},
            date_format='%Y-%m-%d %H:%M:%S',
            parse_dates=['datastart', 'datastop']
        )
        print(f"Sucesso! {len(df_atividades)} registros totais lidos.")
    except FileNotFoundError:
        print(f"\nERRO: Arquivo '{arquivo_entrada}' não encontrado!")
        return
    except Exception as e:
        print(f"\nERRO ao ler o arquivo CSV: {e}")
        return

    inicio_mes = pd.to_datetime(f"{ano}-{mes}-01").tz_localize(None)
    fim_mes = (inicio_mes + pd.offsets.MonthEnd(0)).tz_localize(None)
    
    df_periodo = df_atividades[
        (df_atividades['datastart'] <= fim_mes) & 
        ((df_atividades['datastop'] >= inicio_mes) | (pd.isna(df_atividades['datastop'])))
    ].copy()

    if df_periodo.empty:
        print(f"Nenhuma atividade encontrada para o período de {mes:02d}/{ano}.")
        return

    circuitos_unicos = sorted(df_periodo['circuito'].unique(), key=lambda x: int(re.search(r'\d+', x).group()))
    dias_do_mes = pd.date_range(start=inicio_mes, end=fim_mes, freq='D')
    
    print(f"\nGerando calendário de status para {len(circuitos_unicos)} circuitos potenciais...")
    
    dados_calendario = []
    for circuito in circuitos_unicos:
        for dia in dias_do_mes:
            status = 'Parada Planejada' if dia.weekday() >= 5 else 'Sem Demanda'
            dados_calendario.append({'Circuito': circuito, 'Data': dia, 'Status': status})
    
    calendario_oee_df = pd.DataFrame(dados_calendario).set_index(['Circuito', 'Data'])

    for _, atividade in df_periodo.iterrows():
        
        # --- LÓGICA DEFINITIVA PARA DATAS DE INÍCIO E FIM ---
        start_iter = max(atividade['datastart'].normalize(), inicio_mes)

        if pd.notna(atividade['datastop']):
            # CASO 1: A atividade tem data de fim. Lógica normal.
            end_iter = min(atividade['datastop'].normalize(), fim_mes)
        else:
            # CASO 2: A atividade NÃO tem data de fim.
            # Se a atividade começou ANTES do mês do relatório, é um erro de dado.
            # Contamos apenas 1 dia para não poluir o relatório.
            if atividade['datastart'] < inicio_mes:
                end_iter = start_iter 
            else:
                # Se começou NESTE mês, está em andamento. Contamos até hoje.
                end_iter = min(agora, fim_mes)
        # --- FIM DA LÓGICA DEFINITIVA ---

        if end_iter < start_iter:
            end_iter = start_iter
            
        for dia in pd.date_range(start_iter, end_iter, freq='D'):
            if (atividade['circuito'], dia) in calendario_oee_df.index:
                calendario_oee_df.loc[(atividade['circuito'], dia), 'Status'] = 'Uso Programado'

    circuitos_com_uso_real = calendario_oee_df[
        calendario_oee_df['Status'] == 'Uso Programado'
    ].index.get_level_values('Circuito').unique()

    if circuitos_com_uso_real.empty:
        print(f"Nenhum circuito teve 'Uso Programado' no período de {mes:02d}/{ano}.")
        return

    calendario_oee_df = calendario_oee_df.loc[circuitos_com_uso_real]
    print(f"Filtro aplicado! O relatório será gerado para {len(circuitos_com_uso_real)} circuitos com atividade real no mês.")
    
    relatorio_detalhado_df = calendario_oee_df.unstack(level='Data')
    relatorio_detalhado_df.columns = relatorio_detalhado_df.columns.droplevel(0)
    relatorio_detalhado_df.columns = relatorio_detalhado_df.columns.strftime('%d/%m/%Y')
    relatorio_detalhado_df.index.name = 'Circuito'

    resumo_oee = calendario_oee_df.groupby('Circuito')['Status'].value_counts().unstack(fill_value=0)
    
    for status_col in ['Uso Programado', 'Sem Demanda', 'Parada Planejada']:
        if status_col not in resumo_oee.columns:
            resumo_oee[status_col] = 0
            
    print("\n--- Relatório de OEE por Circuito (Contagem de Dias) ---")
    print(resumo_oee[['Uso Programado', 'Sem Demanda', 'Parada Planejada']].to_string())
    
    nome_arquivo_saida = f"relatorio_oee_visual_{ano}-{mes:02d}.xlsx"
    print(f"\nSalvando e formatando o relatório em '{nome_arquivo_saida}'...")

    with pd.ExcelWriter(nome_arquivo_saida, engine='openpyxl') as writer:
        relatorio_detalhado_df.to_excel(writer, sheet_name='Relatorio_OEE_Diario')
        resumo_oee[['Uso Programado', 'Sem Demanda', 'Parada Planejada']].to_excel(writer, sheet_name='Resumo_Contagem')
        
        ws_detalhado = writer.sheets['Relatorio_OEE_Diario']
        fill_up = PatternFill(fill_type="solid", start_color="C6EFCE")
        fill_sd = PatternFill(fill_type="solid", start_color="FFEB9C")
        fill_pp = PatternFill(fill_type="solid", start_color="FFC7CE")

        max_row = ws_detalhado.max_row
        max_col_letter = get_column_letter(ws_detalhado.max_column)
        cell_range = f"B2:{max_col_letter}{max_row}"

        ws_detalhado.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"Uso Programado"'], fill=fill_up))
        ws_detalhado.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"Sem Demanda"'], fill=fill_sd))
        ws_detalhado.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"Parada Planejada"'], fill=fill_pp))

        ws_detalhado.column_dimensions['A'].width = 15
        for col_idx in range(2, ws_detalhado.max_column + 1):
            ws_detalhado.column_dimensions[get_column_letter(col_idx)].width = 15
            
        ws_resumo = writer.sheets['Resumo_Contagem']
        ws_resumo.column_dimensions['A'].width = 15
        for col_letter in ['B', 'C', 'D']:
            ws_resumo.column_dimensions[col_letter].width = 20

    print(f"\n✅ Relatório OEE visual salvo com sucesso!")

if __name__ == "__main__":
    arquivo_de_dados = "dados_processados.csv"
    if not os.path.exists(arquivo_de_dados):
        print(f"ERRO: Arquivo de entrada '{arquivo_de_dados}' não encontrado.")
    else:
        try:
            ano_atual = datetime.now().year
            mes_atual = datetime.now().month
            ano_desejado_input = input(f"Digite o ano (padrão: {ano_atual}): ")
            ano_desejado = int(ano_desejado_input) if ano_desejado_input else ano_atual
            mes_desejado_input = input(f"Digite o mês (padrão: {mes_atual}): ")
            mes_desejado = int(mes_desejado_input) if mes_desejado_input else mes_atual
            if not (1 <= mes_desejado <= 12):
                raise ValueError("Mês inválido.")
        except ValueError as e:
            print(f"Entrada inválida. {e}.")
        else:
            gerar_relatorio_oee_excel(arquivo_de_dados, ano_desejado, mes_desejado)