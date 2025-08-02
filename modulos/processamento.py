import pandas as pd
import re
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
import locale

try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    pass

def interpretar_data_flexivel(string_data):
    if pd.isna(string_data) or not isinstance(string_data, str):
        return pd.NaT
    string_data = string_data.strip()
    if re.search(r'\b(am|pm)\b', string_data, re.IGNORECASE):
        return pd.NaT
    try:
        return pd.to_datetime(string_data, format="%d/%m/%Y %H:%M:%S", errors='coerce')
    except ValueError:
        try:
            return pd.to_datetime(string_data, format="%d/%m/%Y %H:%M", errors='coerce')
        except ValueError:
            return pd.NaT

def limpar_dados_brutos(lista_arquivos_path, arquivo_saida_path):
    """
    Lê arquivos de texto, limpa os dados e retorna um booleano de sucesso
    e a lista de circuitos únicos encontrados.
    """
    all_lines = []
    for nome_arquivo in lista_arquivos_path:
        try:
            try:
                with open(nome_arquivo, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            except UnicodeDecodeError:
                with open(nome_arquivo, 'r', encoding='latin-1') as f:
                    lines = f.readlines()
            all_lines.extend(lines)
        except FileNotFoundError:
            continue
            
    if not all_lines: return (False, [])
    
    conteudo_total = "\n".join(all_lines)
    circuit_pattern = re.compile(r"Circuit\d+", re.IGNORECASE)
    datetime_pattern = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\b")
    data_list = []
    circuit_matches = list(circuit_pattern.finditer(conteudo_total))

    for i, match in enumerate(circuit_matches):
        circuito = match.group(0)
        start_pos = match.end()
        end_pos = circuit_matches[i+1].start() if i + 1 < len(circuit_matches) else len(conteudo_total)
        data_blob = conteudo_total[start_pos:end_pos]
        found_datetimes = datetime_pattern.findall(data_blob)
        datastart_str = found_datetimes[0] if len(found_datetimes) >= 1 else None
        datastop_str = found_datetimes[1] if len(found_datetimes) >= 2 else None
        if datastart_str:
            data_list.append({'circuito': circuito, 'datastart': datastart_str, 'datastop': datastop_str})
            
    if not data_list: return (False, [])
    
    df = pd.DataFrame(data_list)
    df['datastart'] = df['datastart'].apply(interpretar_data_flexivel)
    df['datastop'] = df['datastop'].apply(interpretar_data_flexivel)
    df.dropna(subset=['datastart'], inplace=True)
    df['circuito_num'] = df['circuito'].str.extract(r'(\d+)').astype(int)
    df.sort_values(by=['circuito_num', 'datastart'], ascending=[True, False], inplace=True)
    df.drop(columns=['circuito_num'], inplace=True)
    df.reset_index(drop=True, inplace=True)
    
    df.to_csv(arquivo_saida_path, index=False, sep=';', date_format='%d/%m/%Y %H:%M:%S')
    
    circuitos_unicos = sorted(df['circuito'].unique(), key=lambda x: int(re.search(r'\d+', x).group()))
    return (True, circuitos_unicos)

def gerar_dashboard_oee(arquivo_entrada_path, arquivo_saida_folder, ano, mes, circuitos_para_forcar=None, tipo_de_force=None):
    """
    Gera o arquivo Excel do OEE com lógicas avançadas de forçamento e formatação.
    """
    if circuitos_para_forcar is None:
        circuitos_para_forcar = []

    try:
        df_atividades = pd.read_csv(arquivo_entrada_path, sep=';', dtype={'circuito': str})
        if 'status' not in df_atividades.columns: df_atividades['status'] = 'UP'
    except Exception:
        return None

    inicio_mes = datetime(ano, mes, 1)
    fim_mes = (inicio_mes + pd.offsets.MonthEnd(0)).to_pydatetime()
    dias_do_mes_range = pd.date_range(start=inicio_mes, end=fim_mes, freq='D')

    todos_os_circuitos = sorted(df_atividades['circuito'].unique(), key=lambda x: int(re.search(r'\d+', x).group()))

    dados_calendario = []
    for circuito in todos_os_circuitos:
        for dia in dias_do_mes_range:
            status = 'PP' if dia.weekday() >= 5 else 'SD'
            dados_calendario.append({'Circuito': circuito, 'Data': dia, 'Status': status})

    if not dados_calendario: return None
    calendario_df = pd.DataFrame(dados_calendario).set_index(['Circuito', 'Data'])

    if circuitos_para_forcar and tipo_de_force:
        for circuito in circuitos_para_forcar:
            if tipo_de_force == "Forçar 100% UP":
                calendario_df.loc[circuito, 'Status'] = 'UP'
            elif tipo_de_force == "Forçar Semana Padrão (Seg-Sex UP)":
                for dia, _ in calendario_df.loc[circuito].iterrows():
                    status_forcado = 'UP' if dia.weekday() < 5 else 'PP'
                    calendario_df.loc[(circuito, dia), 'Status'] = status_forcado

    df_periodo = df_atividades[~df_atividades['circuito'].isin(circuitos_para_forcar)].copy()
    df_periodo = df_periodo[
        (df_periodo['datastart'].notna()) & (df_periodo['datastop'].notna()) &
        (df_periodo['datastart'] <= fim_mes) & (df_periodo['datastop'] >= inicio_mes)
    ]

    for _, atividade in df_periodo.iterrows():
        start_iter = max(atividade['datastart'].normalize(), pd.Timestamp(inicio_mes))
        end_iter = min(atividade['datastop'].normalize(), pd.Timestamp(fim_mes))
        for dia in pd.date_range(start_iter, end_iter, freq='D'):
            if (atividade['circuito'], dia) in calendario_df.index:
                calendario_df.loc[(atividade['circuito'], dia), 'Status'] = 'UP'

    relatorio_detalhado_df = calendario_df.unstack(level='Data')['Status']
    relatorio_detalhado_df.columns = relatorio_detalhado_df.columns.day

    for index, row in relatorio_detalhado_df.iterrows():
        if index not in circuitos_para_forcar and 'UP' not in row.values:
            relatorio_detalhado_df.loc[index] = ''

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Controle_OEE_{ano}_{mes:02d}"
    font_title = Font(name='Calibri', size=40, bold=True, color="FFFFFF")
    font_header_white = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    font_legend = Font(name='Calibri', size=11, bold=True, color="000000")
    font_month_year = Font(name='Calibri', size=36, bold=True, color="FFFFFF")
    font_bold_black = Font(name='Calibri', size=11, bold=True, color="000000")
    fill_dark_green = PatternFill(start_color="006B3D", end_color="006B3D", fill_type="solid")
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border_thin_side = Side(border_style="thin", color="D9D9D9")
    border_thin = Border(left=border_thin_side, right=border_thin_side, top=border_thin_side, bottom=border_thin_side)
    fill_up = PatternFill(fill_type="solid", start_color="92D050")
    fill_pq = PatternFill(fill_type="solid", start_color="FF0000")
    fill_pp = PatternFill(fill_type="solid", start_color="9BC2E6")
    fill_sd = PatternFill(fill_type="solid", start_color="FFEB9C")
    last_day = fim_mes.day
    last_col_calendar = last_day + 1
    last_col_letter_total = get_column_letter(last_col_calendar + 8)
    ws.merge_cells(f'A1:{last_col_letter_total}1')
    cell = ws['A1']; cell.value = "Controle de OEE - Laboratório"
    cell.fill = fill_dark_green; cell.font = font_title; cell.alignment = align_center
    ws.row_dimensions[1].height = 60
    ws.row_dimensions[3].height = 25
    legend_items = {"UP": fill_up, "PQ": fill_pq, "PP": fill_pp, "SD": fill_sd}
    start_col = 2
    for status_code, fill_color in legend_items.items():
        cell = ws.cell(row=3, column=start_col)
        cell.value = status_code; cell.fill = fill_color; cell.font = font_legend; cell.alignment = align_center
        ws.column_dimensions[get_column_letter(start_col)].width = 8
        start_col += 1
    capacidade_col_start = get_column_letter(last_col_calendar + 1)
    capacidade_col_end = get_column_letter(last_col_calendar + 8)
    ws.merge_cells(f'{capacidade_col_start}3:{capacidade_col_end}3')
    cell = ws[f'{capacidade_col_start}3']; cell.value = "Capacidade de utilização: 300 circuitos"
    cell.fill = fill_dark_green; cell.font = font_header_white; cell.alignment = align_center
    month_year_row = 5
    ws.row_dimensions[month_year_row].height = 50
    month_col_end = get_column_letter(last_col_calendar // 2)
    ws.merge_cells(f'A{month_year_row}:{month_col_end}{month_year_row}')
    cell = ws[f'A{month_year_row}']; cell.value = inicio_mes.strftime('%B').capitalize()
    cell.fill = fill_dark_green; cell.font = font_month_year; cell.alignment = align_center
    year_col_start = get_column_letter((last_col_calendar // 2) + 1)
    ws.merge_cells(f'{year_col_start}{month_year_row}:{get_column_letter(last_col_calendar)}{month_year_row}')
    cell = ws[f'{year_col_start}{month_year_row}']; cell.value = str(ano)
    cell.fill = fill_dark_green; cell.font = font_month_year; cell.alignment = align_center
    header_row_dias_semana, header_row_numeros = 6, 7
    ws.row_dimensions[header_row_dias_semana].height = 25
    ws.row_dimensions[header_row_numeros].height = 25
    ws.merge_cells('A6:A7'); cell = ws['A6']
    cell.value = "Circuitos"; cell.fill = fill_dark_green; cell.font = font_header_white; cell.alignment = align_center
    ws.column_dimensions['A'].width = 20
    dias_semana_map = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb']
    for i, dia in enumerate(dias_do_mes_range, start=2):
        col_letter = get_column_letter(i)
        ws.cell(row=header_row_dias_semana, column=i, value=dias_semana_map[int(dia.strftime('%w'))]).fill = fill_dark_green; ws.cell(row=header_row_dias_semana, column=i).font = font_header_white; ws.cell(row=header_row_dias_semana, column=i).alignment = align_center
        ws.cell(row=header_row_numeros, column=i, value=dia.day).fill = fill_dark_green; ws.cell(row=header_row_numeros, column=i).font = font_header_white; ws.cell(row=header_row_numeros, column=i).alignment = align_center
        ws.column_dimensions[col_letter].width = 5
    summary_start_col = last_col_calendar + 3
    ws.merge_cells(start_row=6, start_column=summary_start_col, end_row=6, end_column=summary_start_col + 3)
    cell = ws.cell(row=6, column=summary_start_col)
    cell.value = "Compilação"; cell.fill = fill_dark_green; cell.font = font_header_white; cell.alignment = align_center
    for i, header in enumerate(["UP", "PQ", "PP", "SD"]):
        col_idx = summary_start_col + i
        ws.cell(row=7, column=col_idx, value=header).fill = fill_dark_green; ws.cell(row=7, column=col_idx).font = font_header_white; ws.cell(row=7, column=col_idx).alignment = align_center
        ws.column_dimensions[get_column_letter(col_idx)].width = 7
    current_row = header_row_numeros + 1
    if not relatorio_detalhado_df.empty:
        for circuito, data_row in relatorio_detalhado_df.iterrows():
            ws.cell(row=current_row, column=1, value=circuito).alignment = align_center
            for day_num, status in data_row.items():
                cell = ws.cell(row=current_row, column=day_num + 1, value=status)
                cell.alignment = align_center; cell.border = border_thin
            data_range = f"B{current_row}:{get_column_letter(last_col_calendar)}{current_row}"
            for i, status_code in enumerate(["UP", "PQ", "PP", "SD"]):
                ws.cell(row=current_row, column=summary_start_col + i, value=f'=COUNTIF({data_range}, "{status_code}")').font = font_bold_black
            current_row += 1
    data_range_format = f"B8:{get_column_letter(last_col_calendar)}{current_row + 5}"
    for status, fill_color in [("UP", fill_up), ("PQ", fill_pq), ("PP", fill_pp), ("SD", fill_sd)]:
        ws.conditional_formatting.add(data_range_format, CellIsRule(operator='equal', formula=[f'"{status}"'], fill=fill_color))
    ws.freeze_panes = 'B8'
    
    nome_arquivo_saida = f"Excel_OEE_{ano}_{mes:02d}.xlsx"
    caminho_completo_saida = os.path.join(arquivo_saida_folder, nome_arquivo_saida)
    try:
        wb.save(caminho_completo_saida)
        return caminho_completo_saida
    except Exception:
        return None