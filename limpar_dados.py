import pandas as pd
import re
import os

def interpretar_data_flexivel(string_data):
    """
    Interpreta apenas datas no formato brasileiro (dd/mm/yyyy hh:mm[:ss]).
    Ignora datas com AM/PM ou formatos diferentes.
    """
    if pd.isna(string_data) or not isinstance(string_data, str):
        return pd.NaT

    string_data = string_data.strip()

    # Rejeita strings com AM ou PM
    if re.search(r'\b(am|pm)\b', string_data, re.IGNORECASE):
        return pd.NaT

    try:
        return pd.to_datetime(string_data, format="%d/%m/%Y %H:%M:%S", errors='coerce')
    except ValueError:
        try:
            return pd.to_datetime(string_data, format="%d/%m/%Y %H:%M", errors='coerce')
        except ValueError:
            return pd.NaT

def limpar_dados_brutos(lista_arquivos, arquivo_saida):
    """
    Lê uma lista de arquivos de texto brutos, extrai e limpa os dados no formato brasileiro.
    """
    all_lines = []
    print("--- INICIANDO SCRIPT DE LIMPEZA DE DADOS (VERSÃO BRASILEIRA SOMENTE) ---")

    for nome_arquivo in lista_arquivos:
        try:
            try:
                with open(nome_arquivo, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            except UnicodeDecodeError:
                with open(nome_arquivo, 'r', encoding='latin-1') as f:
                    lines = f.readlines()
            all_lines.extend(lines)
        except FileNotFoundError:
            print(f"  - AVISO: Arquivo '{nome_arquivo}' não encontrado. Pulando.")
            
    if not all_lines:
        print("Nenhum dado para processar.")
        return

    conteudo_total = "\n".join(all_lines)

    print("\nExtraindo dados...")

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
            data_list.append({
                'circuito': circuito,
                'datastart': datastart_str,
                'datastop': datastop_str
            })

    if not data_list:
        print("Nenhum dado válido extraído. Verifique o formato dos arquivos.")
        return

    print(f"Encontrados e processados {len(data_list)} registros.")
    df = pd.DataFrame(data_list)

    print("\nConvertendo datas no formato brasileiro...")
    df['datastart'] = df['datastart'].apply(interpretar_data_flexivel)
    df['datastop'] = df['datastop'].apply(interpretar_data_flexivel)
    df.dropna(subset=['datastart'], inplace=True)

    df['circuito_num'] = df['circuito'].str.extract(r'(\d+)').astype(int)
    df.sort_values(by=['circuito_num', 'datastart'], inplace=True)
    df.drop(columns=['circuito_num'], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Salvar em formato brasileiro
    df.to_csv(arquivo_saida, index=False, sep=';', date_format='%d/%m/%Y %H:%M:%S')
    print(f"\n✅ Arquivo salvo com sucesso em: '{arquivo_saida}'")


if __name__ == "__main__":
    arquivos_de_entrada = ['dig01.txt', 'dig02.txt', 'dig03.txt', 'dig04.txt'] 
    arquivo_de_saida_processado = "dados_processados.csv"
    limpar_dados_brutos(arquivos_de_entrada, arquivo_de_saida_processado)
