<<<<<<< HEAD
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
=======
import pandas as pd
import re
import os

def interpretar_data_flexivel(string_data):
    """
    Interpreta uma string de data que pode estar no formato americano (com AM/PM)
    ou brasileiro (24h). Retorna um objeto datetime ou NaT (Not a Time) em caso de erro.
    """
    if pd.isna(string_data) or not isinstance(string_data, str):
        return pd.NaT
    string_data_lower = string_data.strip().lower()
    if not any(char.isdigit() for char in string_data_lower):
        return pd.NaT
    try:
        # A presença de AM/PM é o indicador mais forte do formato americano (mês primeiro)
        if 'am' in string_data_lower or 'pm' in string_data_lower:
            return pd.to_datetime(string_data, dayfirst=False)
        else:
            return pd.to_datetime(string_data, dayfirst=True)
    except (ValueError, TypeError):
        return pd.NaT

def limpar_dados_brutos(lista_arquivos, arquivo_saida):
    """
    Lê uma lista de arquivos de texto brutos, extrai e limpa os dados usando uma lógica
    robusta de duas etapas para lidar com formatos de dados inconsistentes e concatenados.
    """
    all_lines = []
    print("--- INICIANDO SCRIPT DE LIMPEZA DE DADOS (VERSÃO ROBUSTA) ---")
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

    # Junta todo o conteúdo em uma única string para facilitar a busca entre linhas
    conteudo_total = "\n".join(all_lines)

    print("\nExtraindo dados com lógica de extração rigorosa...")
    
    # --- NOVA LÓGICA DE EXTRAÇÃO ---
    # 1. Padrão para encontrar os marcadores de circuito (ex: "Circuit443")
    circuit_pattern = re.compile(r"Circuit\d+", re.IGNORECASE)
    
    # 2. Padrão para encontrar um padrão completo de data e hora
    datetime_pattern = re.compile(
        r"\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?", 
        re.IGNORECASE
    )

    data_list = []
    # Encontra todos os marcadores de circuito e suas posições no texto
    circuit_matches = list(circuit_pattern.finditer(conteudo_total))

    for i, match in enumerate(circuit_matches):
        circuito = match.group(0)
        start_pos = match.end() # Posição onde o nome do circuito termina

        # Define o fim do bloco de dados: ou é o início do próximo circuito, ou o fim do texto
        if i + 1 < len(circuit_matches):
            end_pos = circuit_matches[i+1].start()
        else:
            end_pos = len(conteudo_total)
            
        # Isola o bloco de texto pertencente a este circuito
        data_blob = conteudo_total[start_pos:end_pos]
        
        # 3. Procura todas as ocorrências de data/hora dentro do bloco isolado
        found_datetimes = datetime_pattern.findall(data_blob)
        
        datastart_str = None
        datastop_str = None
        
        if len(found_datetimes) >= 1:
            datastart_str = found_datetimes[0]
        if len(found_datetimes) >= 2:
            datastop_str = found_datetimes[1]
            
        if datastart_str: # Só adiciona à lista se encontrou pelo menos a data de início
             data_list.append({
                'circuito': circuito, 
                'datastart': datastart_str, 
                'datastop': datastop_str
            })
    
    if not data_list:
        print("Nenhum dado de circuito/data válido foi extraído. Verifique o formato dos arquivos.")
        return
        
    print(f"Encontrados e processados {len(data_list)} registros brutos.")
    df = pd.DataFrame(data_list)

    print("\nConvertendo e padronizando as datas...")
    df['datastart'] = df['datastart'].apply(interpretar_data_flexivel)
    df['datastop'] = df['datastop'].apply(interpretar_data_flexivel)
    df.dropna(subset=['datastart'], inplace=True)
    
    df['circuito_num'] = df['circuito'].str.extract(r'(\d+)').astype(int)
    df.sort_values(by=['circuito_num', 'datastart'], inplace=True)
    df.drop(columns=['circuito_num'], inplace=True)
    df.reset_index(drop=True, inplace=True)
    
    df.to_csv(arquivo_saida, index=False, sep=';', date_format='%Y-%m-%d %H:%M:%S')
    print(f"\n✅ Processamento de limpeza concluído! Dados salvos em '{arquivo_saida}'")


if __name__ == "__main__":
    arquivos_de_entrada = ['dig01.txt', 'dig02.txt', 'dig03.txt', 'dig04.txt'] 
    arquivo_de_saida_processado = "dados_processados.csv"
    limpar_dados_brutos(arquivos_de_entrada, arquivo_de_saida_processado)
>>>>>>> b168e54c4147bafe255aaec9747ff4a520d69bd9
