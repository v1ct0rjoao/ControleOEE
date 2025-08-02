# config.py
import os

# Pasta para onde os arquivos .txt enviados pelo usuário serão salvos.
UPLOAD_FOLDER = 'dados_brutos'

# Pasta onde os relatórios .xlsx finais serão armazenados.
OUTPUT_FOLDER = 'relatorios'

# Nome do arquivo CSV intermediário.
PROCESSED_CSV_FILENAME = 'dados_processados.csv'

# Caminho completo para o arquivo CSV processado.
PROCESSED_CSV_PATH = os.path.join(OUTPUT_FOLDER, PROCESSED_CSV_FILENAME)


# Garante que os diretórios existam.
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)