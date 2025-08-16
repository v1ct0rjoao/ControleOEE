import os

UPLOAD_FOLDER = 'dados_brutos'

OUTPUT_FOLDER = 'relatorios'

PROCESSED_CSV_FILENAME = 'dados_processados.csv'

PROCESSED_CSV_PATH = os.path.join(OUTPUT_FOLDER, PROCESSED_CSV_FILENAME)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)