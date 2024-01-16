import pandas as pd

def gerarXlsx(arquivo_csv,arquivo_xlsx):
    dados_csv = pd.read_csv(arquivo_csv)

    # Converter e salvar como XLSX
    dados_csv.to_excel(arquivo_xlsx, index=False)

