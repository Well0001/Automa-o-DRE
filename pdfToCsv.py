import camelot
import pandas as pd

def pdf_to_csv(pdf_path, csv_path):
    # Extrair tabelas do PDF
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')

    # Verificar se há tabelas a serem salvas em CSV
    if not tables:
        print("Nenhuma tabela encontrada no PDF.")
        return

    # Criar um único DataFrame concatenando todas as tabelas
    df = pd.concat(table.df for table in tables)

    # Salvar o DataFrame em um arquivo CSV
    df.to_csv(csv_path, index=False)
