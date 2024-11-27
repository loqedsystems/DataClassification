import pandas as pd

# Nome do arquivo CSV de entrada
arquivo_csv = "C:/Users/Keslon Magdiel/OneDrive - Loqed Systems Ltda/Documents/DataClassification/DataClassification/82MI.csv"

# Nome do arquivo CSV de saída
arquivo_saida = "saida_distinct.csv"

# Nome da coluna onde você quer aplicar o distinct
coluna = "nome_da_coluna"

# Carregar o arquivo CSV em um DataFrame
df = pd.read_csv(arquivo_csv)
print(df.columns.tolist())


# # Obter valores únicos da coluna especificada
# df_distinct = df[[coluna]].drop_duplicates()

# # Salvar o resultado em um novo arquivo CSV
# df_distinct.to_csv(arquivo_saida, index=False)

# print(f"Os valores distintos da coluna '{coluna}' foram salvos em '{arquivo_saida}'.")
