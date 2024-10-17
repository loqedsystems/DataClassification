import pandas as pd
from database_config import get_database_engine
from utilities import convert_seconds_to_hhmmss, classify_activity, get_day_of_year  # Importar a função de utilidade
import re
import time
from sqlalchemy.exc import OperationalError
from dotenv import load_dotenv
import os
from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError

def execute_query_with_retry(engine, query, retries=3, delay=5):
    """
    Executa a consulta SQL com lógica de retry em caso de falha de transação.
    """
    attempt = 0
    while attempt < retries:
        try:
            with engine.connect() as connection:
                data = pd.read_sql(query, connection)
            return data
        except OperationalError as e:
            print(f"Erro de transação, tentativa {attempt + 1} de {retries}: {e}")
            attempt += 1
            time.sleep(delay)  # Aguarda antes de tentar novamente
    raise RuntimeError("Falha ao executar a consulta após várias tentativas")

def extract_computer_name(name):
    """
    Extrai a palavra do nome do computador que começa com 'df', 
    """
    match = re.search(r'df[-\w]*', name.lower())
    if match:
        return match.group(0)
    return None

def process_data(data):
    """
    Processa o DataFrame: converte datas, tempos e classifica atividades.
    """
    # Separar a coluna "Date" em "Data" e "Hora"
    data['Date'] = pd.to_datetime(data['Date'])  # Converte para o tipo datetime
    data['Data Apenas'] = data['Date'].dt.date  # Extrai apenas a data
    data['Hora Apenas'] = data['Date'].dt.time  # Extrai apenas o horário

    # Converter "ActivityTime" de segundos para HH:MM:SS
    data['ActivityTime'] = data['ActivityTime'].apply(lambda x: convert_seconds_to_hhmmss(int(x)))

    # Aplicar a função de classificação
    classified_data = data.apply(classify_activity, axis=1)

    # Verificar se a coluna 'Nome do Computador' está no dataset e aplicar a extração
    if 'MachineName' in data.columns:
        classified_data['Nome Extraído'] = data['MachineName'].apply(extract_computer_name)
    else:
        print("Coluna 'MachineName' não encontrada no dataset.")

    return classified_data

def clean_illegal_characters(text):
    # Remove caracteres não permitidos no Excel (como "‼")
    if isinstance(text, str):
        return re.sub(r'[^\x20-\x7E]', '', text)  # Remove caracteres fora do intervalo ASCII imprimível
    return text

def save_to_excel(data, filename='dados_classificados.xlsx', max_rows_per_sheet=1000000):
    """
    Salva o DataFrame em um arquivo Excel (.xlsx), criando novas planilhas se o limite de 1 milhão de linhas for excedido.
    """
    # Verifica o número de linhas no DataFrame
    num_rows = data.shape[0]

    # Limpa caracteres ilegais antes de salvar
    data = data.applymap(clean_illegal_characters)
    
    if num_rows <= max_rows_per_sheet:
        # Se o número de linhas for menor ou igual ao limite, salva normalmente
        try:
            data.to_excel(filename, index=False, engine='openpyxl')
            print(f"Exportação para Excel concluída com sucesso! Arquivo salvo como {filename}")
        except IllegalCharacterError as e:
            print(f"Erro ao salvar o arquivo Excel devido a caracteres ilegais: {e}")
    else:
        # Criar um novo arquivo Excel e salvar em múltiplas planilhas
        writer = pd.ExcelWriter(filename, engine='openpyxl')
        for i in range(0, num_rows, max_rows_per_sheet):
            # Divida os dados em partes de 'max_rows_per_sheet'
            chunk = data.iloc[i:i + max_rows_per_sheet]
            # Limpa caracteres ilegais no chunk
            chunk = chunk.applymap(clean_illegal_characters)
            # Salve cada parte em uma nova planilha
            sheet_name = f'Sheet_{i // max_rows_per_sheet + 1}'
            try:
                chunk.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Exportando {len(chunk)} registros para {sheet_name}")
            except IllegalCharacterError as e:
                print(f"Erro ao salvar o chunk {sheet_name} devido a caracteres ilegais: {e}")
        
        writer.close()  # Corrigido para 'close'
        print(f"Exportação para Excel em múltiplas planilhas concluída! Arquivo salvo como {filename}")

def save_to_csv(data, filename='dados_classificados.csv'):
    """
    Salva o DataFrame processado em um arquivo CSV (.csv).
    """
    data.to_csv(filename, index=False)
    print(f"Exportação para CSV concluída com sucesso! Arquivo salvo como {filename}")

load_dotenv()

def build_query(start_date=None, end_date=None):
    """
    Constrói a consulta SQL com base no intervalo de datas fornecido pelo usuário.
    Se datas não forem fornecidas, todos os dados serão retornados.
    """
    # Converter as datas para o dia do ano (PartitionID)
    start_partition = get_day_of_year(start_date) if start_date else None
    end_partition = get_day_of_year(end_date) if end_date else None

    # Iniciar a query base
    query = """
    WITH CTE_ProcessDetails AS (
        SELECT 
            e.DictionaryId,
            f.ProcessName,
            e.MD5_Checksum
        FROM [dbo].[utWin_ProcessDic] e
        LEFT JOIN [dbo].[utWin_ProcessNameDic] f ON f.[DictionaryId] = e.[ProcessNameId]
    ),
    CTE_ComputerDetails AS (
        SELECT 
            cbu.ComputerId as Id,
            cbu.OrganizationId as OrgId,
            o.[Name] AS Host,
            c.MachineName AS ClientMachineName,
            c.TcpAddress AS ClientTcpAddress,
            s.MachineName AS ServerMachineName,
            s.TcpAddress AS ServerTcpAddress
        FROM 
            dbo.[utWin_ComputerByUnit] cbu
        LEFT JOIN 
            dbo.[ut_Organization] o ON cbu.[OrganizationId] = o.[Id]
        LEFT JOIN 
            utWinClient_TCPAddress c ON cbu.ComputerId = c.ID
        LEFT JOIN 
            utWinserver_TCPAddress s ON cbu.ComputerId = s.ID
        WHERE 
            cbu.[Type] = 1 
            AND ComputerID IN (
                SELECT DISTINCT [ComputerId] 
                FROM [dbo].[utWin_ComputerByUnit] 
                WHERE [OrganizationId] IN (77, 78, 98, 143, 185, 186, 190) 
                AND Type = 1
            )
    ),
    CTE_UserWebAppActivity AS (
        SELECT 
            u.ComputerId,
            u.UserId,
            u.ProcessId,
            u.[Domain],
            u.[URL],
            u.[WindowTitle],
            u.[UTCActualSliceId],
            u.[ActivityTime],
            u.[IsPublicIp],
            u.[IsConnect],
            u.[PartitionID],
            p.UserName,
            p.DomainName
        FROM utWinClient_UserWebActivity u
        LEFT JOIN [dbo].[utWin_UserNameDic] p ON u.UserId = p.DictionaryId
        WHERE u.UTCActualSliceId BETWEEN 28752660 AND 28801619
    """

    # Adicionar filtros de PartitionID se as datas forem fornecidas
    if start_partition and end_partition:
        query += f" AND (u.[PartitionID] >= {start_partition} AND u.[PartitionID] <= {end_partition})"

    # Continuar a consulta com UNION
    query += """
        UNION ALL

        SELECT  
            u.ComputerId,
            u.UserId,
            u.ProcessId,
            SPACE(0) AS [Domain],
            SPACE(0) AS [URL],
            u.[WindowTitle],
            u.[UTCActualSliceId],
            u.ActiveTime AS [ActivityTime],
            0 AS [IsPublicIp],
            u.[IsConnect],
            u.[PartitionID],
            p.UserName,
            p.DomainName
        FROM utWinClient_UserAppActivity u
        LEFT JOIN [dbo].[utWin_UserNameDic] p ON u.UserId = p.DictionaryId
        WHERE u.UTCActualSliceId BETWEEN 28752660 AND 28801619
    """

    # Adicionar novamente filtros de PartitionID se as datas forem fornecidas
    if start_partition and end_partition:
        query += f" AND (u.[PartitionID] >= {start_partition} AND u.[PartitionID] <= {end_partition})"

    # Fechar a query com a parte final da seleção e junções
    query += """
    ),
    CTE_ServerWebAppActivity AS (
        SELECT  
            u.ComputerId,
            u.UserId,
            u.ProcessId,
            u.[Domain],
            u.[URL],
            u.[WindowTitle],
            u.[UTCActualSliceId],
            u.[ActivityTime],
            u.[IsPublicIp],
            u.[IsConnect],
            u.[PartitionID],
            p.UserName,
            p.DomainName
        FROM utWinServer_UserWebActivity u
        LEFT JOIN [dbo].[utWin_UserNameDic] p ON u.UserId = p.DictionaryId
        WHERE u.UTCActualSliceId BETWEEN 28752660 AND 28801619

        UNION ALL

        SELECT  
            u.ComputerId,
            u.UserId,
            u.ProcessId,
            SPACE(0) AS [Domain],
            SPACE(0) AS [URL],
            u.[WindowTitle],
            u.[UTCActualSliceId],
            u.ActiveTime AS [ActivityTime],
            0 AS [IsPublicIp],
            u.[IsConnect],
            u.[PartitionID],
            p.UserName,
            p.DomainName
        FROM utWinServer_UserAppActivity u
        LEFT JOIN [dbo].[utWin_UserNameDic] p ON u.UserId = p.DictionaryId
        WHERE u.UTCActualSliceId BETWEEN 28752660 AND 28801619
    )

    SELECT
        data.ComputerId,
        cd.[Host] AS OrganizationId,
        COALESCE(cd.ClientMachineName, cd.ServerMachineName) AS [MachineName],
        COALESCE(cd.ClientTcpAddress, cd.ServerTcpAddress) AS [IpAddress],
        cd.[Host] AS HostName,
        (CASE WHEN LEN(p.UserName) = 0 THEN SPACE(0) ELSE p.[DomainName] + CHAR(92) + p.UserName END) AS [UserName],
        ([dbo].[uf_GetDateByUTCSliceId]([UTCActualSliceId], 1, 0)) AS [Date],
        pd.[ProcessName],
        [Domain],
        [URL] AS [URL_Name],
        [WindowTitle],
        [UTCActualSliceId],
        [ActivityTime],
        [IsConnect]
    FROM (
        SELECT * FROM CTE_UserWebAppActivity
        UNION ALL
        SELECT * FROM CTE_ServerWebAppActivity
    ) AS data
    JOIN CTE_ComputerDetails cd ON data.ComputerId = cd.Id
    JOIN CTE_ProcessDetails pd ON data.ProcessId = pd.DictionaryId
    LEFT JOIN [dbo].[utWin_UserNameDic] p ON p.DictionaryId = data.UserId;
    """

    return query

def filter_by_user(data, usernames):
    """
    Filtra o DataFrame para incluir apenas as linhas onde o UserName corresponde aos usernames fornecidos.
    Considera o formato 'NOME_DA_MAQUINA\\user_name' e faz o filtro somente pela parte do 'user_name'.
    Aceita múltiplos usernames separados por vírgula.
    """

    # Limpar e separar os usernames por vírgula, se múltiplos forem fornecidos
    if isinstance(usernames, str) and usernames.strip():  # Verifica se a string não está vazia
        usernames = [user.strip().lower() for user in usernames.split(",")]
    else:
        usernames = []  # Se estiver vazio ou apenas espaços, definir como lista vazia

    # Separar o 'UserName' no DataFrame após a barra invertida (caso exista) e converter para minúsculas
    data['UserName'] = data['UserName'].apply(lambda x: x.split("\\")[-1].strip().lower() if "\\" in x else x.strip().lower())

    # Filtrar pelo nome de usuário na lista
    if usernames:
        filtered_data = data[data['UserName'].isin(usernames)]
    else:
        filtered_data = data  # Se não houver username, retornar todos os dados

    return filtered_data




if __name__ == "__main__":
    # Solicitar nome(s) do(s) usuário(s) e intervalo de datas via input
    usernames = input("Digite o(s) nome(s) do(s) usuário(s) (separados por vírgula, ou deixe vazio para todos): ")
    start_date = input("Digite a data inicial (YYYY-MM-DD, ou deixe vazio para todas): ")
    end_date = input("Digite a data final (YYYY-MM-DD, ou deixe vazio para todas): ")

    # Configurações de conexão para SQL Server
    load_dotenv()
    db_host = os.getenv('DB_HOST')
    db_name = os.getenv('DB_NAME')
    db_user = os.getenv('DB_USER')
    db_password = os.getenv('DB_PASSWORD')

    # Obter engine de conexão
    engine = get_database_engine(db_host, db_name, db_user, db_password)

    # Construir e executar a consulta com base no intervalo de datas
    query = build_query(start_date, end_date)
    data = execute_query_with_retry(engine, query)

    # Aplicar o filtro de username(s) no Python após a consulta SQL
    filtered_data = filter_by_user(data, usernames)

    # Processar os dados
    processed_data = process_data(filtered_data)

    # Salvar os dados em arquivos Excel e CSV
    save_to_excel(processed_data, filename='resultado_dados_classificados.xlsx')
    save_to_csv(processed_data, filename='resultado_dados_classificados.csv')

