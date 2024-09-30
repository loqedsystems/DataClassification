
import pandas as pd
from database_config import get_database_engine
from utilities import convert_seconds_to_hhmmss, classify_activity
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

if __name__ == "__main__":
    # Configurações de conexão para SQL Server
    db_host = os.getenv('DB_HOST')
    db_name = os.getenv('DB_NAME')
    db_user = os.getenv('DB_USER')
    db_password = os.getenv('DB_PASSWORD')

    # Obter engine de conexão
    engine = get_database_engine(db_host, db_name, db_user, db_password)

    # Consulta SQL
    query = """
 WITH CTE_ComputerDetails AS (
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
)

SELECT
    data.ComputerId,
    cd.[Host] AS OrganizationId,
    COALESCE(cd.ClientMachineName, cd.ServerMachineName) AS [MachineName],
    COALESCE(cd.ClientTcpAddress, cd.ServerTcpAddress) AS [TcpAddress],
    cd.[Host] AS HostName,
    (CASE WHEN LEN(p.UserName) = 0 THEN '' ELSE p.[DomainName] + CHAR(92) + p.UserName END) AS [UserName],
    ([dbo].[uf_GetDateByUTCSliceId]([UTCActualSliceId], 1, 0)) AS [Date],
    d.[ProcessName],
    [Domain],
    [URL] AS [URL_Name],
    [WindowTitle],
    [UTCActualSliceId],
    [ActivityTime],
    [IsConnect]
FROM 
    (
        SELECT * 
        FROM (
            SELECT 
                ComputerId,
                UserId,
                ProcessId,
                [Domain],
                [URL],
                [WindowTitle],
                [UTCActualSliceId],
                [ActivityTime],
                [IsPublicIp],
                [IsConnect],   
                [PartitionID]
            FROM utWinClient_UserWebActivity
            WHERE UTCActualSliceId BETWEEN 28663380 AND 28771379  
            AND ([PartitionID] >= 183 AND [PartitionID] <= 257)
            
            UNION 
            
            SELECT  
                ComputerId,
                UserId,
                ProcessId,
                SPACE(0) AS [Domain],
                SPACE(0) AS [URL],
                [WindowTitle],
                [UTCActualSliceId],
                ActiveTime AS [ActivityTime],
                0 AS [IsPublicIp],
                [IsConnect],   
                [PartitionID]
            FROM utWinClient_UserAppActivity
            WHERE UTCActualSliceId BETWEEN 28663380 AND 28771379  
            AND ([PartitionID] >= 183 AND [PartitionID] <= 257)
        ) cli

        UNION ALL

        SELECT * 
        FROM (
            SELECT 
                ComputerId,
                UserId,
                ProcessId,
                [Domain],
                [URL],
                [WindowTitle],
                [UTCActualSliceId],
                [ActivityTime],
                [IsPublicIp],
                [IsConnect],   
                [PartitionID]
            FROM utWinServer_UserWebActivity
            WHERE UTCActualSliceId BETWEEN 28663380 AND 28771379  
            AND ([PartitionID] >= 183 AND [PartitionID] <= 257)
            
            UNION 
            
            SELECT  
                ComputerId,
                UserId,
                ProcessId,
                SPACE(0) AS [Domain],
                SPACE(0) AS [URL],
                [WindowTitle],
                [UTCActualSliceId],
                ActiveTime AS [ActivityTime],
                0 AS [IsPublicIp],
                [IsConnect],   
                [PartitionID]
            FROM utWinServer_UserAppActivity
            WHERE UTCActualSliceId BETWEEN 28663380 AND 28771379  
            AND ([PartitionID] >= 183 AND [PartitionID] <= 257)
        ) srv
    ) data 
JOIN CTE_ComputerDetails cd ON data.ComputerId = cd.Id
LEFT JOIN [dbo].[utWin_ProcessDic] c ON c.DictionaryId = data.ProcessId
LEFT JOIN [dbo].[utWin_ProcessNameDic] d ON d.DictionaryId = c.ProcessNameId
LEFT JOIN [dbo].[utWin_UserNameDic] p ON p.[DictionaryId] = data.UserId;







    """


    # Executar a consulta com retry
    data = execute_query_with_retry(engine, query)
    
    # Processar e salvar os dados
    processed_data = process_data(data)
    save_to_excel(processed_data)
    save_to_csv(processed_data)
