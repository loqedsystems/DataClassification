from sqlalchemy import create_engine
from sqlalchemy.exc import OperationalError

def get_database_engine(db_host, db_name, db_user, db_password):
    """
    Cria e retorna a engine de conexão com o banco de dados SQL Server.
    Verifica se a conexão foi bem-sucedida.
    """
    connection_string = f"mssql+pyodbc://{db_user}:{db_password}@{db_host}/{db_name}?driver=ODBC+Driver+17+for+SQL+Server"
    engine = create_engine(connection_string)
    
    try:
        # Testa a conexão
        with engine.connect() as connection:
            print("Conexão com o banco de dados bem-sucedida!")
    except OperationalError as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
    
    return engine

