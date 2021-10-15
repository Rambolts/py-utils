import pyodbc
import pandas as pd
from functools import lru_cache
import re

@lru_cache(maxsize=1)
def get_sql_obdc_driver():
    """
    Função que retorna o driver ODBC mais atualizado para conexão com SQL

    :returns: driver ODBC mais atualizado disponível.
    """
    drivers = pyodbc.drivers()
    
    sql_driver_pattern = r'ODBC Driver \d+ for SQL Server'
    sql_drivers = sorted([driver for driver in drivers if re.match(sql_driver_pattern, driver)])
    
    assert len(sql_drivers)>0, 'Nenhum driver ODBC para conexão com SQL encontrado.'
    return sql_drivers[-1]


class SqlDbDataSource():
    """Classe de abstração da conexão com banco de dados SQL Server"""

    def __init__(self, server, database, username, password):
        """
        Instancia um objeto de conexão com banco de dados SQL Server.

        :param server: nome do servidor ou IP
        :param database: nome da base de dados
        :param username: usuário da base de dados
        :param password: senha da base de dados

        Observação: precisa do driver ODBC 13 para conexão com SQL Server instalado.
        """
        sql_driver = '{' + get_sql_obdc_driver() + '}'
        self._server = server
        self._database = database
        self._username = username

        self._connection_string = f'Driver={sql_driver}; Server={server}; Database={database}; uid={username}; pwd={password}; '
        self._connection = pyodbc.connect(self._connection_string)

    @lru_cache(maxsize=None)
    def read_sql(self, query, **pd_kwargs):
        """
        Executa query SQL

        :param query: comando SQL, normalmente do tipo SELECT
        :returns: Pandas DataFrame com o resultado da consulta
        """
        return pd.read_sql(query, self._connection, **pd_kwargs)

    def read_sql_without_cache(self, query, **pd_kwargs):
        """
        Executa query SQL sem uso de cache

        :param query: comando SQL, normalmente do tipo SELECT
        :returns: Pandas DataFrame com o resultado da consulta
        """
        return self.read_sql.__wrapped__(self, query, **pd_kwargs)

    def getdate(self):
        """
        Retorna a data e hora do servidor SQL

        :returns: datetime atual do servidor
        """
        cursor = self._connection.cursor()
        dt = cursor.execute('SELECT GETDATE()').fetchval()
        return dt

    def commit(self):
        """
        Conclui transação SQL
        """
        self._connection.commit()

    def rollback(self):
        """
        Restaura estado antes da transação SQL
        """
        self._connection.rollback()
        
    def connect(self):
        return self._connection