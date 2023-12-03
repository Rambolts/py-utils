import pyodbc
import pandas as pd
from functools import lru_cache
from datetime import datetime
import re
from typing import Any

import warnings
warnings.filterwarnings('ignore', category=UserWarning)

@lru_cache(maxsize=1)
def get_sql_obdc_driver() -> str:
    """
    Função que retorna o driver ODBC mais atualizado para conexão com SQL

    :returns: driver ODBC mais atualizado disponível.
    """
    drivers = pyodbc.drivers()
    
    sql_driver_pattern = r'ODBC Driver \d+ for SQL Server'
    sql_drivers = sorted([driver for driver in drivers if re.match(sql_driver_pattern, driver)])
    
    assert len(sql_drivers)>0, 'Nenhum driver ODBC para conexão com SQL encontrado.'
    return sql_drivers[-1]

class SqlServerPlugin():
    """Classe de abstração da conexão com banco de dados SQL Server"""

    def __init__(self, server: str, database: str, username: str, password: str, trusted: bool = False):
        """
        Instancia um objeto de conexão com banco de dados SQL Server.

        :param server: nome do servidor ou IP
        :param database: nome da base de dados
        :param username: usuário da base de dados
        :param password: senha da base de dados

        Observação: precisa do driver ODBC para conexão com SQL Server instalado.
        """
        sql_driver = '{' + get_sql_obdc_driver() + '}'
        self._server = server
        self._database = database
        self._username = username
        self._trusted = trusted

        trusted = 'yes' if trusted else 'no'
        self._connection_string = f'Driver={sql_driver}; Server={server}; Database={database}; uid={username}; pwd={password}; trusted={trusted};'
        self._connection = pyodbc.connect(self._connection_string)

    @property
    def connection(self) -> pyodbc.Connection:
        return self._connection
    
    @property
    def cursor(self) -> pyodbc.Cursor:
        return self._connection.cursor()
    
    @property
    def connection_string(self) -> str:
        return self._connection_string

    def read(self, query: str, **pd_kwargs) -> pd.DataFrame:
        """
        Executa query SQL

        :param query: comando SQL, normalmente do tipo SELECT
        :returns: Pandas DataFrame com o resultado da consulta
        """
        return pd.read_sql(query, self._connection, **pd_kwargs)

    @lru_cache(maxsize=None)
    def cached_read(self, query: str, **pd_kwargs) -> pd.DataFrame:
        """
        Executa query SQL com uso de cache

        :param query: comando SQL, normalmente do tipo SELECT
        :returns: Pandas DataFrame com o resultado da consulta
        """
        return self.read(query, **pd_kwargs)
    
    def read_value(self, query: str, params: Any = None) -> Any:
        """
        Executa query SQL para retornar um único valor

        :param query: comando SQL
        :returns: valor retornado pela consulta
        """
        cursor = self._connection.cursor()
        if params:
            val = cursor.execute(query, params).fetchval()
        else:
            val = cursor.execute(query).fetchval()
        return val
    
    @lru_cache(maxsize=None)
    def cached_read_value(self, query: str, params: Any = None) -> Any:
        """
        Executa query SQL para retornar um único valor com uso de cache

        :param query: comando SQL
        :returns: valor retornado pela consulta
        """
        return self.read_value(query, params)

    def get_date(self) -> datetime:
        """
        Retorna a data e hora do servidor SQL

        :returns: datetime atual do servidor
        """
        cursor = self._connection.cursor()
        dt = cursor.execute('SELECT GETDATE()').fetchval()
        return dt

    def execute(self, sql: str, values: Any = None) -> int:
        """
        Executa comando SQL

        :param sql: comando sql a ser executado
        :param values: valores a serem passados como parâmetros
        :returns: quantidade de registros afetados
        """
        cursor = self._connection.cursor()
        if values:
            rowcount = cursor.execute(sql, values).rowcount
        else:
            rowcount = cursor.execute(sql).rowcount
        return rowcount
        
    def execute_many(self, sql: str, values: Any = None, fast: bool = False) -> int:
        """
        Executa comando múltiplo SQL

        :param sql: comando sql a ser executado
        :param values: valores a serem passados como parâmetros
        :param fast: indica se deve ser feito de forma rápida
        :returns: quantidade de valores informados
        """
        cursor = self._connection.cursor()
        old_fast_executemany = cursor.fast_executemany
        cursor.fast_executemany = fast
        
        cursor.executemany(sql, values)
        cursor.fast_executemany = old_fast_executemany
        
        rowcount = len(values)
        return rowcount

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