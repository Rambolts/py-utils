import pandas as pd
from functools import lru_cache
from datetime import datetime
from hdbcli import dbapi
from typing import Any

class SapHanaPlugin():
    """Classe de abstração da conexão com banco de dados Hana"""

    def __init__(self, host: str, port: int, user: str, password: str):
        """
        Instancia um objeto de conexão com banco de dados Hana.

        :param host: nome do servidor ou IP
        :param port: porta de acesso
        :param user: usuário da base de dados
        :param password: senha da base de dados
        """
        self._host = host
        self._port = port
        self._user = user
        self._connection = dbapi.connect(address=host, port=port, user=user, password=password)

    def __del__(self):
        """
        Destrutor da classe
        """
        try:
            self._connection.close()
        except:
            pass

    def read(self, query: str, params: Any = None) -> pd.DataFrame:
        """
        Executa query SQL

        :param query: comando SQL, normalmente do tipo SELECT
        :returns: Pandas DataFrame com o resultado da consulta
        """
        cursor = self._connection.cursor()
        cursor.execute(query, params)
        cols = [col_info[0] for col_info in cursor.description]
        df = pd.DataFrame(cursor.fetchall(), columns=cols)
        return df

    @lru_cache(maxsize=None)
    def cached_read(self, query: str, params: Any = None) -> pd.DataFrame:
        """
        Executa query SQL sem uso de cache

        :param query: comando SQL, normalmente do tipo SELECT
        :returns: Pandas DataFrame com o resultado da consulta
        """
        return self.read(query, params)

    def get_date(self) -> datetime:
        """
        Retorna a data e hora do servidor SQL

        :returns: datetime atual do servidor
        """
        cursor = self._connection.cursor()
        dt = cursor.execute('SELECT GETDATE()').fetchval()
        return dt
