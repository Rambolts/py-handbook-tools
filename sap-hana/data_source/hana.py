import pandas as pd
from hdbcli import dbapi

class HanaDataSource():
    """ Representa a abstração da conexão com uma banco de dados SQL Hana """    
    
    def __init__(self, host, port, user, pswd):
        """
        Instancia um objeto de conexão com o Hana.

        :param host: numero do host
        :param port: numero da porta
        :param username: usuário hana
        :param password: senha para acesso
        """       
        self._connection = dbapi.connect(address=host, port=port, user=user, password=pswd)

    def query(self, query: str, **pd_kwargs) -> pd.DataFrame:
        """ 
        Executa uma consulta e retorna um dataframe com os dados
        
        :param query: comando SQL, do tipo SELECT
        :returns: Pandas DataFrame com o resultado da consulta
        """
        return pd.read_sql(query, self._connection, **pd_kwargs)