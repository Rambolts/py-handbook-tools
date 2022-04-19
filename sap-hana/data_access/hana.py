from data_source.hana import HanaDataSource
from app.configurations import get_configuration_by_path
import pandas as pd

class HanaDataAccess():
    """ Representa o acesso de dados à base de dados Hana """

    def __init__(self, dso: HanaDataSource = None):
        """
        Construtor

        :param dso: fonte de dados SQL. Quando None, puxa a partir do arquivo de configuração
        """
        self._dso = dso or HanaDataSource(**get_configuration_by_path('data_source/hana'))

    def faça_sua_query(self) -> pd.DataFrame:
        """ Aqui você irá utilizar da maneira que melhor lhe servir. A princício utilizando alguma query de busca que te retorne algum pd.DataFrame """
        pass