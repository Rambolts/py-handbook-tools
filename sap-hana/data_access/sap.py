import asyncio
from win32com.client import CDispatch
from app.configurations import get_configuration_by_path as get_config

from data_source.sap_gui.session import SAPGuiDataSource
from data_source.sap_gui.object import ObjectDataSource as objsap

class SapDataAccess():
    """ Representa a abstração do processamento no SAP de um item """
    
    def __init__(self):
        """ Contrutor da classe """
        self._sapgui = SAPGuiDataSource(**(get_config('data_source/sapgui')))
         
    def open(self) -> CDispatch:
        """ 
        Abrir o aplicativo sapgui.
        
        :returns : objeto de conexao com a sessao aberta do SAP.
        """
        return self._sapgui.new_session_connection()
        
    def login(self, connection: CDispatch):
        """ 
        Executa o login de acordo com os parametros passados no construtor do objeto.
        
        :param connection: objeto que representa uma conexão ao sap
        """
        self._sapgui.login(connection)
    
    def close(self):
        """ Matar qualquer instancia do SAPgui aberta """
        self._sapgui.kill()   
    
    async def _func_async_sap(self, f, timeout):
        """
        Funcao auxiliar, para executar uma outra funcao (f), com um delimitador de timeout.
        
        :param f: funcao
        :param timeout: tempo limite para a funcao
        """
        await asyncio.wait_for(f, timeout=timeout) 
    

class F04():
    """ Classe que representa a automacao para o procedimento para a atividade R064 """
    
    def __init__(self, connection: CDispatch):
        """ 
        Construtor da classe 
        
        :param conection: Objeto que representa uma conexão ao sap
        """
        self._conn = connection

    def select_transaction(self):
        """ Executa a selecao de uma transacao """
        c = self._conn
        objsap(c, 'wnd[0]/tbar[0]/okcd').write('F-04')
        objsap(c,'wnd[0]').key_press(0)

    def reclassificacao(self, entry):
        """ Processa a reclassificação de um lançamento """
        data = self.format_date(entry)
        self.select_transaction()
        self.selecionar_partida(
            data     = data,
            tipo_doc = entry['Tipo_Documento'], 
            empresa  = entry['Empresa'], 
            moeda    = entry['Moeda_Interna'], 
            periodo  = entry['Mes']
        )
        
    def format_date(entry: dict):
        """ Captura as datas do lançamento e transforma de acordo com o SAP GUI """
        return str(entry['Dia'])+'.'+str(entry['Mes'])+'.'+str(entry['Ano'])    

    def dados_cabecalho(self, data:str, tipo_doc:str, empresa:str, moeda:str, periodo:int):
        c = self._conn
        objsap(c, 'wnd[0]/usr/ctxtBKPF-BLDAT').write(data)
        objsap(c, 'wnd[0]/usr/ctxtBKPF-BUDAT').write(data)
        objsap(c, 'wnd[0]/usr/ctxtBKPF-BLART').write(tipo_doc)
        objsap(c, 'wnd[0]/usr/ctxtBKPF-BUKRS').write(empresa)
        objsap(c, 'wnd[0]/usr/ctxtBKPF-WAERS').write(moeda)
        objsap(c, 'wnd[0]/usr/txtBKPF-MONAT').write(periodo)
        objsap(c, 'wnd[0]/usr/sub:SAPMF05A:0122/radRF05A-XPOS1[3,0]').select()
        objsap(c, 'wnd[0]/tbar[1]/btn[6]').btn_press()

    def selecionar_partida(self, conta:str, tipo_conta:str, empresa:str, cod_razao:str):
        c = self._conn
        objsap(c, 'wnd[0]/usr/ctxtRF05A-AGKON').write(conta) # DEPENDE DO TIPO CONTA
        objsap(c, 'wnd[0]/usr/ctxtRF05A-AGKOA').write(tipo_conta)
        objsap(c, 'wnd[0]/usr/ctxtRF05A-AGBUK').write(empresa)
        objsap(c, 'wnd[0]/usr/ctxtRF05A-AGUMS').write(cod_razao)
        objsap(c, 'wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[2,0]').select()
        objsap(c, 'wnd[0]/tbar[1]/btn[16]').btn_press()

    def dar_baixa_a_diferenca(self, num_documento:str):
        c = self._conn
        objsap(c, 'wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]').write(num_documento)
        objsap(c, 'wnd[0]/tbar[1]/btn[16]').btn_press()
        objsap(c, 'wnd[0]/tbar[1]/btn[7]').btn_press()

    def inserir_item_cliente(self, chave_lancamento:str, cliente:str, cod_razao:str, montante:float, centro_lucro:str):
        c = self._conn
        objsap(c, 'wnd[0]/usr/ctxtRF05A-NEWBS').write(chave_lancamento)
        objsap(c, 'wnd[0]/usr/ctxtRF05A-NEWKO').write(cliente)
        objsap(c, 'wnd[0]/usr/ctxtRF05A-NEWUM').write(cod_razao)
        objsap(c, 'wnd[0]').key_press(0)

        objsap(c, 'wnd[0]/usr/txtBSEG-WRBTR').write(montante)
        objsap(c, 'wnd[0]/usr/txtBSEG-ZUONR').write('F' + centro_lucro[1:])
        objsap(c, 'wnd[0]/tbar[1]/btn[14]').btn_press()

    def corrigir_item_cliente(self, ):
        c = self._conn
        objsap(c, 'wnd[0]/mbar/menu[0]/menu[3]').key_toolbar_press()

        pass

    def capture_result(self) -> str:
        """ 
        Captura o resultado da execucao.
        
        :returns: string com os detalhes do processamento.
        """ 
        return objsap(self._conn, 'wnd[0]/sbar').get_text()