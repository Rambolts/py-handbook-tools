from data_source.sap_gui.object import ObjectDataSource as objsap
import win32com.client
import os, subprocess, time

class SAPGuiDataSource():
    """ Contrutor da fonte de execucao da aplicacao SAP Autogui """
    
    def __init__(self, path: str, environment: str, username: str, password: str):
        """ 
        Construtor da classe
        :params path: caminho do executavel Sapgui local
        """
        self._path = path
        self._name_exc = self._path.split('\\')[-1]
        self._env  = environment
        self._user = username
        self._pass = password
    
    def _validate_path(self) -> bool:
        """ 
        Executa a validacao do caminho executavel informado
        :returns: boleano com o resultado da verificacao
        """        
        return True if os.path.isfile(self._path) else True
        
    def _open_check_subprocess(self) -> bool:
        """ 
        Executa a validacao de pelo menos 1 sessão aberta do SAPGui no windows local
        :returns: boleano com o resultado da verificacao
        """
        tlcall = 'TASKLIST', '/FI', 'imagename eq %s' % self._name_exc
        tlproc = subprocess.Popen(tlcall, shell=True, stdout=subprocess.PIPE)
        msg = str(tlproc.communicate()[0])
        return True if self._name_exc in msg else False
        
    def _terminate(self):
        """ Força terminar todos as sessoes do SAPGui em execucao no momento """
        subprocess.call(["taskkill", "/IM", self._name_exc, "/T", "/F"])

    def new_session_connection(self) -> win32com.client.CDispatch:
        """ 
        Cria uma conexão com a sessao do sap aberta 
        
        :returns : objeto de conexao com a sessao aberta do SAP.
        """
        if self._open_check_subprocess():
            self._terminate()
        
        if self._validate_path():
            subprocess.Popen(self._path)
            time.sleep(5)
        
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection(self._env, True)
        session = connection.Children(0)
        return session
     
    def login(self, conection: win32com.client.CDispatch):
        """ 
        Executa login de um usuario 
        
        :param conection: Objeto que representa uma conexão ao sap
        """
        objsap(conection, 'wnd[0]/usr/txtRSYST-BNAME').write(self._user)
        objsap(conection, 'wnd[0]/usr/pwdRSYST-BCODE').write(self._pass)
        objsap(conection, 'wnd[0]').key_press(0)
    
    def kill(self):
        """ Executa a finalizacao de todas sessoes abertas """
        self._terminate() 