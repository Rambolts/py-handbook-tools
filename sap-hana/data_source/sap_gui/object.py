from win32com.client import CDispatch

class ObjectDataSource():
    """ Representa a abstração das funções de um objeto refernciado SAP """    

    def __init__(self, connection: CDispatch, id_object: str):
        """ 
        Construtor da classe
        
        :param conection: objeto que representa uma conexão ao sap
        :param id_object: string que representa o identificador de um objeto mapeavel SAP   
        """
        self._id_object = id_object
        self._obj = connection.findById(id_object)
               
    def write(self, text: str):
        """ 
        Executa a escrita em um objeto SAP
        
        :param text: string a ser escrita
        """
        self._obj.text = text
        
    def get_text(self) -> str:
        """ 
        Executa a captura do texto de um objeto SAP
        
        :returns: string capturado
        """
        return self._obj.text
     
    def btn_press(self):
        """
        Executa o click em um objeto do tipo button, sem referencia especifica
        """
        self._obj.press()
        
    def key_press(self, id_key: int):
        """ 
        Executa um click em um botao da janela sap informada por meio do identificador. Exemplo Enter(0)
        
        :param id_key: inteiro com a identificador do botao
        """
        self._obj.sendVKey(id_key)
     
    def key_toolbar_press(self, id_key: str):
        """ 
        Executa um click em um botao mapeado na barra de ferramentas do objeto mapeado por meio de um referencia
        
        param id_key: referencia do botao
        """
        self._obj.pressToolbarButton(id_key)
    
    def set_focus(self, caret_position: int):
        """
        Seleciona campo como foco
        
        :param caret_position: posição na tela
        """
        self._obj.setFocus()
        self._obj.caretPosition = caret_position

    def select(self):
        """
        Seleciona opções do tipo Radius, Checkbox ou do menu superior
        """
        self._obj.select()

    def select_column(self, id_column: str):
        """
        Executa a selecao de uma coluna de tabela
        
        :param id_row: inteiro identificador da coluna
        """
        self._obj.selectColumn(id_column)
     
    def select_row(self, id_row: int):
        """ 
        Executa a selecao de uma linha
        
        :param id_row: inteiro identificador da linha 
        """
        self._obj.selectedRows = id_row
        
    def cell_row_set(self, id_row: int, column_ref: str, value: str):
        """ 
        Executa a modificação do valor de uma celula de uma linha de uma tabela
        
        param id_row: inteiro identificador da linha
        param column_ref: string com o identificador da coluna da tabela
        param value: valor a ser inserido na modificacao
        """
        self._obj.modifyCell(id_row, column_ref, value) 
    
    def trigger_modified(self):
        """
        Executa a modificação no cache do SAPGui
        """
        self._obj.triggerModified()
