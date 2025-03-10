import subprocess
import time
import win32com.client
import psutil

from .exceptions import ElementNotFoundException


def connect(ambiente: str):
    """
    Inicia o SAP GUI instalado na maquina a partir do caminho padrão C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe

    Args:
        ambiente(str): Ambiente a ser iniciado no SAP GUI
    
    Retorna:
        Session: Retorna uma session usada para manipular a interface do SAP Gui.
    """
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(5)
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    conn = application.OpenConnection(rf"{ambiente}", True)

    session = conn.Children(0)
    session.findById("wnd[0]").maximize
    return session


def closeSAPProcess():
    """
    Encerra todos os processos do SAP Gui em execução
    """
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if 'sapgui.exe' in proc.info['name'].lower() or 'saplogon.exe' in proc.info['name'].lower():
                print(f"Fechando o processo SAP: {proc.info['name']} (PID: {proc.info['pid']})")
                proc.terminate() 
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


def login(session, user: str, passw: str):
    """
    Realiza o login no ambiente selecionado

    Args:
        session(session): Sessão de referencia retornada pela função connect
        user(str): Usuário do ambiente SAP informado
        passw(str):  Usuário do ambiente SAP informado na função connect
    """
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = passw
    session.findById("wnd[0]").sendVKey(0)
    return


def startTransaction(session, transaction: str):
    """
    Inicia uma transação com a função /n

    Args:
        session(session): Sessão de referencia retornada pela função connect
        transaction(str): Transação a ser iniciada
    """
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{transaction}"
    session.findById("wnd[0]").sendVKey(0)
    return


class Element:
    def __init__(self, session, elementId: str):
        self.session = session
        self.elementId = elementId
        self.rawElement = None

    def find(self):
        """
        Busca o elemento e retorna o próprio objeto Element, permitindo encadeamento
        """
        try:
            self.rawElement = self.session.findById(rf"{self.elementId}")
        except Exception as e:
            raise ElementNotFoundException(self.elementId) from e
        return self

    def getText(self):
        """
        Retorna o texto do elemento encontrado
        """
        if self.rawElement:
            return self.rawElement.text.strip()
        return None

    def setText(self, text: str):
        """
        Insere um texto no elemento
        """
        if self.rawElement:
            self.rawElement.text = text
        return self

    def press(self):
        """
        Pressiona o elemento, se for um botão
        """
        if self.rawElement:
            self.rawElement.press()
        return self
    
    def select(self):
        """
        Seleciona o elemento ou aba
        """
        if self.rawElement:
            self.rawElement.select()
        return self
    
    def close(self):
        """
        Seleciona o elemento ou aba
        """
        if self.rawElement:
            self.rawElement.close()
        return self

    def setVScrollPosition(self, pos:int):
        """
        Configura uma nova posição para uma ScrollBar Vertical
        """
        if self.rawElement:
            self.rawElement.verticalScrollbar.position = pos
        return self

def getElement(session, elementId: str):
    """
    Busca um elemento na transação atual e retorna um objeto Element
    Args:
        session(session): Sessão de referencia retornada pela função connect
        elementId(str): ID do elemento a ser buscado
    Retorna:
        Element: Objeto que encapsula o elemento SAP
    """
    return Element(session, elementId).find()


def pressEnter(session):
    """
    Pressiona a tecla ENTER dentro de uma transação no SAP GUI
    Args:
        session(session): Sessão de referencia retornada pela função connect
    """
    session.findById("wnd[0]").sendVKey(0)
    return


def executeTransaction(session):
    """
    Simula o pressionar da tecla F8 (executar) dentro das transações que possuem suporte
    Args:
        session(session): Sessão de referencia retornada pela função connect
    """
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    return


def turnBack(session):
    """
    Simula o pressionar da tecla F3 (voltar) dentro das transações 
    Args:
        session(session): Sessão de referencia retornada pela função connect
    """
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    return
