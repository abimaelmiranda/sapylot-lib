import subprocess
import time
from typing import Type, TypeVar
from pywintypes import com_error
import win32com.client
import psutil

from .element import Element


def get_sapgui_session() -> win32com.client.CDispatch | None:
        """
        Função para obter o SAPgui
        Esta função coleta a ultima sessão ativa do SAP, caso não exista, utilize a função login()
        """
        try:
            sap_gui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
            conn = sap_gui.Connections(0)
            sessions = conn.Sessions.Count - 1
            session = conn.Sessions(sessions)
            return session
        except com_error:
            return None


def login(ambient: str, client: str, lang: str, username:str, password: str) -> win32com.client.CDispatch | None:
    """
    Inicia o SAP GUI instalado na maquina a partir do caminho padrão 'C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe'

    Args:
        ambient: Ambiente a ser iniciado no SAP GUI
        client: Mandante do ambiente
        lang: Idioma do ambiente
        username: Usuario utilizado para login neste ambiente
        password: Senha utilizada para login neste ambiente
    
    Returns:
        Session: Retorna uma session usada para manipular a interface do SAP Gui.
    """
    APP_PATH = 'C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe'
    

    processo = subprocess.Popen([APP_PATH , f'-system={ambient}', f'-client={client}', f'-user={username}', f'-pw={password}', f'-language={lang}'])
    processo.communicate()

    time.sleep(5)
    session = get_sapgui_session() 
    return session

def start_transaction(session: win32com.client.CDispatch, transaction: str):
    """
    Inicia uma transação com a função /n

    Args:
        session: Sessão de referencia retornada pela função login ou get_sapgui_session
        transaction: Transação a ser iniciada
    """
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{transaction}"
    session.findById("wnd[0]").sendVKey(0)
    return

def start_transaction_new_window(session: win32com.client.CDispatch, transaction: str):
    """
    Inicia uma transação com a função /o

    Args:
        session: Sessão de referencia retornada pela função login ou get_sapgui_session
        transaction: Transação a ser iniciada
    """
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/o{transaction}"
    session.findById("wnd[0]").sendVKey(0)

    new_session = get_sapgui_session()
    return new_session


T = TypeVar("T", bound="Element")	        
def get_element(session: win32com.client.CDispatch, elementId: str,  elementType: Type[T] = Element) -> T:
    """
    Busca um elemento na transação atual e retorna um objeto do tipo especificado.

    Args:
        session: Sessão de referência retornada pela função login ou get_sap_session.
        elementType: Classe do tipo de elemento esperado (ex: GridView, TableControl).
        elementId: ID do elemento a ser buscado.

    Returns:
        T: Instância do tipo solicitado (por padrão, um Element).
    """
    el = elementType(session, elementId)
    return el


def press_enter(session: win32com.client.CDispatch, wndIndex:int=0):
    """
    Pressiona a tecla ENTER dentro de uma transação no SAP GUI

    Args:
        session: Sessão de referencia retornada pela função get_sapgui_session
        wndIndex: Indice da janela a ser pressionada a tecla ENTER
    """
    session.findById(f"wnd[{wndIndex}]").sendVKey(0)
    return


def execute_transaction(session: win32com.client.CDispatch):
    """
    Simula o pressionar da tecla F8 (executar) dentro das transações que possuem suporte

    Args:
        session: Sessão de referencia retornada pela função get_sapgui_session
    """
    session.findById("wnd[0]").sendVKey(8)
    return


def turn_back(session: win32com.client.CDispatch):
    """
    Simula o pressionar da tecla F3 (voltar) dentro das transações

    Args:
        session: Sessão de referencia retornada pela função get_sapgui_session
    """
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    return




    # def press(self):
    #     """
    #     Pressiona o elemento, se for um botão
    #     """
    #     if self.rawElement:
    #         self.rawElement.press()
    #     return self
    
    # def select(self):
    #     """
    #     Seleciona o elemento ou aba
    #     """
    #     if self.rawElement:
    #         self.rawElement.select()
    #     return self
    
    # def close(self):
    #     """
    #     Fecha o elemento ou aba
    #     """
    #     if self.rawElement:
    #         self.rawElement.close()
    #     return self


def close_sap_process():
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
