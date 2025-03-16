"""Script de testes"""
# pylint: disable=line-too-long
import subprocess
import tkinter as tk
from pywintypes import com_error
import win32com
from win32com.client import GetObject
import time

class LoginWindow():
    """Classe para tela de login"""
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Login Form")
        self.username = None
        self.password = None
        self.__login_window()

    def __center_window(self, window):
        window.update()
        scr_width, scr_height = window.winfo_screenwidth(), window.winfo_screenheight()
        border_width = window.winfo_rootx() - window.winfo_x()
        title_height = window.winfo_rooty() - window.winfo_y()
        win_width = window.winfo_width() + border_width + border_width
        win_height = window.winfo_height() + title_height + border_width
        x = (scr_width - win_width) // 2
        y = (scr_height - win_height) // 2
        window.geometry(f"+{x}+{y}")

    def __login_window(self):
        def enter(event=None) -> None:
            self.__login()

        self.window.username_label = tk.Label(self.window, text="Username:")
        self.window.username_label.pack()
        self.window.username_entry = tk.Entry(self.window)
        self.window.username_entry.pack()
        self.window.password_label = tk.Label(self.window, text="Password:")
        self.window.password_label.pack()
        self.window.password_entry = tk.Entry(self.window, show="*")
        self.window.password_entry.pack()
        self.window.login_button = tk.Button(self.window, text="Login", command=self.__login)
        self.window.login_button.pack()
        self.window.attributes("-topmost", True)
        self.window.geometry("200x120")
        self.window.bind('<Return>', enter)
        self.__center_window(self.window)
        self.window.focus()
        self.window.username_entry.focus()
        self.window.mainloop()

    def __login(self):
        self.username = self.window.username_entry.get()
        self.password = self.window.password_entry.get()
        self.window.quit()

class SAPAuto():
    @classmethod
    def get_sapgui(cls):
        """Função para obter o SAPgui"""
        try:
            sap_gui = GetObject("SAPGUI").GetScriptingEngine
            conn = sap_gui.Connections(0)
            sessions = conn.Sessions.Count - 1
            session = conn.Sessions(sessions)
            return session
        except com_error:
            return None

    @classmethod
    def login_sap(cls):
        window = LoginWindow()
        username = window.username
        password = window.password

        app_path = 'C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe'
        ambiente = 'PS4'
        mandante = '400'
        idioma = 'PT'

        if username is None or username == '' or password is None or password == '':
            print("Nenhum username ou password foi informado.")
            return

        processo = subprocess.Popen([app_path, f'-system={ambiente}', f'-client={mandante}', f'-user={username}', f'-pw={password}', f'-language={idioma}'])
        processo.communicate()

        time.sleep(3)
        session = cls.get_sapgui()

        return session

    @classmethod
    def create_transaction_window(cls, transaction: str, session: win32com.client.CDispatch) -> win32com.client.CDispatch:
        session.findById("wnd[0]/tbar[0]/okcd").text = f"/o{transaction}"
        session.findById("wnd[0]").sendVKey(0)
        new_session = cls.get_sapgui()

        return new_session

def va23(session) -> None:
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "20026219"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

def va25(session) -> None:
    session.findById("wnd[0]/usr/ctxtSVKORG-LOW").text = "2001"
    session.findById("wnd[0]/usr/ctxtSVKORG-HIGH").text = "2002"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 0
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()

def main():
    """Função principal"""
    session = SAPAuto.get_sapgui() #Tenta pegar o controle do SAP aberto
    if session is None:
        session = SAPAuto.login_sap()

    sessao_va25 = SAPAuto.create_transaction_window("VA25", session) #Cria sessão para controle da transação
    print(sessao_va25.findById("wnd[0]/sbar").text)
    if "Sem autorização" in sessao_va25.findById("wnd[0]/sbar").text:
        print("Sem permissão para a transação")
    else:
        sessao_va23 = SAPAuto.create_transaction_window("VA23", session) #Cria sessão para controle da transação
        va25(session=sessao_va25) #Executa transação em janela própria
        va23(session=sessao_va23) #Executa transação em janela própria

        session.StartTransaction(Transaction="VA05") #Executa transação sobre a janela principal que estiver aberta

if __name__ == "__main__":
    main()
