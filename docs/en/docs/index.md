# Introdução

## Sobre o projeto

Sapylot é uma biblioteca em Python desenhada para facilitar a criação de scripts de automação para o SAP GUI.

A biblioteca se baseia na documentação original da [SAP Scripting API](https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/babdf65f4d0a4bd8b40f5ff132cb12fa.html) para criar uma espécie de *ponte* para facilitar a criação dos scripts em Python, fornecendo tipagem e *auto complete* para as IDEs.

### Um pouco de contexto

A API de scripts do SAP é baseada em Microsoft COM (Component Object Model) e sua documentação nativa, indica exemplos implementados em VBA. Pensando em um cenário atual, Python é uma linguagem extremamente em alta e excelente para a criação de scripts porém, não possui nenhuma forma de interagir com o SAP utilizando COM e mantendo o *auto complete* ao mesmo tempo.

Pensando nisso, o projeto SAPylot foi criado. Criado utilizando a biblioteca pywin32 (win32com.client), o nosso objetivo é criar representações dos componentes de interface do SAP em Python e fornecer tipos, métodos e atributos para manipulação destes elementos de maneira simplificada.

## Um exemplo real

Abaixo um script que utiliza chamadas COM para interagir com o SAP

OBS: Todas as funções chamadas apartir dos objetos COM, não possuem nenhum tipo de *auto complete* ou *type hinting*

```{.py3 title='main.py'}
import win32com.client

def connect_sap():
    try:
        # Conectar ao SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)  # Assume que já há uma conexão ativa
        session = connection.Children(0)  # Assume que há uma sessão aberta

        return session
    except Exception as e:
        print(f"Erro ao conectar ao SAP: {e}")
        return None

def open_va03(session, order_number):
    try:
        # Acessa a transação VA03
        session.StartTransaction("VA03")
        
        # Insere o número da ordem de venda
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = order_number
        
        # Pressiona Enter para carregar a ordem
        session.findById("wnd[0]").sendVKey(0)

        print(f"Ordem {order_number} carregada com sucesso!")
    
    except Exception as e:
        print(f"Erro ao acessar VA03: {e}")

if __name__ == "__main__":
    session = connect_sap()
    
    if session:
        open_va03(session, "1234567890")
```

Agora, um exemplo do mesmo código implementando a lib sapylot

```{.py3 title='main.py'}
import sapylot


def script():
    session = sapylot.connect('S/4')
    sapylot.startTransaction(session, "VA03")


if __name__ == "__main__":
    script()

```