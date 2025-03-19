
# Criando nosso primeiro script

---

## Como o SAP GUI identifica os compentes internos

Existem diversa maneiras de se indentificar e selecionar um componente de interface no SAP GUI como: names, ids, parents, etc. Sendo a maneira mais comum de se utilizar o ID, iremos entender como ele funciona.

```{.py3 }
"wnd[0]/usr/ctxtVBAK-VBELN"
```

/// admonition | Dica
    type: tip

Sempre que for passar um ID como parâmetro para uma função, use raw strings!
///

Este é o ID de um campo de texto na interface. Onde o "wnd" representa a janela, o "usr" representa a área de usuário e o "ctxtVBAK-VBELN" representa o nome do campo de texto.

## Criando nosso script

Sabendo utilizar a ferramenta de gravação de scripts e como funciona um ID, vamos por em prática o exemplo proposto no inicio da seção.

/// admonition | Aviso
    type: warning

Para este exemplo, iremos assumir que o usuário já fez o login no ambiente.
///

Lembrando que o .VBS é apenas uma base. Em nosso exemplo usaremos python. Por este motivo primeiro devemos instalar a biblioteca ``pywin32``

```{.bash }
pip install pywin32
```

Isso nos permitirá criar uma conexão com o SAP GUI e executar comandos.

Após instalar o ``pywin32``, importamos ``win32com.client`` e criamos uma conexão ``COM`` junto ao SAP GUI.

```{.py3}
import win32com.client

def connect_sap():
    # Cria conexão ao SAP GUI e retorna a sessão ativa.
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        return session
    except Exception as e:
        print(f"Erro ao conectar ao SAP: {e}")
        return None
```

Agora que ja temos nossa ``session``, podemos utilizar o código gerado pela [Ferramenta de gravação de scripts](./sap-recording-tool.md) para seguir com a automação. 

```{.py3}
def script():
    session = connect_sap() # Obtém a sessão

    session.findById(r"wnd[0]/tbar[0]/okcd").text = "va03" # pesquisa "VA03" na barra de busca

    session.findById(r"wnd[0]").sendVKey(0) # Pressiona a tecla enter

    session.findById(r"wnd[0]/usr/ctxtVBAK-VBELN").text = "400923" # Pesquisa pela ordem de venda

    session.findById(r"wnd[0]/usr/btnBT_SUCH").press() # Pressiona o botão de pesquisa

    session.findById(r"wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press() # Abre o cabeçalho da ordem de venda

    so_type = session.findById(r"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-AUART").text 
    # Coleta o texto do campo de "tipo de ordem"

    return so_type

```

Este foi apenas um exemplo simples de como uma automação funciona e como a ferramenta de gravação de scripts do SAP pode ser útil. Nas próximas sessões entraremos a fundo em como nossa biblioteca funciona e como ela pode lhe auxliar ao desenvolver estas automações.