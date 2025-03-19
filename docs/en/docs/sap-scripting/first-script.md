# Criando nosso primeiro script

---

## Como o SAP GUI identifica os compentes internos

Existem diversa maneiras de se indentificar e selecionar um componente de interface no SAP GUI como: names, ids, parents, etc. Sendo a maneira mais comum de se utilizar o ID, iremos entender como ele funciona.

```{.py3 }
"wnd[0]/usr/ctxtVBAK-VBELN"
```

/// details | Dica
    open: False
    type: tip

Sempre que for passar um ID como parâmetro para uma função, use raw strings!
///

Este é o ID de um campo de texto na interface. Onde o "wnd" representa a janela, o "usr" representa a área de usuário e o "ctxtVBAK-VBELN" representa o nome do campo de texto.

## Criando nosso script

Sabendo utilizar a ferramenta de gravação de scripts e como funciona um ID, vamos por em prática o exemplo proposto no inicio da seção.

Lembrando que o .VBS é apenas uma base. Em nosso exemplo usaremos python. Por este motivo primeiro devemos instalar a biblioteca ``pywin32``

```{.bash }
pip install pywin32
```