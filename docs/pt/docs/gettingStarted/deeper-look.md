# Um olhar aprofundado

O objetivo desta seção é lhe apresentar um olhar mais aprofudado sobre o funcionamento da biblioteca.

---

## Como o SAPylot funciona

Agora que você já possui o conhecimento sobre como funcionam as automações do SAP GUI, vamos entender como o SAPylot funciona e como ele pode tornar seu desenvolvimento mais produtivo

Dentro do SAP GUI, todo componente de interface possui um tipo. O SAPylot busca traduzir estes tipos para dentro do python utilizando o paradigma de classes. Com isso visamos ajudar o desenvolvedor fornecendo *auto complete*, *type hinting* e melhor tratamento de erros a partir de exceções personalizadas

## Classe ``Element``

A classe ``Element`` é a base da nossa biblioteca. Ela representa os elementos **GuiComponent Object** e **GuiVComponent Object** do SAP GUI.

Todas as demais classes da biblioteca herdam de ``Element``.

```{.py3 }
class Element:
    def __init__(self, session, elementId: str):
        self.session = session
        self.id = elementId
        self.name = None
        self.parent = None
        self.type = None
        self.rawElement = None

    def _find(self, session):
        """
        Busca o elemento e retorna o próprio objeto Element, permitindo encadeamento
        """
        try:
            self.rawElement = self.session.findById(r"{self.id}")
        except Exception:
            raise ElementNotFoundException(self.id)
        
        return self   
```
