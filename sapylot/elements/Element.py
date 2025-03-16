import win32com.client

from ..exceptions import ElementNotFoundException


class Element:
    """Classe raiz para todas as outras. Baseada no GuiComponent Object e em GuiVComponent Object"""
    def __init__(self, session: win32com.client.CDispatch, elementId: str):
        self.session = session
        self.id = elementId
        self.name = None
        self.parentId = None
        self.type = None
        self.rawElement = None
        self._find()

    def _find(self):
        """
        Busca o elemento e retorna o próprio objeto Element, permitindo encadeamento
        """
        try:
            self.rawElement = self.session.findById(rf"{self.id}")
            self.name = self.rawElement.name
            self.parentId = self.rawElement.parent.id
            self.type = self.rawElement.type
        except Exception as e:
            raise ElementNotFoundException(self.id) from e
        return self
    
    def set_focus(self):
        """Configura o foco para o elemento atual"""
        self.rawElement.setFocus()
        return

    def get_text(self):
        """
        Retorna o texto do elemento encontrado
        """
        if self.rawElement:
            return self.rawElement.text.strip()
        return None

    def set_text(self, text: str):
        """
        Insere um texto no elemento
        """
        if self.rawElement:
            self.rawElement.text = text
        return self