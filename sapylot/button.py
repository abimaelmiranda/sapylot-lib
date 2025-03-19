from .element import Element

from .exceptions import TypeMismatchException

class Button(Element):
    """Classe raiz para representação de Botões"""
    def __init__(self, session, elementId: str):
        super().__init__(session, elementId)
        self._check_type()

    def _check_type(self):
        if not self.rawElement.type in 'GuiButton':
            raise TypeMismatchException(self.id, self.type, Button)
        return

    def press(self):
        """Simula a ação de pressionar o botão"""
        return self.rawElement.press()