from typing import List
from sapylot.exceptions import InvalidIndexException, TypeMismatchException
from pywintypes import com_error
from .element import Element

class TabStrip(Element):
    """
    Classe para representação de objetos GuiTableStrip
    """
    def __init__(self, session, elementId):
        super().__init__(session, elementId)
        self._check_type()
        self.tabs = self._get_tabs()

    def select_tab(self, tab:int):
        """
        Seleciona uma tabela na UI 

        Args:
            tab: indice da tabela a ser selecionada. Pode ser verificado através do atributo tabs
        """
        try:
            id_build = rf"{self.id}/tabpT\{self.tabs[tab]}"
            self.session.findById(id_build).select()
        except com_error:
            raise InvalidIndexException(tab)


    def _check_type(self):
        TYPE = 'GuiTabStrip'
        if not self.rawElement.type == TYPE:
            raise TypeMismatchException(self.id, self.type, TabStrip)

    def _get_tabs(self) -> List[str]:
        """Retorna uma lista com as abas da tabStrip"""
        tabs = []
        i = 0
        try:
            while True:
                self.rawElement.Children(i)
                tabs.append(f"0{i+1}")
                i+=1
        except Exception:
            pass
        return tabs