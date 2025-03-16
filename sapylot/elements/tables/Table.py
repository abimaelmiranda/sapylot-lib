from typing import List
from pywintypes import com_error
from ..Element import Element

from ...exceptions import TypeMismatchException, InvalidIndexException

class Table(Element):
    """Classe raiz para representação de tabelas"""
    def __init__(self, session, elementId: str):
        super().__init__(session, elementId)
        self._check_type()
        self.row_count = self.rawElement.rowCount
        self.visible_row_count = self.rawElement.visibleRowCount
        self.column_count = self._get_column_count()

    def _check_type(self):
        SHELLTYPES = ['GuiShell', 'GuiTableControl', 'GuiGridView']
        if not self.rawElement.type in SHELLTYPES:
            raise TypeMismatchException(self.id, self.type, Table)
        return

    def _get_column_count(self):
        result = None
        try:
            result = self.rawElement.columnCount
        except Exception:
            pass
        try:
            result = self.rawElement.Columns.Count
        except Exception:
            pass
        return result
        

    def get_cell_value(self, row:int, column:int) -> str:
        """
        Coleta o valor atual de uma celula

        Args:
            row(Int): Index da linha na tabela
            column(Int): Index da coluna na tabela
        """
        return self.rawElement.getCell(row, column).text

    def set_cell_value(self, row:int, column:int, value) -> str:
        """
        Tenta configurar o valor atual de uma celula

        Args:
            row(Int): Index da linha na tabela
            column(Int): Index da coluna na tabela
            value(any): Novo valor a ser configurado
        """
        return self.rawElement.setCellValue(row, column, value)
    
    