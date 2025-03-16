from re import match
from typing import List
from pywintypes import com_error
from .Table import Table
from ...exceptions import TypeMismatchException, InvalidIndexException

class TableShell(Table):
    """Classe para representação de tabelas GuiShell"""
    def __init__(self, session, elementId):
        super().__init__(session, elementId)
        self._check_type()
        self.columns = self._get_column_names()


    def _check_type(self):
        TYPE = 'GuiShell'
        if not self.rawElement.type == TYPE:
            raise TypeMismatchException(self.id, self.type, TableShell)

    def _get_column_names(self) -> List[str]:
        """Retorna uma lista com os nomes das colunas da tabela"""
        column_names = []
        for i in range(self.rawElement.columnCount):
            column_names.append(self.rawElement.columnOrder(i))
        return column_names

    def get_column_values(self, column:str) -> List[str]:
        """
        Coleta os valores de uma coluna

        Args:
            column(str): Id da coluna na tabela
        """
        values = []
        for i in range(self.rawElement.rowCount):
            values.append(self.get_cell_value(i, column))
        return values

    def get_cell_value(self, row:int, column:str):
        """
        Coleta o valor atual de uma celula

        Args: 
            row(Int): Index da linha na tabela
            column(str): Id da coluna na tabela
        """
        return self.rawElement.getCellValue(row, column)
    
    def sort_column(self, column:str, sort_type=0):
        """
        Ordena uma coluna da tabela

        Args:
            column(str): Id da coluna na tabela
            sort_type(int): Ordem de ordenação (0 = Ascendente, 1 = Descendente)
        """
        sort_order = None
        match(sort_type):
            case 0:
                sort_order = "&SORT_ASC"
            case 1:
                sort_order = "&SORT_DSC"

        self.rawElement.setCurrentCell(-1, column)
        self.rawElement.selectColumn(column)
        self.rawElement.contextMenu()
        self.rawElement.selectContextMenuItem(sort_order)
        return
    
    def get_column_name(self, column:int):
        """Retorna o ID da coluna com base no index"""
        column_name = self.columns[column]
        return column_name