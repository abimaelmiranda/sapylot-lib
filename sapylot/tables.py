from typing import List
from pywintypes import com_error
import win32com.client
from .element import Element

from .exceptions import TypeMismatchException, InvalidIndexException

class Table(Element):
    """
    Classe raiz para representação de tabelas

    Attributes:
        row_count: Quantidade de linhas na tabela
        visible_row_count: Quantidade de linhas visíveis na tabela
        column_count: Quantidade de colunas na tabela
    """
    def __init__(self, session, elementId: str):
        super().__init__(session, elementId)
        self._check_type()
        self.row_count:int = self.rawElement.rowCount
        self.visible_row_count:int = self.rawElement.visibleRowCount
        self.column_count:int = self._get_column_count()

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
            row: Index da linha na tabela
            column: Index da coluna na tabela
        """
        return self.rawElement.getCell(row, column).text

    def set_cell_value(self, row:int, column:int, value:str) -> str:
        """
        Tenta configurar o valor atual de uma celula

        Args:
            row: Index da linha na tabela
            column: Index da coluna na tabela
            value: Novo valor a ser configurado
        """
        return self.rawElement.setCellValue(row, column, value)
    
#########################################################

class TableControl(Table):
    """
    Classe para representação de tabelas GuiTableControl
    
    Attributes:
        row_count: Quantidade de linhas na tabela
        visible_row_count: Quantidade de linhas visíveis na tabela
        column_count: Quantidade de colunas na tabela
    """
    def __init__(self, session, elementId: str):
        super().__init__(session, elementId)
        self._check_type()


    def find_in(self, value: str, column: int) -> int:
        """ 
        Tenta buscar um valor específico em uma coluna e retorna o índice da linha 
        onde o valor foi encontrado. Interrompe o loop quando o valor é encontrado.

        Args:
            value: Valor a ser buscado
            column: Index da coluna
        """
        self.session.findById("wnd[0]").maximize
        for i in range(self.row_count):
            try:
                T = self.session.findById(self.id)
                init_pos = T.verticalScrollbar.position

                if T.verticalScrollbar.position >= T.verticalScrollbar.maximum - 1:
                    break

                visible_column_values = self._get_visible_column_values(T, column)
                
                if value in visible_column_values:
                    return visible_column_values.index(value) + (i * self.visible_row_count)

                self._increase_scroll_pos(T, self.visible_row_count)

                T = self.session.findById(self.id)
                new_pos = T.verticalScrollbar.position
                
                if new_pos == init_pos:
                    break

            except com_error: 
                raise InvalidIndexException(column)
        return -1

    def select_row(self, index: int):
        """
        Seleciona a linha de uma tabela
        
        Args:
            index: Indice da linha a ser selecionada
        """
        self.rawElement.getAbsoluteRow(index).selected = True
        return
    

    def get_cell_value(self, row:int, column:int, initialScrollPos:int=0) -> str:
        """
        Coleta o valor atual de uma celula
        
        OBS: o valor máximo da linha é o valor maximo da propriedade visible_row_count

        Args:
            row: Index da linha na tabela
            column: Index da coluna na tabela
            initialScrollPos: Posição inicial do scroll vertical na tabela
        """
        try:
            if initialScrollPos != 0:
                self.rawElement.verticalScrollbar.position = initialScrollPos
                T = self.session.findById(self.id)
                return T.GetCell(row, column).Text

            return self.rawElement.getCell(row, column).text
        except com_error:
            raise InvalidIndexException(column)

    def get_column_values(self, column: int) -> List[str]:
        """ 
        Coleta todos os valores de uma determinada coluna e retorna em um array.

        Args:
            column: Index da coluna
        """
        data: List[str] = []
        last_fetched = set()

        self.session.findById("wnd[0]").maximize

        while True:
            T = self.session.findById(self.id)
            init_pos = T.verticalScrollbar.position

            if T.verticalScrollbar.position >= T.verticalScrollbar.maximum - 1:
                break
            
            temp_data = self._get_visible_column_values(T, column)
            
            for value in temp_data:
                if value not in last_fetched:
                    data.append(value)
                    last_fetched.add(value)

            self._increase_scroll_pos(T, self.visible_row_count)

            T = self.session.findById(self.id)
            new_pos = T.verticalScrollbar.position
            
            if new_pos == init_pos:
                break

        return data

    def _increase_scroll_pos(self, T: win32com.client.CDispatch, increment:int):
        """
        Incrementa a posição do scroll vertical da tabela

        Args:
            T: Objeto COM que representa uma tabela
            increment: Valor do incremento
        """
        pos = T.verticalScrollbar.position + increment
        T.verticalScrollbar.position = pos
        return pos
    
    def _check_type(self):
        TYPE = 'GuiTableControl'
        if not self.rawElement.type in TYPE:
            raise TypeMismatchException(self.id, self.type, TableControl)
        return   

    def _get_visible_column_values(self, T: win32com.client.CDispatch, column:int) -> List[str]:
        """
        Coleta os valores de todas colunas visiveis
        
        Args:
            T: Objeto COM que representa uma tabela
            column: Index da coluna
        """
        data:str = []
        for i in range(self.visible_row_count):
            data.append(T.GetCell(i, column).Text)
        return data 
    
#######################################

class TableShell(Table):
    """
    Classe para representação de tabelas GuiShell
    
    Attributes:
        row_count: Quantidade de linhas na tabela
        visible_row_count: Quantidade de linhas visíveis na tabela
        column_count: Quantidade de colunas na tabela
    """
    def __init__(self, session, elementId):
        super().__init__(session, elementId)
        self._check_type()
        self.columns = self._get_column_names()

    def get_column_values(self, column:str) -> List[str]:
        """
        Coleta os valores de uma coluna

        Args:
            column: Id da coluna na tabela
        """
        values = []
        for i in range(self.rawElement.rowCount):
            values.append(self.get_cell_value(i, column))
        return values

    def get_cell_value(self, row:int, column:str):
        """
        Coleta o valor atual de uma celula

        Args: 
            row: Index da linha na tabela
            column: Id da coluna na tabela
        """
        return self.rawElement.getCellValue(row, column)
    
    def sort_column(self, column:str, sort_type:int=0):
        """
        Ordena uma coluna da tabela

        Args:
            column: Id da coluna na tabela
            sort_type: Ordem de ordenação (0 = Ascendente, 1 = Descendente)
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