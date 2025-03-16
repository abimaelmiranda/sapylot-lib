from typing import List
from pywintypes import com_error
from .Table import Table
from ...exceptions import TypeMismatchException, InvalidIndexException

class TableControl(Table):
    """Classe para representação de tabelas GuiTableControl"""
    def __init__(self, session, elementId: str):
        super().__init__(session, elementId)
        self._check_type()


    def _check_type(self):
        TYPE = 'GuiTableControl'
        if not self.rawElement.type in TYPE:
            raise TypeMismatchException(self.id, self.type, TableControl)
        return   


    def _get_visible_column_values(self, T, column:int) -> List[str]: 
        data:str = []
        for i in range(self.visible_row_count):
            data.append(T.GetCell(i, column).Text)
        return data 

    def get_cell_value(self, row:int, column:int, initialScrollPos=0) -> str:
        """
        Coleta o valor atual de uma celula
        
        OBS: o valor máximo da linha é o valor maximo da propriedade visible_row_count

        Args:
            row(Int): Index da linha na tabela
            column(Int): Index da coluna na tabela
            initialScrollPos(Int): Posição inicial do scroll vertical na tabela
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
        """
        data: List[str] = []
        last_fetched = set()
        
        # Loop de coleta de dados da tabela
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

    def _increase_scroll_pos(self, T, increment):
        """
        Incrementa a posição do scroll vertical da tabela
        """
        pos = T.verticalScrollbar.position + increment
        T.verticalScrollbar.position = pos
        return pos
    
    def find_in(self, value: str, column: int) -> int:
        """ 
        Tenta buscar um valor específico em uma coluna e retorna o índice da linha 
        onde o valor foi encontrado. Interrompe o loop quando o valor é encontrado.
        """
        
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