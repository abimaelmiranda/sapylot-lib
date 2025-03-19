class ElementNotFoundException(Exception):
    """
    Exceção utilizada quando um elemento não foi encontrado na UI
    
    """
    def __init__(self, element_id, message="Elemento não encontrado"):
        self.element_id = element_id
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        return f"{self.message}: {self.element_id}"
    

class TypeMismatchException(Exception):
    """
    Exceção utilizada quando o tipo interno do SAP diverge da classe sapylot informada para representação
    
    """
    def __init__(self, element_id, invalid_type, correct_type):
        self.element_id = element_id
        self.message = f"O tipo {invalid_type} é incompativel com o tipo {correct_type}"
        super().__init__(self.message)

    def __str__(self):
        return f"{self.message}"
    
    
class InvalidIndexException(Exception):
    """
    Exceção utilizada quando o usuário tenta acessar algum indice inválido
    
    """
    def __init__(self, index):
        self.message = f"O indice {index} é inválido para esta operação"
        super().__init__(self.message)

    def __str__(self):
        return f"{self.message}"