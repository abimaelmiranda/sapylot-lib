class ElementNotFoundException(Exception):
    def __init__(self, element_id, message="Elemento não encontrado"):
        self.element_id = element_id
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        return f"{self.message}: {self.element_id}"
    

class TypeMismatchException(Exception):
    def __init__(self, element_id, invalid_type, correct_type):
        self.element_id = element_id
        self.message = f"O tipo {invalid_type} é incompativel com o tipo {correct_type}"
        super().__init__(self.message)

    def __str__(self):
        return f"{self.message}"
    
    
class InvalidIndexException(Exception):
    def __init__(self, index):
        self.message = f"O indice {index} é inválido para esta operação"
        super().__init__(self.message)

    def __str__(self):
        return f"{self.message}"