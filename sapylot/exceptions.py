class ElementNotFoundException(Exception):
    def __init__(self, element_id, message="Elemento não encontrado"):
        self.element_id = element_id
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        return f"{self.message}: {self.element_id}"