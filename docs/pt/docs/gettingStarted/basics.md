O objetivo desta seção é apresentar as principais funções da biblioteca e seu uso

---

## get_element()
A função ``get_element()`` é responsável por buscar um elemento na interface e retorná-lo. 
```{.py3}
element = get_element(session, "wnd[0]/usr/txtockd")

```
A sintaxe da função é

- session -> Parâmetro que representa a conexão em aberto com o SAP GUI
- elementType -> Parâmetro que representa o tipo do elemento, sendo o valor padrão, a classe Element
- elementId -> String que representa o ID do elemento a ser buscado

Funcionamento

- A função busca o elemento na interface
- Verifica se o tipo interno do elemento é compatível com o tipo sapylot passado na chamada da função
- Caso a validação de tipos ocorra sem erros, o elemento é retornado com o tipo correto inferido.

---

## start_transaction()
A função ``start_transaction()`` inicia uma transação na janela atual do SAP GUI 
```{.py3}
start_transaction(session, "VA03")
```
A sintaxe da função é

- session -> Parâmetro que representa a conexão em aberto com o SAP GUI
- transaction -> String que representa o código da transação

Funcionamento

- A função utiliza a sessão atual e tenta iniciar uma transação, utilizando o prefixo "/n" do SAP.