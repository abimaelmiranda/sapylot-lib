# O Básico

O objetivo desta seção é apresentar as principais funções da biblioteca e seu uso

---

## get_sapgui_session()

A função ``get_sapgui_session()`` buscar e retonar a ultima sessão ativa do SAP GUI

```{.py3}
session = get_sapgui_session()
```

Funcionamento

- Verifica se existe uma ou mais sessões do SAP GUI ativas
- Retorna como ``session`` a última destas sessões
- Caso não exista sessão, retorna ``None``

---

## login() 

A função ``login()`` tenta realizar o login em um ambiente SAP caso este não esteja aberto e retorna uma ``session``. 

```{.py3}
session = login("s4", "200", "en", "user", "12345")
```
A sintaxe da função é

- ambient -> Parâmetro que representa o ambiente a ser aberto no SAP GUI
- cliente -> Parâmetro que representa o client/mandante da aplicação
- lang -> Linguagem na qual o ambiente deve ser executado
- username -> Usuário do ambiente informado
- password -> Senha do ambiente informado

Funcionamento

- Executa o sapshcut.exe através da linha de comando, usando os parametros informados
- Chama a função ``get_sapgui_session()`` para salvar a ultima sessão aberta
- Retorna a sessão salva

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