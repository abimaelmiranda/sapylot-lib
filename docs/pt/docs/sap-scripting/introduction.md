# Introdução a automações SAP GUI

Esta seção é dedicada a aqueles que não possuem conhecimento de como automatizar o SAP GUI

---

Se você utiliza a ferramenta SAP com certeza já se deparou com algum processo que precisa ser feito repetidamente, e que certamente poderia ser automatizado.
Pensando nisso, a própria SAP disponibiliza uma API nativa para criação de scripts/automações.

Para introduzir você aos conceitos principais do processo de criação de automações, vamos desenvolver uma automação simples que acessa a transação "VA03", pesquisa e acessa uma Ordem de venda (OV) e finaliza nos retornando o tipo desta ordem.
