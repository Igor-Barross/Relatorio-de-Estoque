
Esse é meu primeiro projeto simples, apenas uma pequena automação que fiz para uma tarefa repetiva no meu trabalho, onde eu tinha que sempre escrever uma nova planilha manualmente sempre que quisesse fazer essa comparação. Decidi utilizar python para tornar isso mais automático.

# 📊 Projeto de Comparação de Estoque

## 📌 Descrição

Este projeto tem como objetivo automatizar a conferência de estoque entre:

- 📥 Estoque do sistema (extraído de um relatório primário)

- 📦 Estoque físico (digitado manualmente pelo usuário na planilha gerada)

**A planilha final traz:**

- Coluna para inserir estoque físico

- Cálculo automático da diferença entre físico e sistema

- Cálculo do custo da diferença

- Estilização da planilha (largura de colunas, alinhamento, bordas, formatação de números)

## ⚙️ Tecnologias Utilizadas

Python 3.x

pandas
 → manipulação de dados

openpyxl
 → escrita, fórmulas e formatação no Excel


## 🖼️ Exemplo de Saída

| Produto              | Estoque Físico | Estoque Sistema | Diferença | Custo Rep | Total    |
| -------------------- | -------------- | --------------- | --------- | --------- | -------- |
| Aquamix PR 300 25 KG | 6              | 12              | -6        | R\$ 83,72 | -R\$ 502 |
| Avetop Crescimento   | 200            | 640             | -440      | R\$ 2,23  | -R\$ 981 |
