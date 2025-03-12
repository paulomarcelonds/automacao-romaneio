Gerador de PDFs a partir de Planilhas Excel

Este script lê um arquivo Excel contendo registros de movimentações de materiais, processa os dados e gera arquivos PDF organizados por tipo de movimento e técnico responsável.

📌 Funcionalidades

Lê um arquivo Excel com dados de movimentação de materiais.

Formata e agrupa os dados por movimento e técnico.

Gera um PDF para cada combinação de "movimento" e "técnico".

Adiciona mensagens específicas para "ENTREGA" e "DEVOLUÇÃO".

🛠 Tecnologias Utilizadas

Python

Pandas (para manipulação de dados)

FPDF (para geração de PDFs)

📂 Estrutura dos Dados

O script processa as seguintes colunas do Excel:

sku: Código do item.

qtd: Quantidade.

descricao: Descrição do item.

movimento: Tipo de movimento (ENTREGA ou DEVOLUÇÃO).

tecnico: Nome do técnico responsável.

data: Data da movimentação.

📋 Como Usar

Instale as bibliotecas necessárias:

pip install pandas fpdf

Coloque o arquivo marcelo-entrega-devolucao.xlsx na mesma pasta do script.

Execute o script Python:

python script.py

Os PDFs gerados serão salvos no mesmo diretório do script.

📜 Exemplo de Saída

Os PDFs serão nomeados no formato:

movimento-tecnico-ddmmaaaa.pdf

Exemplo:

entrega-marcelo-12032024.pdf

🔍 Observações

Se a quantidade (qtd) estiver vazia, será substituída por um espaço em branco.

Datas são formatadas para DD-MM-YYYY na exibição e DDMMAAAA no nome do arquivo.

Mensagens específicas são adicionadas no final do PDF para ENTREGA e DEVOLUÇÃO.

🚀 Agora é só rodar o script e gerar seus PDFs automaticamente!

