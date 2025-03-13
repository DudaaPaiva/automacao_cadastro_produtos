# Automação de Cadastro de Produtos

Este projeto foi desenvolvido para automatizar o preenchimento de campos em uma plataforma de cadastro de produtos a partir de uma planilha do Excel. O script utiliza bibliotecas como `openpyxl`, `pyperclip`, `pyautogui` e `time` para manipular os dados da planilha e interagir com a interface gráfica da plataforma, preenchendo os campos de forma eficiente e precisa.

## Requisitos

Antes de executar o código, certifique-se de ter as seguintes bibliotecas instaladas:

- `openpyxl`: Para manipulação da planilha Excel.
- `pyperclip`: Para copiar e colar os dados mantendo os acentos.
- `pyautogui`: Para automação da interface gráfica.
- `time`: Para adicionar intervalos entre as ações de clique e colagem.

Você pode instalar essas bibliotecas executando:

pip install openpyxl pyperclip pyautogui

## Funcionalidade

O script lê os dados de um arquivo Excel chamado `produtos_ficticios.xlsx` e preenche automaticamente os campos de cadastro de produtos na plataforma, incluindo informações como:

- Nome do produto
- Descrição
- Categoria
- Código do produto
- Peso
- Dimensões
- Preço
- Quantidade
- Validade
- Cor
- Tamanho
- Material
- Fabricante
- País de origem
- Observações
- Código de barras
- Localização no armazém

### Fluxo do Código

1. **Carregar a Planilha**: O código começa carregando a planilha `produtos_ficticios.xlsx` e seleciona a aba de produtos.
2. **Iteração nas Linhas**: Para cada linha da planilha (iniciando da segunda linha), o script copia os valores dos campos e os cola nas respectivas áreas na plataforma.
3. **Automação de Cliques e Colagens**: Utiliza `pyautogui` para clicar nas posições específicas da tela e `pyperclip` para colar as informações copiadas.
4. **Botões de Navegação**: Após preencher os campos de um produto, o script clica no botão "Próximo" para avançar para o próximo cadastro.

## Como Usar

1. **Prepare a Planilha**: Certifique-se de que o arquivo `produtos_ficticios.xlsx` está no formato correto, com os dados organizados conforme o esperado.
2. **Execute o Script**: Execute o script em seu ambiente Python. Ele preencherá os campos da plataforma automaticamente.

### Exemplo de como rodar o script:

python automatizar_cadastro.py 

## Observações

- O script depende das coordenadas exatas dos campos na plataforma. As coordenadas de clique (por exemplo, `pyautogui.click(846,236)`) podem precisar ser ajustadas para se adequar ao layout da sua tela.
- Certifique-se de que a plataforma de cadastro esteja aberta e visível na tela ao rodar o script, para garantir que os cliques e as colagens ocorram nos locais corretos.

## Contribuições

Contribuições são bem-vindas! Se você tiver sugestões de melhorias ou correções, fique à vontade para enviar um pull request ou abrir uma issue.


