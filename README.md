# Processador de Arquivos

O Processador de Arquivos é uma ferramenta desenvolvida para ajudar na limpeza e organização de dados de produtos, categorizando-os corretamente e salvando os resultados em um arquivo Excel existente. Esta ferramenta foi criada para ser fácil de usar, permitindo que você selecione os arquivos necessários através de uma interface gráfica intuitiva.

## Requisitos

- Sistema Operacional: Windows
- Arquivo JSON contendo os dados dos produtos
- Arquivo Excel onde os resultados serão salvos

## Instalação

1. Clone o repositório:
    ```sh
    git clone https://github.com/seu-usuario/processador-de-arquivos.git
    ```
2. Navegue até o diretório do projeto:
    ```sh
    cd processador-de-arquivos
    ```
3. Instale as dependências (requer Python instalado):
    ```sh
    pip install pandas openpyxl tk
    ```

## Como Usar

### 1. Preparação

1. Certifique-se de ter o arquivo JSON contendo os dados dos produtos.
2. Tenha um arquivo Excel pronto para receber os dados processados.

### 2. Executando o Programa

1. **Abra o Processador de Arquivos**:
    - Execute o script Python:
    ```sh
    python gerador_tabela.py
    ```
    - Ou use o executável se disponível (`gerador_tabela.exe`).

2. **Selecione o Arquivo JSON**:
    - Clique no botão "Procurar" ao lado do campo "Selecione o arquivo JSON".
    - Navegue até o local onde o arquivo JSON está armazenado.
    - Selecione o arquivo JSON e clique em "Abrir".

3. **Selecione o Arquivo Excel**:
    - Clique no botão "Procurar" ao lado do campo "Selecione o arquivo Excel".
    - Navegue até o local onde o arquivo Excel está armazenado.
    - Selecione o arquivo Excel e clique em "Abrir".

4. **Processar os Arquivos**:
    - Com os arquivos JSON e Excel selecionados, clique no botão "Processar".
    - O programa irá carregar os dados do JSON, limpar e categorizar as informações, e salvar os resultados na aba "Base - Ordem de Estoque" do arquivo Excel.
    - Uma mensagem de sucesso será exibida ao final do processo.

### 3. Resultado

- O arquivo Excel selecionado será atualizado com os dados processados na aba "Base - Ordem de Estoque".
- As linhas onde a coluna "Categoria" estiver vazia serão removidas.
- Produtos com "manga curta" no nome serão categorizados como "Camisas Sociais MC".
- Todos os itens da categoria "Camisas Sociais" serão exibidos no console para verificação.

### 4. Erros Comuns e Soluções

- **Erro ao Selecionar Arquivos**: Certifique-se de que você selecionou corretamente os arquivos JSON e Excel antes de clicar em "Processar".
- **Arquivo JSON ou Excel Inválido**: Verifique se os arquivos selecionados estão no formato correto e não estão corrompidos.
- **Permissões de Arquivo**: Certifique-se de ter permissões adequadas para ler o arquivo JSON e modificar o arquivo Excel.

## Contribuição

Se você encontrar quaisquer problemas ou tiver sugestões de melhorias, por favor, sinta-se à vontade para abrir uma issue ou enviar um pull request.

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.
