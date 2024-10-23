Designações de Reunião - Aplicação Console
==========================================

Descrição
---------

Esta aplicação console foi desenvolvida para auxiliar na gestão de designações de reuniões. Ela permite exportar a programação de reuniões para arquivos Excel e Word, além de preencher automaticamente as designações de partes específicas com base em um arquivo Excel fornecido.

A aplicação utiliza um **WebScraper** para buscar a programação das reuniões de um site (exemplo fictício) e oferece três funcionalidades principais:

1.  **Exportar a programação de um mês específico para Excel**: Gera um arquivo Excel com a programação das reuniões de um mês específico, permitindo o preenchimento manual das designações.
2.  **Preencher designações com base em um arquivo Excel**: Preenche automaticamente as designações de reuniões com base em um arquivo Excel fornecido pelo usuário.
3.  **Exportar todas as programações disponíveis a partir do mês atual**: Gera arquivos Excel e Word com a programação de todas as reuniões disponíveis a partir do mês atual.

Funcionalidades
---------------

### 1\. Exportar programação de um mês específico para Excel

Esta opção permite que o usuário selecione um mês e ano específicos para exportar a programação das reuniões em um arquivo Excel. O arquivo gerado pode ser utilizado para preencher manualmente as designações.

### 2\. Preencher designações com base em um arquivo Excel

Com esta opção, o usuário pode fornecer um arquivo Excel contendo as designações preenchidas. A aplicação então preenche automaticamente as designações nas reuniões programadas, atualizando os campos de presidente, orações, sessões e partes.

### 3\. Exportar todas as programações disponíveis a partir do mês atual

Esta opção busca automaticamente todas as programações de reuniões disponíveis a partir do mês atual e exporta os dados para arquivos Excel e Word. A busca continua até que não haja mais programações disponíveis.

Requisitos
----------

-   .NET 6.0 ou superior
-   Pacotes NuGet:
    -   `iText7` para manipulação de PDFs
    -   `DocumentFormat.OpenXml` para manipulação de arquivos Word
    -   `EPPlus` para manipulação de arquivos Excel

Como usar
---------

### 1\. Clonar o repositório

`

`1git clone https://github.com/seu-usuario/designacoes-reuniao.git 2cd designacoes-reuniao`

`

### 2\. Configurar o ambiente

Certifique-se de que o ambiente está configurado corretamente para o uso do iText7 com o Bouncy Castle:

`

`1Environment.SetEnvironmentVariable("ITEXT_BOUNCY_CASTLE_FACTORY_NAME", "bouncy-castle");`

`

### 3\. Executar a aplicação

Compile e execute a aplicação:

`

`1dotnet run`

`

### 4\. Escolher uma opção

Ao iniciar a aplicação, você verá as seguintes opções no console:

`O que deseja fazer? 1. Exportar programação da reunião de um mês específico em excel para preenchimento das designações 2. Com base em arquivo excel, preencher designados das reuniões de um mês específico 3. Exportar todas as programações de reuniões disponíveis a partir do mês atual em excel para preenchimento das designações`

Digite o número da opção desejada e siga as instruções fornecidas.

### 5\. Exportar ou preencher designações

Dependendo da opção escolhida, a aplicação solicitará informações adicionais, como o mês, ano ou o caminho do arquivo Excel com as designações preenchidas.

Estrutura do Projeto
--------------------

-   **Program.cs**: Contém a lógica principal da aplicação, incluindo a interação com o usuário e a execução das funcionalidades.
-   **WebScraper.cs**: Responsável por buscar a programação das reuniões no site.
-   **ExcelExporter.cs**: Exporta a programação das reuniões para arquivos Excel.
-   **WordExporter.cs**: Exporta a programação das reuniões para arquivos Word.
-   **PdfEditor.cs**: Preenche automaticamente as partes dos estudantes em um arquivo PDF.

Exemplo de Uso
--------------

### Exportar programação de um mês específico

1.  Escolha a opção `1` no menu.
2.  Informe o mês e o ano desejados.
3.  A aplicação exportará a programação para um arquivo Excel no diretório `ExcelModelo`.

### Preencher designações com base em um arquivo Excel

1.  Escolha a opção `2` no menu.
2.  Informe o mês e o ano desejados.
3.  Forneça o caminho completo do arquivo Excel com as designações preenchidas.
4.  A aplicação preencherá automaticamente as designações e exportará os arquivos atualizados para os diretórios `ExcelDesignacoesPreenchidas` e `WordDesignacoesPreenchidas`.

### Exportar todas as programações disponíveis a partir do mês atual

1.  Escolha a opção `3` no menu.
2.  A aplicação buscará todas as programações disponíveis a partir do mês atual e exportará os arquivos para os diretórios `ExcelModelo` e `WordModelos`.

Estrutura de Arquivos Gerados
-----------------------------

-   **ExcelModelo**: Contém os arquivos Excel com a programação das reuniões para preenchimento manual.
-   **ExcelDesignacoesPreenchidas**: Contém os arquivos Excel com as designações preenchidas automaticamente.
-   **WordModelos**: Contém os arquivos Word com a programação das reuniões.
-   **WordDesignacoesPreenchidas**: Contém os arquivos Word com as designações preenchidas automaticamente.
-   **PartesEstudantes**: Contém os arquivos PDF com as partes dos estudantes preenchidas.

Contribuição
------------

Sinta-se à vontade para contribuir com melhorias ou correções. Para contribuir:

1.  Faça um fork do repositório.
2.  Crie uma nova branch para sua feature ou correção de bug.
3.  Envie um pull request com suas alterações.

Licença
-------

Este projeto está licenciado sob a [MIT License](https://file+.vscode-resource.vscode-cdn.net/c%3A/Users/Douglas/.vscode/extensions/stackspotai.stackspotai-1.8.8/packages/webview/LICENSE).

* * * * *

**Nota**: O site utilizado para buscar a programação das reuniões é fictício e serve apenas como exemplo.