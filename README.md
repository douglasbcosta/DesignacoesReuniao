Designa��es de Reuni�o - Aplica��o Console
==========================================

Descri��o
---------

Esta aplica��o console foi desenvolvida para auxiliar na gest�o de designa��es de reuni�es. Ela permite exportar a programa��o de reuni�es para arquivos Excel e Word, al�m de preencher automaticamente as designa��es de partes espec�ficas com base em um arquivo Excel fornecido.

A aplica��o utiliza um�**WebScraper**�para buscar a programa��o das reuni�es de um site (exemplo fict�cio) e oferece tr�s funcionalidades principais:

1.  **Exportar a programa��o de um m�s espec�fico para Excel**: Gera um arquivo Excel com a programa��o das reuni�es de um m�s espec�fico, permitindo o preenchimento manual das designa��es.
2.  **Preencher designa��es com base em um arquivo Excel**: Preenche automaticamente as designa��es de reuni�es com base em um arquivo Excel fornecido pelo usu�rio.
3.  **Exportar todas as programa��es dispon�veis a partir do m�s atual**: Gera arquivos Excel e Word com a programa��o de todas as reuni�es dispon�veis a partir do m�s atual.

Funcionalidades
---------------

### 1\. Exportar programa��o de um m�s espec�fico para Excel

Esta op��o permite que o usu�rio selecione um m�s e ano espec�ficos para exportar a programa��o das reuni�es em um arquivo Excel. O arquivo gerado pode ser utilizado para preencher manualmente as designa��es.

### 2\. Preencher designa��es com base em um arquivo Excel

Com esta op��o, o usu�rio pode fornecer um arquivo Excel contendo as designa��es preenchidas. A aplica��o ent�o preenche automaticamente as designa��es nas reuni�es programadas, atualizando os campos de presidente, ora��es, sess�es e partes.

### 3\. Exportar todas as programa��es dispon�veis a partir do m�s atual

Esta op��o busca automaticamente todas as programa��es de reuni�es dispon�veis a partir do m�s atual e exporta os dados para arquivos Excel e Word. A busca continua at� que n�o haja mais programa��es dispon�veis.

Requisitos
----------

-   .NET 6.0 ou superior
-   Pacotes NuGet:
    -   `iText7`�para manipula��o de PDFs
    -   `DocumentFormat.OpenXml`�para manipula��o de arquivos Word
    -   `EPPlus`�para manipula��o de arquivos Excel

Como usar
---------

### 1\. Clonar o reposit�rio

`

`1git clone https://github.com/seu-usuario/designacoes-reuniao.git 2cd designacoes-reuniao`

`

### 2\. Configurar o ambiente

Certifique-se de que o ambiente est� configurado corretamente para o uso do iText7 com o Bouncy Castle:

`

`1Environment.SetEnvironmentVariable("ITEXT_BOUNCY_CASTLE_FACTORY_NAME", "bouncy-castle");`

`

### 3\. Executar a aplica��o

Compile e execute a aplica��o:

`

`1dotnet run`

`

### 4\. Escolher uma op��o

Ao iniciar a aplica��o, voc� ver� as seguintes op��es no console:

`O que deseja fazer? 1. Exportar programa��o da reuni�o de um m�s espec�fico em excel para preenchimento das designa��es 2. Com base em arquivo excel, preencher designados das reuni�es de um m�s espec�fico 3. Exportar todas as programa��es de reuni�es dispon�veis a partir do m�s atual em excel para preenchimento das designa��es`

Digite o n�mero da op��o desejada e siga as instru��es fornecidas.

### 5\. Exportar ou preencher designa��es

Dependendo da op��o escolhida, a aplica��o solicitar� informa��es adicionais, como o m�s, ano ou o caminho do arquivo Excel com as designa��es preenchidas.

Estrutura do Projeto
--------------------

-   **Program.cs**: Cont�m a l�gica principal da aplica��o, incluindo a intera��o com o usu�rio e a execu��o das funcionalidades.
-   **WebScraper.cs**: Respons�vel por buscar a programa��o das reuni�es no site.
-   **ExcelExporter.cs**: Exporta a programa��o das reuni�es para arquivos Excel.
-   **WordExporter.cs**: Exporta a programa��o das reuni�es para arquivos Word.
-   **PdfEditor.cs**: Preenche automaticamente as partes dos estudantes em um arquivo PDF.

Exemplo de Uso
--------------

### Exportar programa��o de um m�s espec�fico

1.  Escolha a op��o�`1`�no menu.
2.  Informe o m�s e o ano desejados.
3.  A aplica��o exportar� a programa��o para um arquivo Excel no diret�rio�`ExcelModelo`.

### Preencher designa��es com base em um arquivo Excel

1.  Escolha a op��o�`2`�no menu.
2.  Informe o m�s e o ano desejados.
3.  Forne�a o caminho completo do arquivo Excel com as designa��es preenchidas.
4.  A aplica��o preencher� automaticamente as designa��es e exportar� os arquivos atualizados para os diret�rios�`ExcelDesignacoesPreenchidas`�e�`WordDesignacoesPreenchidas`.

### Exportar todas as programa��es dispon�veis a partir do m�s atual

1.  Escolha a op��o�`3`�no menu.
2.  A aplica��o buscar� todas as programa��es dispon�veis a partir do m�s atual e exportar� os arquivos para os diret�rios�`ExcelModelo`�e�`WordModelos`.

Estrutura de Arquivos Gerados
-----------------------------

-   **ExcelModelo**: Cont�m os arquivos Excel com a programa��o das reuni�es para preenchimento manual.
-   **ExcelDesignacoesPreenchidas**: Cont�m os arquivos Excel com as designa��es preenchidas automaticamente.
-   **WordModelos**: Cont�m os arquivos Word com a programa��o das reuni�es.
-   **WordDesignacoesPreenchidas**: Cont�m os arquivos Word com as designa��es preenchidas automaticamente.
-   **PartesEstudantes**: Cont�m os arquivos PDF com as partes dos estudantes preenchidas.

Contribui��o
------------

Sinta-se � vontade para contribuir com melhorias ou corre��es. Para contribuir:

1.  Fa�a um fork do reposit�rio.
2.  Crie uma nova branch para sua feature ou corre��o de bug.
3.  Envie um pull request com suas altera��es.

Licen�a
-------

Este projeto est� licenciado sob a�[MIT License](https://file+.vscode-resource.vscode-cdn.net/c%3A/Users/Douglas/.vscode/extensions/stackspotai.stackspotai-1.8.8/packages/webview/LICENSE).

* * * * *

**Nota**: O site utilizado para buscar a programa��o das reuni�es � fict�cio e serve apenas como exemplo.