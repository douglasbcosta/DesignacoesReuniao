using DesignacoesReuniao.Domain.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DesignacoesReuniao.Infra.Word
{
    public class WordExporter 
    {
        // Definindo as cores de fundo como variáveis globais (constantes)
        private const string COR_TESOUROS_DA_PALAVRA_DE_DEUS = "2a6b77";
        private const string COR_FACA_SEU_MELHOR_NO_MINISTERIO = "9b6d17";
        private const string COR_NOSSA_VIDA_CRISTA = "942926";
        private const string COR_CINZA = "808080"; // Cor cinza para a segunda coluna

        public void ExportarReunioesParaWord(List<Reuniao> reunioes, string caminhoArquivo)
        {
            FileInfo fileInfo = new FileInfo(caminhoArquivo);

            // Verifica se o diretório existe, se não, cria o diretório
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            string caminhoCompleto = fileInfo.FullName;
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(caminhoArquivo, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                AdicionarCabecalho(wordDocument);
                DefinirMargens(body);

                int contadorReunioes = 0;
                int qtdReunioes = reunioes.Count;

                foreach (var reuniao in reunioes)
                {
                    contadorReunioes++;
                    Table tabela = CriarTabelaInvisivel();

                    AdicionarLinhaNaTabelaComNegrito(tabela, $"{reuniao.Semana} | {reuniao.LeituraDaSemana}", "Presidente:", reuniao.Presidente?.ToString());
                    AdicionarLinhaNaTabela(tabela, reuniao.Canticos[0], "Oração Inicial:", reuniao.OracaoInicial?.ToString());
                    body.Append(tabela);

                    foreach (var sessao in reuniao.Sessoes)
                    {
                        string corFundo = ObterCorFundoPorSessao(sessao.TituloSessao);
                        AdicionarParagrafoComCorDeFundo(body, sessao.TituloSessao, JustificationValues.Left, corFundo, bold: true, corTexto: "FFFFFF");

                        if (sessao.TituloSessao == "NOSSA VIDA CRISTÃ")
                        {
                            AdicionarParagrafo(body, reuniao.Canticos[1]);
                        }

                        tabela = CriarTabelaInvisivel();
                        foreach (var parte in sessao.Partes)
                        {
                            string textoSegundaColuna = ObterTextoSegundaColuna(sessao.TituloSessao, parte.TituloParte);
                            string designados = parte.ObterNomesDesignadoEAjudante();
                            AdicionarLinhaNaTabela(tabela, $"{parte.TituloParte} ({parte.TempoMinutos} min)", textoSegundaColuna, designados);
                        }
                        body.Append(tabela);
                    }

                    AdicionarLinhaNaTabela(tabela, reuniao.Canticos[2], "Oração Final:", reuniao.OracaoFinal.ToString());
                    VerificarQuebraDePagina(body, qtdReunioes, contadorReunioes);
                }

                mainPart.Document.Append(body);
                mainPart.Document.Save();
                
            }

            Console.WriteLine($"Arquivo Word criado com sucesso em: {caminhoCompleto}");
        }

        private void DefinirMargens(Body body)
        {
            SectionProperties sectionProperties = new SectionProperties();
            PageMargin pageMargin = new PageMargin()
            {
                Top = 720, // 0.5 polegadas
                Right = 1440, // 0.5 polegadas
                Bottom = 0, // 0.5 polegadas
                Left = 720 // 0.5 polegadas
            };
            sectionProperties.Append(pageMargin);
            body.Append(sectionProperties);
        }

        private void VerificarQuebraDePagina(Body body, int qtdReunioes, int contadorReunioes)
        {
            if (qtdReunioes != contadorReunioes && contadorReunioes % 2 == 0)
            {
                AdicionarQuebraDePagina(body);
            }
            else if (qtdReunioes != contadorReunioes && contadorReunioes % 2 != 0)
            {
                AdicionarParagrafo(body, "");
            }
        }

        private void AdicionarCabecalho(WordprocessingDocument wordDocument)
        {
            HeaderPart headerPart = wordDocument.MainDocumentPart.AddNewPart<HeaderPart>();
            string headerPartId = wordDocument.MainDocumentPart.GetIdOfPart(headerPart);

            Header header = new Header();
            Table tabela = CriarTabelaInvisivel();

            TableCell celulaColuna1 = CriarCelula("ANDORINHA DA MATA", "3000", true, "000000", "1");
            TableCell celulaColuna2 = CriarCelula("Programação da reunião do meio de semana", "7000", true, "000000", "2", "28", JustificationValues.Right);

            AdicionarLinhaNaTabela(tabela, celulaColuna1, celulaColuna2);
            header.Append(tabela);

            headerPart.Header = header;
            AssociarCabecalhoAoDocumento(wordDocument, headerPartId);
        }

        private void AssociarCabecalhoAoDocumento(WordprocessingDocument wordDocument, string headerPartId)
        {
            if (wordDocument.MainDocumentPart.Document.Body == null)
            {
                wordDocument.MainDocumentPart.Document.Body = new Body();
            }

            SectionProperties sectionProperties = new SectionProperties();
            HeaderReference headerReference = new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId };
            sectionProperties.Append(headerReference);
            wordDocument.MainDocumentPart.Document.Body.Append(sectionProperties);
        }

        private TableCell CriarCelula(string texto, string largura, bool negrito, string corTexto, string margemInferior, string tamanhoFonte = "24", JustificationValues? alinhamento = null)
        {
            // Define o alinhamento padrão como JustificationValues.Left, caso não seja fornecido
            JustificationValues alinhamentoFinal = alinhamento ?? JustificationValues.Left;

            Run run = new Run(new Text(texto));
            run.RunProperties = new RunProperties(new Bold(), new FontSize { Val = tamanhoFonte }, new Color { Val = corTexto });

            ParagraphProperties paragraphProperties = new ParagraphProperties(new Justification { Val = alinhamentoFinal });
            TableCell celula = new TableCell(new Paragraph(paragraphProperties, run));
            celula.TableCellProperties = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = largura },
                new TableCellBorders(new BottomBorder { Val = BorderValues.Double, Size = 12, Color = "000000" }),
                new TableCellMargin { BottomMargin = new BottomMargin { Width = margemInferior, Type = TableWidthUnitValues.Dxa } }
            );
            return celula;
        }

        private void AdicionarQuebraDePagina(Body body)
        {
            Paragraph paragraph = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
            body.Append(paragraph);
        }

        private void AdicionarParagrafo(Body body, string texto)
        {
            Paragraph paragraph = new Paragraph(new Run(new Text(texto)));
            body.Append(paragraph);
        }

        private void AdicionarParagrafoComCorDeFundo(Body body, string texto, JustificationValues alinhamento, string corHex, bool bold = false, string corTexto = "000000")
        {
            // Cria o Run com o texto
            Run run = new Run(new Text(texto));
            run.RunProperties = new RunProperties(bold ? new Bold() : null, new Color { Val = corTexto });

            // Define a cor de fundo para o parágrafo
            Shading shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = corHex };

            // Propriedades do parágrafo, incluindo a indentação do texto
            ParagraphProperties paragraphProperties = new ParagraphProperties(new Justification { Val = alinhamento });


            // Aplica o fundo colorido ao parágrafo
            paragraphProperties.Shading = shading;

            // Cria o parágrafo com as propriedades e o texto
            Paragraph paragraph = new Paragraph(run);
            paragraph.ParagraphProperties = paragraphProperties;

            // Adiciona o parágrafo ao corpo do documento
            body.Append(paragraph);
        }

        private Table CriarTabelaInvisivel()
        {
            Table tabela = new Table();
            TableProperties props = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.None },
                    new BottomBorder { Val = BorderValues.None },
                    new LeftBorder { Val = BorderValues.None },
                    new RightBorder { Val = BorderValues.None },
                    new InsideHorizontalBorder { Val = BorderValues.None },
                    new InsideVerticalBorder { Val = BorderValues.None }
                ),
                new TableCellSpacing() { Width = "0" }
            );
            tabela.AppendChild(props);
            return tabela;
        }

        private void AdicionarLinhaNaTabela(Table tabela, TableCell celula1, TableCell celula2)
        {
            TableRow linha = new TableRow();
            linha.Append(celula1, celula2);
            tabela.Append(linha);
        }

        private void AdicionarLinhaNaTabela(Table tabela, TableCell celula1, TableCell celula2, TableCell celula3)
        {
            TableRow linha = new TableRow();
            linha.Append(celula1, celula2, celula3);
            tabela.Append(linha);
        }

        private void AdicionarLinhaNaTabela(Table tabela, string coluna1, string coluna2, string coluna3)
        {
            TableCell celulaColuna1 = CriarCelulaSimples(coluna1, "6000");
            TableCell celulaColuna2 = AdicionarColuna2(coluna2);
            TableCell celulaColuna3 = AdicionarColuna3(coluna3);

            AdicionarLinhaNaTabela(tabela, celulaColuna1, celulaColuna2, celulaColuna3);
        }

        private TableCell CriarCelulaSimples(string texto, string largura, bool negrito = false)
        {
            // Cria o Run com o texto
            Run run = new Run(new Text(texto));

            // Se o parâmetro negrito for true, aplica a formatação de negrito
            if (negrito)
            {
                run.RunProperties = new RunProperties(new Bold());
            }

            // Define as propriedades do parágrafo
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.SpacingBetweenLines = new SpacingBetweenLines() { Before = "0", After = "0" };

            // Cria a célula com o parágrafo e o texto
            TableCell celula = new TableCell(new Paragraph(paragraphProperties, run));

            // Define as propriedades da célula, como largura e margens
            celula.TableCellProperties = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = largura },
                new TableCellMargin
                {
                    TopMargin = new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa }
                }
            );

            return celula;
        }


        private TableCell AdicionarColuna3(string coluna3)
        {
            // Cria a célula para a terceira coluna (intermediária) com espaçamento inicial
            TableCell celulaColuna3 = new TableCell(new Paragraph(new Run(new Text(coluna3))));

            // Define as propriedades da célula, como largura e margens
            celulaColuna3.TableCellProperties = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }, // Largura da célula
                new TableCellMargin
                {
                    LeftMargin = new LeftMargin { Width = "200", Type = TableWidthUnitValues.Dxa }, // Espaçamento à esquerda
                    TopMargin = new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa }, // Reduz o espaçamento superior
                    BottomMargin = new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa } // Reduz o espaçamento inferior
                }
            );

            // Retorna a célula criada
            return celulaColuna3;
        }
        private static TableCell AdicionarColuna2(string coluna2)
        {
            return CriarCelulaComMargem(coluna2, "2000", "0", JustificationValues.Right, "18", "808080", true);
        }

        private static TableCell CriarCelulaComMargem(string texto, string largura, string margemEsquerda, JustificationValues? alinhamento = null, string tamanhoFonte = "24", string corTexto = "000000", bool negrito = false)
        {
            // Define o alinhamento padrão como JustificationValues.Left, caso não seja fornecido
            JustificationValues alinhamentoFinal = alinhamento ?? JustificationValues.Left;

            Run run = new Run(new Text(texto));
            run.RunProperties = new RunProperties(negrito ? new Bold() : null, new FontSize { Val = tamanhoFonte }, new Color { Val = corTexto });

            ParagraphProperties paragraphProperties = new ParagraphProperties(new Justification { Val = alinhamentoFinal });
            TableCell celula = new TableCell(new Paragraph(paragraphProperties, run));
            celula.TableCellProperties = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = largura },
                new TableCellMargin
                {
                    LeftMargin = new LeftMargin { Width = margemEsquerda, Type = TableWidthUnitValues.Dxa },
                    TopMargin = new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa }
                }
            );
            return celula;
        }

        private string ObterTextoSegundaColuna(string tituloSessao, string tituloParte)
        {
            return tituloSessao switch
            {
                "TESOUROS DA PALAVRA DE DEUS" when tituloParte.Contains("Leitura da Bíblia") => "Estudante:",
                "FAÇA SEU MELHOR NO MINISTÉRIO" when tituloParte.Contains("Discurso") => "Estudante:",
                "FAÇA SEU MELHOR NO MINISTÉRIO" => "Estudante/ajudante:",
                "NOSSA VIDA CRISTÃ" when tituloParte.Contains("Estudo bíblico de congregação") => "Dirigente:",
                _ => ""
            };
        }

        private void AdicionarLinhaNaTabelaComNegrito(Table tabela, string coluna1, string coluna2, string coluna3)
        {
            TableCell celulaColuna1 = CriarCelulaSimples(coluna1, "6000", true);
            TableCell celulaColuna2 = AdicionarColuna2(coluna2);
            TableCell celulaColuna3 = AdicionarColuna3(coluna3);

            AdicionarLinhaNaTabela(tabela, celulaColuna1, celulaColuna2, celulaColuna3);
        }

        private string ObterCorFundoPorSessao(string tituloSessao)
        {
            return tituloSessao switch
            {
                "TESOUROS DA PALAVRA DE DEUS" => COR_TESOUROS_DA_PALAVRA_DE_DEUS,
                "FAÇA SEU MELHOR NO MINISTÉRIO" => COR_FACA_SEU_MELHOR_NO_MINISTERIO,
                "NOSSA VIDA CRISTÃ" => COR_NOSSA_VIDA_CRISTA,
                _ => "FFFFFF"
            };
        }
    }
}