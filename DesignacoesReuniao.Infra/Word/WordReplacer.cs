using DesignacoesReuniao.Domain.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace DesignacoesReuniao.Infra.Word
{
    public class WordReplacer
    {

        public void PreencherReunioesEmModelo(string caminhoModelo, string caminhoReuniaoPreenchida, List<Reuniao> reunioes)
        {
            FileInfo fileInfo = new FileInfo(caminhoReuniaoPreenchida);

            // Verifica se o diretório existe, se não, cria o diretório
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            // Copia o template para o destino
            File.Copy(caminhoModelo, caminhoReuniaoPreenchida, true);

            // Abre o documento Word
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(caminhoReuniaoPreenchida, true))
            {
                // Obtém o corpo do documento
                var body = wordDoc.MainDocumentPart.Document.Body;

                var substituicoesPadrao = GetSubstituicoesPadrao();
                foreach (var substituicao in substituicoesPadrao)
                {
                    ReplaceTodasOcorrencias(body, substituicao.Key, substituicao.Value);
                }
                foreach (var reuniao in reunioes)
                {
                    var substituicoes = GetSubstituicoes(reuniao);
                    foreach (var substituicao in substituicoes)
                    {
                        if (substituicao.ValorSubstituicao.Contains("Cântico"))
                        {
                            ReplaceCantico(body, substituicao.ValorSubstituicao);
                        }
                        if (string.IsNullOrEmpty(substituicao.Sessao))
                        {
                            ReplacePrimeiraOcorrencia(body, substituicao.ValorOriginal, substituicao.ValorSubstituicao);
                        }
                        else if (string.IsNullOrEmpty(substituicao.Tema))
                        {
                            ReplacePrimeiraOcorrenciaNaSessao(body, substituicao.ValorOriginal, substituicao.ValorSubstituicao, substituicao.Sessao);
                        }
                        else
                        {
                            ReplacePrimeiraOcorrenciaNaSessaoETema(body, substituicao.ValorOriginal, substituicao.ValorSubstituicao, substituicao.Sessao, substituicao.Tema);
                        }
                    }
                }
                RemoverLinhasComPartesVazias(body);
                RemoverLinhasComPrimeiraCelulaVazia(body);
                if (reunioes.Count < 5)
                {
                    RemoverTabelasAMais(body);
                }

                ReplaceIndices(body);
                AjustarFormatacaoEstudantes(body);
                // Chama o novo método para ajustar as linhas com 5 colunas
                AjustarLarguraLinhasComCincoColunas(body);
                // Chama o novo método para ajustar as linhas com 5 colunas
                AjustarLarguraLinhasComQuatroColunas(body);
                // Salva as alterações no documento
                RetirarParagrafosComSomenteDoisPontos(body);
                AdicionarQuebraDePagina(body);
                AdicionarHorariosReuniao(body);
                wordDoc.MainDocumentPart.Document.Save();
            }

            Console.WriteLine($"Arquivo Word gerado com sucesso em: {caminhoReuniaoPreenchida}");
        }

        private void AdicionarHorariosReuniao(Body body)
        {
            TimeOnly horario = new TimeOnly(20,0);
            string sessaoAtual = "";
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (GetSessoesReunioes().Contains(text.Text))
                        {
                            sessaoAtual = text.Text;
                        }

                        if (text.Text.Contains("Presidente"))
                        {
                            horario = new TimeOnly(20, 0);
                        }

                        if (text.Text.Contains("Cântico"))
                        {
                            horario = horario.AddMinutes(5);
                        }

                        int minutos = ExtrairTempo(text.Text);
                        if (minutos > 0)
                        {
                            if(sessaoAtual == "FAÇA SEU MELHOR NO MINISTÉRIO" && minutos <= 5) 
                            {
                                minutos++;
                            }

                            horario = horario.AddMinutes(minutos);
                        }
                        if (text.Text.Contains("0:00"))
                        {
                            text.Text = text.Text.Replace(text.Text, horario.ToString("HH:mm"));
                        }
                    }
                }
            }
        }

        private int ExtrairTempo(string texto)
        {
            var match = Regex.Match(texto, @"\((\d+)\s*min\)");
            if (match.Success)
            {
                return int.Parse(match.Groups[1].Value);
            }
            else
            {
                return 0;
            }
        }
        private void AdicionarQuebraDePagina(Body body)
        {
            int ocorrencias = 0;

            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Percorre todas as linhas da tabela
                foreach (var linha in tabela.Descendants<TableRow>())
                {
                    foreach (var celula in linha.Descendants<TableCell>())
                    {
                        foreach (var paragrafo in celula.Descendants<Paragraph>())
                        {
                            foreach (var run in paragrafo.Descendants<Run>())
                            {
                                foreach (var text in run.Descendants<Text>())
                                {
                                    if (text.Text.Contains("ANDORINHA DA MATA"))
                                    {
                                        ocorrencias++;

                                        // Adiciona uma quebra de página a partir da segunda ocorrência
                                        if (ocorrencias >= 2)
                                        {
                                            // Cria uma nova quebra de página
                                            var quebraDePagina = new Paragraph(
                                                new Run(
                                                    new Break() { Type = BreakValues.Page }
                                                )
                                            );

                                            // Insere a quebra de página antes da linha atual
                                            linha.InsertBeforeSelf(quebraDePagina);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void RetirarParagrafosComSomenteDoisPontos(Body body)
        {
            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Percorre todas as linhas da tabela
                foreach (var linha in tabela.Descendants<TableRow>())
                {
                    var celulas = linha.Descendants<TableCell>().ToList();

                    // Percorre todas as células da linha, exceto a última
                    for (int i = 0; i < celulas.Count - 1; i++)
                    {
                        var textoCelula = celulas[i].InnerText.Trim();
                        if (textoCelula.Contains("0:00"))
                        {
                            int width = ObterLarguraCelula(celulas[i]);
                            width = width + 90;
                            AjustarLarguraCelula(celulas[i], width);
                        }

                        // Verifica se a célula contém "Estudante"
                        if (textoCelula == ":")
                        {

                            celulas[i].RemoveAllChildren<Paragraph>();
                            break;
                        }
                    }
                }
            }
        }

       private void AjustarLarguraLinhasComCincoColunas(Body body)
        {
            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Percorre todas as linhas da tabela
                foreach (var linha in tabela.Descendants<TableRow>())
                {
                    var celulas = linha.Descendants<TableCell>().ToList();

                    // Verifica se a linha tem exatamente 5 colunas
                    if (celulas.Count == 5)
                    {
                        // Ajusta a largura de cada célula
                        for (int i = 0; i < celulas.Count; i++)
                        {
                            if (i == 1)
                            {
                                // Aumenta a largura da primeira célula
                                AjustarLarguraCelula(celulas[i], 4500);
                            }
                            if (i == 2)
                            {
                                // Aumenta a largura da segunda célula
                                AjustarLarguraCelula(celulas[i], 49);
                            }
                        }
                    }
                }
            }
        }
        private void AjustarLarguraLinhasComQuatroColunas(Body body)
        {
            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Percorre todas as linhas da tabela
                foreach (var linha in tabela.Descendants<TableRow>())
                {
                    var celulas = linha.Descendants<TableCell>().ToList();

                    // Verifica se a linha tem exatamente 5 colunas
                    if (celulas.Count == 4)
                    {
                        // Ajusta a largura de cada célula
                        for (int i = 0; i < celulas.Count; i++)
                        {
                            if (i == 1)
                            {
                                // Aumenta a largura da primeira célula
                                AjustarLarguraCelula(celulas[i], 4549);
                            }
                            else if (i == 2)
                            {
                                // Aumenta a largura da segunda célula
                                AjustarLarguraCelula(celulas[i], 49);
                            }
                        }
                    }
                }
            }
        }
        private void AjustarFormatacaoEstudantes(Body body)
        {
            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Percorre todas as linhas da tabela
                foreach (var linha in tabela.Descendants<TableRow>())
                {
                    var celulas = linha.Descendants<TableCell>().ToList();

                    // Percorre todas as células da linha, exceto a última
                    for (int i = 0; i < celulas.Count - 1; i++)
                    {
                        var textoCelula = celulas[i].InnerText.Trim();

                        // Verifica se a célula contém "Estudante"
                        if (textoCelula.Contains("Estudante"))
                        {
                            // Move o conteúdo da célula atual para a célula seguinte, copiando também a formatação
                            var proximaCelula = celulas[i + 1];
                            proximaCelula.RemoveAllChildren<Paragraph>();

                            // Copia todos os parágrafos da célula original para a célula seguinte
                            foreach (var paragraph in celulas[i].Elements<Paragraph>())
                            {
                                // Clona o parágrafo para manter a formatação
                                var novoParagrafo = (Paragraph)paragraph.CloneNode(true);
                                proximaCelula.Append(novoParagrafo);
                            }
                            celulas[i].RemoveAllChildren<Paragraph>();

                            // Sai do loop para evitar problemas com a mesclagem
                            break;
                        }
                    }
                }
            }
        }
        private int ObterLarguraCelula(TableCell celula)
        {
            var largura = 0;
            var cellWidth = celula.TableCellProperties?.TableCellWidth;

            if (cellWidth != null && cellWidth.Width != null)
            {
                largura = int.Parse(cellWidth.Width);
            }

            return largura;
        }

        private void AjustarLarguraCelula(TableCell celula, int novaLargura)
        {
            if (celula.TableCellProperties == null)
            {
                celula.TableCellProperties = new TableCellProperties();
            }

            var cellWidth = celula.TableCellProperties.TableCellWidth;
            if (cellWidth == null)
            {
                cellWidth = new TableCellWidth();
                celula.TableCellProperties.Append(cellWidth);
            }

            cellWidth.Width = novaLargura.ToString();
            cellWidth.Type = TableWidthUnitValues.Dxa; // Define a unidade de medida como twips (1/20 de ponto)
        }

        private static void ReplaceTodasOcorrencias(Body? body, string textoOriginal, string textoAlterado)
        {
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (text.Text.Contains(textoOriginal))
                        {
                            text.Text = text.Text.Replace(textoOriginal, textoAlterado);

                        }
                    }
                }
            }
        }

        private static void ReplaceValorIgual(Body? body, string textoOriginal, string textoAlterado)
        {
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (text.Text == textoOriginal)
                        {
                            text.Text = text.Text.Replace(textoOriginal, textoAlterado);

                        }
                    }
                }
            }
        }
        private static void ReplaceCantico(Body? body, string canticoAlteracao)
        {
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (text.Text == "Cântico número")
                        {
                            text.Text = text.Text.Replace("Cântico número", canticoAlteracao);
                        }

                        if (text.Text == "número")
                        {
                            text.Text = text.Text.Replace("número", canticoAlteracao.Split(' ')[1]);
                        }
                    }
                }
            }
        }
        private static void ReplaceIndices(Body? body)
        {
            int indice = 0;
            Regex regexNumeracao = new Regex(@"^\d+\.$");
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (!text.Text.Contains("1.º") && text.Text.Contains("."))
                        {
                            if (text.Text.Contains("1."))
                            {
                                indice = 1;
                            }
                            if (text.Text != $"{indice}.")
                            {
                                text.Text = text.Text.Replace(text.Text, $"{indice}.");
                            }
                            indice++;
                        }
                    }
                }
            }
        }
        private static void ReplacePrimeiraOcorrencia(Body? body, string textoOriginal, string textoAlterado)
        {
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (text.Text.Contains(textoOriginal))
                        {
                            text.Text = text.Text.Replace(text.Text, textoAlterado);
                            return;
                        }
                    }
                }
            }
        }

        private void ReplacePrimeiraOcorrenciaNaSessao(Body? body, string textoOriginal, string textoAlterado, string sessao)
        {
            string sessaoAtual = "";
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (GetSessoesReunioes().Contains(text.Text))
                        {
                            sessaoAtual = text.Text;
                        }

                        if (sessaoAtual.Contains(sessao) && text.Text.Contains(textoOriginal))
                        {
                            text.Text = text.Text.Replace(text.Text, textoAlterado);
                            return;
                        }
                    }
                }
            }
        }

        private void ReplacePrimeiraOcorrenciaNaSessaoETema(Body? body, string textoOriginal, string textoAlterado, string sessao, string tema)
        {
            string sessaoAtual = "";
            string temaAtual = "";
            // Substitui os textos conforme o dicionário de substituições
            foreach (var paragrafo in body.Descendants<Paragraph>())
            {
                foreach (var run in paragrafo.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (text.Text.Contains(tema))
                        {
                            temaAtual = text.Text;
                        }
                        if (GetSessoesReunioes().Contains(text.Text))
                        {
                            sessaoAtual = text.Text;
                        }

                        if (sessaoAtual.Contains(sessao) && temaAtual.Contains(tema) && text.Text.Contains(textoOriginal))
                        {
                            text.Text = text.Text.Replace(text.Text, textoAlterado);
                            return;
                        }
                    }
                }
            }
        }

        private List<Substituicao> GetSubstituicoes(Reuniao reunioes)
        {
            List<Substituicao> substiticoes = new List<Substituicao>();

            substiticoes.Add(new Substituicao("DATA | LEITURA SEMANAL DA BÍBLIA", $"{reunioes.Semana} | {reunioes.LeituraDaSemana}"));
            foreach (var cantico in reunioes.Canticos)
            {
                substiticoes.Add(new Substituicao("Cântico número", cantico));
            }
            substiticoes.Add(new Substituicao("Nome", reunioes.Presidente));
            substiticoes.Add(new Substituicao("Nome", ""));
            if (string.IsNullOrEmpty(reunioes.OracaoInicial)) 
            {
                substiticoes.Add(new Substituicao("Oração", ""));
            }

            substiticoes.Add(new Substituicao("Nome", reunioes.OracaoInicial));
            foreach (var sessao in reunioes.Sessoes)
            {
                foreach (var parte in sessao.Partes)
                {
                    GerarSubstituicoesTesouros(substiticoes, sessao, parte);
                    GerarSubstituicoesMinisterio(substiticoes, sessao, parte);
                    GerarSubstituicoesVidaCrista(substiticoes, sessao, parte);
                }
                if (sessao.TituloSessao == "FAÇA SEU MELHOR NO MINISTÉRIO")
                {
                    int diferencaQtdPartes = 4 - sessao.Partes.Count;
                    for (int i = 0; i < diferencaQtdPartes; i++)
                    {
                        substiticoes.Add(new Substituicao("Tema", "", sessao.TituloSessao));
                        substiticoes.Add(new Substituicao("(X min)", "", sessao.TituloSessao));
                        substiticoes.Add(new Substituicao("Nome/Nome", "", sessao.TituloSessao));
                        substiticoes.Add(new Substituicao("Nome/Nome", "", sessao.TituloSessao));
                    }
                }
                if (sessao.TituloSessao == "NOSSA VIDA CRISTÃ")
                {
                    int diferencaQtdPartes = 3 - sessao.Partes.Count;
                    for (int i = 0; i < diferencaQtdPartes; i++)
                    {
                        substiticoes.Add(new Substituicao("Tema", "", sessao.TituloSessao));
                        substiticoes.Add(new Substituicao("(XX min)", "", sessao.TituloSessao));
                        substiticoes.Add(new Substituicao("Nome", "", sessao.TituloSessao));
                    }
                }

            }
            if (string.IsNullOrEmpty(reunioes.OracaoFinal))
            {
                substiticoes.Add(new Substituicao("Oração", ""));
            }
            substiticoes.Add(new Substituicao("Nome", reunioes.OracaoFinal));

            return substiticoes;

        }

        private static void GerarSubstituicoesTesouros(List<Substituicao> substiticoes, Sessao sessao, Parte parte)
        {
            if (sessao.TituloSessao == "TESOUROS DA PALAVRA DE DEUS")
            {
                if (parte.TituloParte.Contains("Joias espirituais"))
                {
                    substiticoes.Add(new Substituicao("Joias espirituais", parte.TituloParte, sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("Nome", parte.Designados.FirstOrDefault(), sessao.TituloSessao));
                }
                else if (parte.TituloParte.Contains("Leitura da Bíblia"))
                {
                    substiticoes.Add(new Substituicao("Leitura da Bíblia", parte.TituloParte, sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("Nome", "", sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("Nome", parte.Designados.FirstOrDefault(), sessao.TituloSessao));

                }
                else
                {
                    substiticoes.Add(new Substituicao("Tema", parte.TituloParte, sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("Nome", parte.Designados.FirstOrDefault(), sessao.TituloSessao));
                }
            }
        }

        private static void GerarSubstituicoesMinisterio(List<Substituicao> substiticoes, Sessao sessao, Parte parte)
        {
            if (sessao.TituloSessao == "FAÇA SEU MELHOR NO MINISTÉRIO")
            {
                substiticoes.Add(new Substituicao("Tema", parte.TituloParte, sessao.TituloSessao));
                substiticoes.Add(new Substituicao("(X min)", $"({parte.TempoMinutos} min)", sessao.TituloSessao));
                substiticoes.Add(new Substituicao("Nome/Nome", "", sessao.TituloSessao));
                if (parte.TempoMinutos > 5)
                {
                    substiticoes.Add(new Substituicao("Estudante/ajudante", "", sessao.TituloSessao, parte.TituloParte));
                }

                if (parte.Designados.Count == 1 || parte.TituloParte.Contains("Discurso"))
                {
                    substiticoes.Add(new Substituicao("Estudante/ajudante", "Estudante", sessao.TituloSessao, parte.TituloParte));
                }
                substiticoes.Add(new Substituicao("Nome/Nome", string.Join("/ ", parte.Designados), sessao.TituloSessao));
            }
        }

        private static void GerarSubstituicoesVidaCrista(List<Substituicao> substiticoes, Sessao sessao, Parte parte)
        {
            if (sessao.TituloSessao == "NOSSA VIDA CRISTÃ")
            {
                if (parte.TituloParte.Contains("Estudo bíblico de congregação"))
                {
                    substiticoes.Add(new Substituicao("Estudo bíblico de congregação", parte.TituloParte, sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("(30 min)", $"({parte.TempoMinutos} min)", sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("Nome/Nome", parte.Designados.FirstOrDefault(), sessao.TituloSessao));
                }
                else
                {
                    substiticoes.Add(new Substituicao("Tema", parte.TituloParte, sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("(XX min)", $"({parte.TempoMinutos} min)", sessao.TituloSessao));
                    substiticoes.Add(new Substituicao("Nome", parte.Designados.FirstOrDefault(), sessao.TituloSessao));
                }
            }
        }

        private string[] GetSessoesReunioes()
        {
            return new string[] { "TESOUROS DA PALAVRA DE DEUS", "FAÇA SEU MELHOR NO MINISTÉRIO", "NOSSA VIDA CRISTÃ" };
        }

        private Dictionary<string, string> GetSubstituicoesPadrao()
        {
            Dictionary<string, string> substiticoes = new Dictionary<string, string>();

            substiticoes.Add("[", "");
            substiticoes.Add("]", "");
            substiticoes.Add("NOME DA CONGREGAÇÃO", "ANDORINHA DA MATA");
            substiticoes.Add("Conselheiro da sala B", "");
            substiticoes.Add("Sala B", "");
            substiticoes.Add("Dirigente/leitor", "Dirigente");

            return substiticoes;

        }

        private string FormatarTextoComPrimeiraLetraMaiuscula(string texto)
        {
            if (string.IsNullOrEmpty(texto))
            {
                return texto;
            }
            return System.Text.RegularExpressions.Regex.Replace(texto.ToLower(), @"\b\w", m => m.Value.ToUpper());
        }

        public void RemoverLinhasComPartesVazias(Body body)
        {

            Regex regexNumeracao = new Regex(@"^\d+\.$");
            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Lista de linhas a serem removidas
                var linhasParaRemover = tabela.Descendants<TableRow>()
                    .Where(linha =>
                    {
                        // Obtém a célula da coluna especificada (indiceColuna)
                        var celula = linha.Descendants<TableCell>().ElementAtOrDefault(1);
                        if (celula != null)
                        {
                            // Verifica se a célula está vazia
                            var textoCelula = celula.InnerText.Trim();
                            return regexNumeracao.IsMatch(textoCelula);
                        }
                        return false;
                    }).ToList();

                // Remove as linhas que atendem à condição
                foreach (var linha in linhasParaRemover)
                {
                    linha.Remove();
                }
            }
        }
        public void RemoverLinhasComPrimeiraCelulaVazia(Body body)
        {

            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Lista de linhas a serem removidas
                var linhasParaRemover = tabela.Descendants<TableRow>()
                    .Where((linha, index) =>
                    {
                        // Obtém a célula da coluna especificada (indiceColuna)
                        var celula = linha.Descendants<TableCell>().ElementAtOrDefault(0);
                        if (celula != null)
                        {
                            // Verifica se a célula está vazia
                            var textoCelula = celula.InnerText.Trim();
                            if (string.IsNullOrEmpty(textoCelula))
                            {
                                // Verifica se existe uma linha abaixo
                                var proximaLinha = tabela.Descendants<TableRow>().ElementAtOrDefault(index + 1);
                                if (proximaLinha != null)
                                {
                                    // Obtém a célula da próxima linha
                                    var celulaProximaLinha = proximaLinha.Descendants<TableCell>().ElementAtOrDefault(0);
                                    if (celulaProximaLinha != null)
                                    {
                                        var textoProximaCelula = celulaProximaLinha.InnerText.Trim();

                                        // Verifica se a célula da próxima linha contém uma das palavras da lista
                                        foreach (var palavra in GetSessoesReunioes())
                                        {
                                            if (textoProximaCelula.Contains(palavra, StringComparison.OrdinalIgnoreCase))
                                            {
                                                return true; // A linha atual será removida
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        return false;
                    }).ToList();

                // Remove as linhas que atendem à condição
                foreach (var linha in linhasParaRemover)
                {
                    linha.Remove();
                }
            }
        }
        public void RemoverTabelasAMais(Body body)
        {
            int ocorrencias = 0;
            bool encontrouQuintaOcorrencia = false;

            // Percorre todas as tabelas no documento
            foreach (var tabela in body.Descendants<Table>())
            {
                // Lista de linhas a serem removidas
                var linhasParaRemover = tabela.Descendants<TableRow>()
                    .Where(linha =>
                    {
                        // Obtém a célula da coluna especificada (indiceColuna)
                        var celula = linha.Descendants<TableCell>().ElementAtOrDefault(0);
                        if (encontrouQuintaOcorrencia)
                        {
                            return true;
                        }
                        if (celula != null)
                        {
                            // Verifica se a célula está vazia
                            var textoCelula = celula.InnerText.Trim();
                            if(textoCelula.Contains("ANDORINHA DA MATA"))
                            {
                                ocorrencias++;
                            }
                            if(ocorrencias == 3)
                            {
                                encontrouQuintaOcorrencia = true;
                                return true;
                            }
                        }
                        return false;
                    }).ToList();

                // Remove as linhas que atendem à condição
                foreach (var linha in linhasParaRemover)
                {
                    linha.Remove();
                }
            }
        }
    }
}
