using DesignacoesReuniao.Domain.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace DesignacoesReuniao.Infra.Excel
{
    public class ExcelExporter
    {
        // Definindo as cores de fundo como variáveis globais (constantes)
        private const string COR_TESOUROS_DA_PALAVRA_DE_DEUS = "#2a6b77";
        private const string COR_FACA_SEU_MELHOR_NO_MINISTERIO = "#9b6d17";
        private const string COR_NOSSA_VIDA_CRISTA = "#942926";
        private const string COR_PADRAO = "#f4a261"; // Cor padrão 

        // Definindo os nomes das sessões como constantes globais
        private const string SESSAO_TESOUROS_DA_PALAVRA_DE_DEUS = "TESOUROS DA PALAVRA DE DEUS";
        private const string SESSAO_FACA_SEU_MELHOR_NO_MINISTERIO = "FAÇA SEU MELHOR NO MINISTÉRIO";
        private const string SESSAO_NOSSA_VIDA_CRISTA = "NOSSA VIDA CRISTÃ";

        public void ExportarReunioesParaExcel(List<Reuniao> reunioes, string caminhoArquivo)
        {
            // Configura a licença do EPPlus (obrigatório a partir da versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage())
            {
                // Cria uma nova planilha
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Reuniões");

                // Configura o cabeçalho
                ConfigurarCabecalho(worksheet);

                int linhaAtual = 2; // Começa na linha 2, pois a linha 1 é o cabeçalho

                // Preenche os dados das reuniões
                foreach (var reuniao in reunioes)
                {
                    // Adiciona as linhas de Presidente e Oração Inicial
                    linhaAtual = AdicionarLinhaPresidenteOracaoInicial(worksheet, reuniao, linhaAtual);

                    // Adiciona as sessões e partes da reunião
                    linhaAtual = AdicionarSessoes(worksheet, reuniao, linhaAtual);

                    // Adiciona a linha de Oração Final
                    linhaAtual = AdicionarLinhaOracaoFinal(worksheet, reuniao, linhaAtual);

                    // Adiciona uma borda para separar as semanas
                    AplicarBordaSeparadora(worksheet, linhaAtual - 1);
                }

                // Aplica bordas ao redor de toda a tabela
                AplicarBordasTabela(worksheet, linhaAtual - 1);

                // Auto ajusta as colunas
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Salva o arquivo Excel no caminho especificado
                SalvarArquivoExcel(package, caminhoArquivo);
            }
        }

        private void ConfigurarCabecalho(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "Semana";
            worksheet.Cells[1, 2].Value = "Sessão";
            worksheet.Cells[1, 3].Value = "Parte";
            worksheet.Cells[1, 4].Value = "Tempo (min)";
            worksheet.Cells[1, 5].Value = "Designados";

            // Aplica borda mais espessa no cabeçalho
            using (var range = worksheet.Cells[1, 1, 1, 5])
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thick;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thick;
                range.Style.Border.Top.Color.SetColor(Color.Black);
                range.Style.Border.Bottom.Color.SetColor(Color.Black);
                range.Style.Border.Left.Color.SetColor(Color.Black);
                range.Style.Border.Right.Color.SetColor(Color.Black);

                // Aplica a cor de fundo do cabeçalho 
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(COR_PADRAO)); 
            }
        }

        private int AdicionarLinhaPresidenteOracaoInicial(ExcelWorksheet worksheet, Reuniao reuniao, int linhaAtual)
        {
            // Adiciona a linha para o Presidente
            AdicionarLinha(worksheet, reuniao.Semana, SESSAO_TESOUROS_DA_PALAVRA_DE_DEUS, "Presidente", reuniao.Presidente, COR_TESOUROS_DA_PALAVRA_DE_DEUS, linhaAtual);
            linhaAtual++;

            // Adiciona a linha para a Oração Inicial
            AdicionarLinha(worksheet, reuniao.Semana, SESSAO_TESOUROS_DA_PALAVRA_DE_DEUS, "Oração Inicial", reuniao.OracaoInicial, COR_TESOUROS_DA_PALAVRA_DE_DEUS, linhaAtual);
            linhaAtual++;

            return linhaAtual;
        }

        private int AdicionarSessoes(ExcelWorksheet worksheet, Reuniao reuniao, int linhaAtual)
        {
            foreach (var sessao in reuniao.Sessoes)
            {
                foreach (var parte in sessao.Partes)
                {
                    string corFundo = ObterCorFundoPorSessao(sessao.TituloSessao);
                    AdicionarLinha(worksheet, reuniao.Semana, sessao.TituloSessao, parte.TituloParte, string.Join(", ", parte.Designados), corFundo, linhaAtual, parte.TempoMinutos);
                    linhaAtual++;
                }
            }
            return linhaAtual;
        }

        private int AdicionarLinhaOracaoFinal(ExcelWorksheet worksheet, Reuniao reuniao, int linhaAtual)
        {
            // Adiciona a linha para a Oração Final
            AdicionarLinha(worksheet, reuniao.Semana, SESSAO_NOSSA_VIDA_CRISTA, "Oração Final", reuniao.OracaoFinal, COR_NOSSA_VIDA_CRISTA, linhaAtual);
            linhaAtual++;

            return linhaAtual;
        }

        private void AdicionarLinha(ExcelWorksheet worksheet, string semana, string sessao, string parte, string designados, string corFundo, int linha, int tempoMinutos = 0)
        {
            worksheet.Cells[linha, 1].Value = semana;
            worksheet.Cells[linha, 2].Value = sessao;
            worksheet.Cells[linha, 3].Value = parte;
            worksheet.Cells[linha, 4].Value = tempoMinutos > 0 ? (object)tempoMinutos : null;
            worksheet.Cells[linha, 5].Value = designados;

            // Aplica a cor de fundo
            AplicarCorDeFundo(worksheet, linha, corFundo);
        }

        private void AplicarCorDeFundo(ExcelWorksheet worksheet, int linha, string corHex)
        {
            worksheet.Cells[linha, 1, linha, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[linha, 1, linha, 5].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(corHex));
        }

        private string ObterCorFundoPorSessao(string tituloSessao)
        {
            return tituloSessao switch
            {
                SESSAO_TESOUROS_DA_PALAVRA_DE_DEUS => COR_TESOUROS_DA_PALAVRA_DE_DEUS,
                SESSAO_FACA_SEU_MELHOR_NO_MINISTERIO => COR_FACA_SEU_MELHOR_NO_MINISTERIO,
                SESSAO_NOSSA_VIDA_CRISTA => COR_NOSSA_VIDA_CRISTA,
                _ => COR_PADRAO // Cor padrão (branco) se não houver correspondência
            };
        }

        private void AplicarBordaSeparadora(ExcelWorksheet worksheet, int linha)
        {
            using (var range = worksheet.Cells[linha, 1, linha, 5])
            {
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Bottom.Color.SetColor(Color.Black);
            }
        }

        private void AplicarBordasTabela(ExcelWorksheet worksheet, int ultimaLinha)
        {
            using (var range = worksheet.Cells[1, 1, ultimaLinha, 5])
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Top.Color.SetColor(Color.Black);
                range.Style.Border.Bottom.Color.SetColor(Color.Black);
                range.Style.Border.Left.Color.SetColor(Color.Black);
                range.Style.Border.Right.Color.SetColor(Color.Black);
            }
        }

        private void SalvarArquivoExcel(ExcelPackage package, string caminhoArquivo)
        {
            FileInfo fileInfo = new FileInfo(caminhoArquivo);

            // Verifica se o diretório existe, se não, cria o diretório
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }

            package.SaveAs(fileInfo);
            Console.WriteLine($"Arquivo Excel criado com sucesso em: {fileInfo.FullName}");
        }
    }
}
