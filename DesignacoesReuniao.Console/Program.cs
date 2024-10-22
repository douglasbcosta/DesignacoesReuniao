using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Excel;
using DesignacoesReuniao.Infra.Pdf;
using DesignacoesReuniao.Infra.Scraper;
using DesignacoesReuniao.Infra.Word;

namespace DesignacoesReuniao;
class Program
{
    static void Main(string[] args)
    {
        Environment.SetEnvironmentVariable("ITEXT_BOUNCY_CASTLE_FACTORY_NAME", "bouncy-castle");
        PdfEditor pdfEditor = new PdfEditor();
        pdfEditor.EditPdfForm("C:\\Users\\Douglas\\Downloads\\Sem título.pdf", @"PartesEstudantes\PartesEstudantes.pdf");

        // URL base do site onde a programação está disponível (exemplo fictício)
        string baseUrl = "https://wol.jw.org/pt/wol/meetings/r5/lp-t";

        // Instancia o WebScraper
        WebScraper scraper = new WebScraper(baseUrl);

        // Exibe as opções para o usuário
        Console.WriteLine("O que deseja fazer?");
        Console.WriteLine("1. Exportar programação da reunião de um mês específico em excel para preenchimento das designações");
        Console.WriteLine("2. Exportar todas as programações de reuniões disponívels a partir do mês atual em excel para preenchimento das designações");
        Console.WriteLine("3. Com base em arquivo excel, preencher designados das reuniões de um mês específico");

        // Lê a escolha do usuário
        Console.Write("Digite o número da opção: ");
        int option = int.Parse(Console.ReadLine());

        if (option == 1)
        {
            // Opção 1: Exportar programação da reunião de um mês específico em excel para preenchimento das designações
            BuscarMesEspecifico(scraper);
        }
        else if (option == 2)
        {
            // Opção 2: Exportar todas as programações de reuniões disponívels a partir do mês atual em excel para preenchimento das designações
            BuscarAutomaticamente(scraper);
        }else if(option == 3)
        {
            // Opção 3: Com base em arquivo excel, preencher designados das reuniões de um mês específico
            PreencherDesignacoes(scraper);
        }



        // Pausa para que o console não feche imediatamente
        Console.WriteLine("\nPressione Enter para sair...");
        Console.ReadLine();
    }

    private static void PreencherDesignacoes(WebScraper scraper)
    {
        int month = ReceberMes();
        int year = ReceberAno();

        // Informe o caminho completo do arquivo
        Console.WriteLine("Informe o caminho completo do arquivo excel com as designações preenchidas:");
        string path = Console.ReadLine();

        ExcelImporter excelImporter = new ExcelImporter();
        var reunioesImportadas = excelImporter.ImportarReunioesDeExcel(path);

        // Chama o WebScraper para buscar as reuniões
        List<Reuniao> reunioesProgramacao = scraper.GetReunioes(year, month);

        // Preenche as designações de reunioesProgramacao com base em reunioesImportadas
        foreach (var reuniaoProgramada in reunioesProgramacao)
        {
            var reuniaoImportada = reunioesImportadas.FirstOrDefault(r => r.Semana == reuniaoProgramada.Semana);
            if (reuniaoImportada != null)
            {
                // Atualiza o presidente e orações
                reuniaoProgramada.Presidente = reuniaoImportada.Presidente;
                reuniaoProgramada.OracaoInicial = reuniaoImportada.OracaoInicial;
                reuniaoProgramada.OracaoFinal = reuniaoImportada.OracaoFinal;

                // Atualiza as sessões e partes
                foreach (var sessaoProgramada in reuniaoProgramada.Sessoes)
                {
                    var sessaoImportada = reuniaoImportada.Sessoes.FirstOrDefault(s => s.TituloSessao == sessaoProgramada.TituloSessao);
                    if (sessaoImportada != null)
                    {
                        foreach (var parteProgramada in sessaoProgramada.Partes)
                        {
                            var parteImportada = sessaoImportada.Partes.FirstOrDefault(p => p.TituloParte == parteProgramada.TituloParte);
                            if (parteImportada != null)
                            {
                                // Atualiza os designados e o tempo da parte
                                parteProgramada.Designados = parteImportada.Designados;
                                parteProgramada.TempoMinutos = parteImportada.TempoMinutos;
                            }
                        }
                    }
                }
            }
        }

        // Gera o arquivo Excel com as informações das reuniões
        ExportarReuniaoParaExcel(month, year, reunioesProgramacao, true);
        // Gera o arquivo Word com as informações das reuniões

        ExportarReuniaoParaWord(month, year, reunioesProgramacao, true);

        PreencherReuniaoWordModelo(month, year, reunioesProgramacao);

    }


    static void BuscarAutomaticamente(WebScraper scraper)
    {
        {
            // Começa a busca a partir do mês atual
            DateTime currentDate = DateTime.Now;
            int year = currentDate.Year;
            int month = currentDate.Month;

            List<Reuniao> reunioes = scraper.GetReunioes(year, month);

            // Continua buscando até que não seja mais encontrada programação
            while (reunioes.Any())
            {

                // Gera o arquivo Excel com as informações das reuniões
                ExportarReuniaoParaExcel(month, year, reunioes);
                // Gera o arquivo Word com as informações das reuniões
                ExportarReuniaoParaWord(month, year, reunioes);

                // Avança para o próximo mês
                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
                reunioes = scraper.GetReunioes(year, month);
            }

            Console.WriteLine("Busca automática finalizada. Não há mais programação disponível.");
        }
    }

    static void BuscarMesEspecifico(WebScraper scraper)
    {
        int month = ReceberMes();
        int year = ReceberAno();

        // Chama o WebScraper para buscar as reuniões
        List<Reuniao> reunioes = scraper.GetReunioes(year, month);

        // Gera o arquivo Excel com as informações das reuniões
        ExportarReuniaoParaExcel(month, year, reunioes);
        // Gera o arquivo Word com as informações das reuniões
        ExportarReuniaoParaWord(month, year, reunioes, true);

        PreencherReuniaoWordModelo(month, year, reunioes);
    }

    private static void PreencherReuniaoWordModelo(int month, int year, List<Reuniao> reunioes)
    {
        WordReplacer wordReplacer = new WordReplacer();
        string caminhoModelo = $@"S-140_T.docx";
        string caminhoArquivo = $@"ReunioesPreenchidas\\Reunioes_{month}_{year}.docx";
        wordReplacer.PreencherReunioesEmModelo(caminhoModelo, caminhoArquivo, reunioes);
    }

    private static int ReceberAno()
    {
        // Exibe as opções de ano
        int currentYear = DateTime.Now.Year;
        Console.WriteLine($"Selecione o ano:");
        Console.WriteLine($"1. {currentYear}");
        Console.WriteLine($"2. {currentYear + 1}");

        // Lê a escolha do ano
        Console.Write("Digite o número do ano: ");
        int yearOption = int.Parse(Console.ReadLine());
        int year = (yearOption == 1) ? currentYear : currentYear + 1;
        return year;
    }

    private static int ReceberMes()
    {
        // Exibe as opções de meses e anos para o usuário
        Console.WriteLine("Selecione o mês e o ano para buscar a programação das reuniões:");
        Console.WriteLine("1. Janeiro");
        Console.WriteLine("2. Fevereiro");
        Console.WriteLine("3. Março");
        Console.WriteLine("4. Abril");
        Console.WriteLine("5. Maio");
        Console.WriteLine("6. Junho");
        Console.WriteLine("7. Julho");
        Console.WriteLine("8. Agosto");
        Console.WriteLine("9. Setembro");
        Console.WriteLine("10. Outubro");
        Console.WriteLine("11. Novembro");
        Console.WriteLine("12. Dezembro");

        // Lê a escolha do mês
        Console.Write("Digite o número do mês: ");
        int month = int.Parse(Console.ReadLine());
        return month;
    }

    private static string ExportarReuniaoParaWord(int month, int year, List<Reuniao> reunioes, bool preenchido = false)
    {
        string pasta = preenchido ? "WordDesignacoesPreenchidas" : "WordModelos";
        WordExporter wordExporter = new WordExporter();
        string caminhoArquivo = $@"{pasta}\\Reunioes_{month}_{year}.docx";
        wordExporter.ExportarReunioesParaWord(reunioes, caminhoArquivo);
        return caminhoArquivo;
    }

    private static void ExportarReuniaoParaExcel(int month, int year, List<Reuniao> reunioes, bool preenchido = false)
    {
        string pasta = preenchido ? "ExcelDesignacoesPreenchidas" : "ExcelModelo";
        ExcelExporter excelExporter = new ExcelExporter();
        string caminhoArquivo = $@"{pasta}\Reunioes_{month}_{year}.xlsx";
        excelExporter.ExportarReunioesParaExcel(reunioes, caminhoArquivo);
    }
}