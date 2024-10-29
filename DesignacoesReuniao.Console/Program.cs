using DesignacoesReuniao.CrossCutting.Injections;
using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Interfaces;
using Microsoft.Extensions.DependencyInjection;

namespace DesignacoesReuniao
{
    class Program
    {
        static void Main(string[] args)
        {
            // Configuração do ServiceCollection
            var services = new ServiceCollection();
            // Configurar o container de DI
            services.ConfigureDependences();
            services.AddTransient<Program>();

            // Construir o ServiceProvider
            var serviceProvider = services.BuildServiceProvider();
            // Resolver a classe Program e executar o método Run
            var program = serviceProvider.GetService<Program>();
            program.Run();
        }


        private readonly IWebScraper _scraper;
        private readonly IExcelExporter _excelExporter;
        private readonly IWordReplacer _wordReplacer;
        private readonly IPdfEditor _pdfEditor;
        private readonly IExcelImporter _excelImporter;

        // Construtor com injeção de dependências
        public Program(IWebScraper scraper, IExcelExporter excelExporter, IWordReplacer wordReplacer, IPdfEditor pdfEditor, IExcelImporter excelImporter)
        {
            _scraper = scraper;
            _excelExporter = excelExporter;
            _wordReplacer = wordReplacer;
            _pdfEditor = pdfEditor;
            _excelImporter = excelImporter;
        }

        public void Run()
        {
            

            // Exibe as opções para o usuário
            Console.WriteLine("O que deseja fazer?");
            Console.WriteLine("1. Exportar programação da reunião de um mês específico em excel para preenchimento das designações");
            Console.WriteLine("2. Com base em arquivo excel, preencher designados das reuniões de um mês específico");
            Console.WriteLine("3. Exportar todas as programações de reuniões disponíveis a partir do mês atual em excel para preenchimento das designações");

            // Lê a escolha do usuário
            Console.Write("Digite o número da opção: ");
            int option = int.Parse(Console.ReadLine());

            if (option == 1)
            {
                // Opção 1: Exportar programação da reunião de um mês específico em excel para preenchimento das designações
                BuscarMesEspecifico();
            }
            else if (option == 2)
            {
                // Opção 2: Com base em arquivo excel, preencher designados das reuniões de um mês específico
                PreencherDesignacoes();
            }
            else if (option == 3)
            {
                // Opção 3: Exportar todas as programações de reuniões disponíveis a partir do mês atual em excel para preenchimento das designações
                BuscarAutomaticamente();
            }

            // Pausa para que o console não feche imediatamente
            Console.WriteLine("\nPressione Enter para sair...");
            Console.ReadLine();
        }

        private void PreencherDesignacoes()
        {
            int month = ReceberMes();
            int year = ReceberAno();

            // Informe o caminho completo do arquivo
            Console.WriteLine("Informe o caminho completo do arquivo excel com as designações preenchidas:");
            string path = Console.ReadLine();

            var reunioesImportadas = _excelImporter.ImportarReunioesDeExcel(path);

            // Chama o WebScraper para buscar as reuniões
            List<Reuniao> reunioesProgramacao = _scraper.GetReunioes(year, month);

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

            _excelExporter.ExportarReunioesParaExcel(month, year, reunioesProgramacao);
            _wordReplacer.PreencherReunioesEmModelo(month, year, reunioesProgramacao);
            _pdfEditor.EditPdfForm(month, year, reunioesProgramacao);
        }

        private void BuscarAutomaticamente()
        {
            // Começa a busca a partir do mês atual
            DateTime currentDate = DateTime.Now;
            int year = currentDate.Year;
            int month = currentDate.Month;

            List<Reuniao> reunioes = _scraper.GetReunioes(year, month);

            // Continua buscando até que não seja mais encontrada programação
            while (reunioes.Any())
            {
                ExportarReuniaoParaExcel(month, year, reunioes);
                _wordReplacer.PreencherReunioesEmModelo(month, year, reunioes);
                _pdfEditor.EditPdfForm(month,year, reunioes);

                // Avança para o próximo mês
                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
                reunioes = _scraper.GetReunioes(year, month);
            }

            Console.WriteLine("Busca automática finalizada. Não há mais programação disponível.");
        }

        private void BuscarMesEspecifico()
        {
            int month = ReceberMes();
            int year = ReceberAno();

            // Chama o WebScraper para buscar as reuniões
            List<Reuniao> reunioes = _scraper.GetReunioes(year, month);

            ExportarReuniaoParaExcel(month, year, reunioes);
            _wordReplacer.PreencherReunioesEmModelo(month, year, reunioes);
            _pdfEditor.EditPdfForm(month, year, reunioes);
        }


        private int ReceberAno()
        {
            int currentYear = DateTime.Now.Year;
            Console.WriteLine($"Selecione o ano:");
            Console.WriteLine($"1. {currentYear}");
            Console.WriteLine($"2. {currentYear + 1}");

            int yearOption = int.Parse(Console.ReadLine());
            return (yearOption == 1) ? currentYear : currentYear + 1;
        }

        private int ReceberMes()
        {
            Console.WriteLine("Selecione o mês:");
            for (int i = 1; i <= 12; i++)
            {
                Console.WriteLine($"{i}. {new DateTime(1, i, 1).ToString("MMMM")}");
            }
            return int.Parse(Console.ReadLine());
        }

        private void ExportarReuniaoParaExcel(int month, int year, List<Reuniao> reunioes)
        {
            _excelExporter.ExportarReunioesParaExcel(month, year, reunioes);
        }

    }
}