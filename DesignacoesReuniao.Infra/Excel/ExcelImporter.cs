using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Interfaces;
using DesignacoesReuniao.Infra.Repostories.Interface;
using OfficeOpenXml;

namespace DesignacoesReuniao.Infra.Excel
{
    
    public class ExcelImporter : IExcelImporter
    {
        private readonly IPessoaRepository _pessoasRepository;

        public ExcelImporter(IPessoaRepository pessoasRepository)
        {
            _pessoasRepository = pessoasRepository;
        }

        public List<Reuniao> ImportarReunioesDeExcel(string caminhoArquivo)
        {
            // Configura a licença do EPPlus (obrigatório a partir da versão 5)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var reunioes = new List<Reuniao>();

            using (var package = new ExcelPackage(new FileInfo(caminhoArquivo)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assume que a planilha de reuniões é a primeira

                int linhaAtual = 2; // Começa na linha 2, pois a linha 1 é o cabeçalho
                Reuniao reuniaoAtual = null;

                while (worksheet.Cells[linhaAtual, 1].Value != null)
                {
                    string semana = worksheet.Cells[linhaAtual, 1].Text;
                    string sessao = worksheet.Cells[linhaAtual, 2].Text;
                    string[] textoParte = worksheet.Cells[linhaAtual, 3].Text.Split('.');
                    int indiceParte = textoParte.Count() > 1 ? int.Parse(textoParte[0]) : 0;
                    string parte = textoParte.Count() > 1 ? textoParte[1] : textoParte[0];
                    string designado = worksheet.Cells[linhaAtual, 5].Text;
                    string ajudante = worksheet.Cells[linhaAtual, 6].Text;
                    int tempoMinutos = worksheet.Cells[linhaAtual, 4].GetValue<int>();

                    // Verifica se é uma nova reunião (baseado na semana)
                    if (reuniaoAtual == null || reuniaoAtual.Semana != semana)
                    {
                        if (reuniaoAtual != null)
                        {
                            reunioes.Add(reuniaoAtual);
                        }

                        reuniaoAtual = new Reuniao
                        {
                            Semana = semana,
                            Sessoes = new List<Sessao>()
                        };
                    }

                    // Verifica a sessão e preenche as partes correspondentes
                    var sessaoAtual = reuniaoAtual.Sessoes.Find(s => s.TituloSessao == sessao);
                    if (sessaoAtual == null)
                    {
                        sessaoAtual = new Sessao(sessao);
                        reuniaoAtual.Sessoes.Add(sessaoAtual);
                    }

                    // Verifica se é Presidente, Oração Inicial ou Oração Final
                    if (parte == "Presidente")
                    {
                        var pessoa = _pessoasRepository.BuscarPessoa(designado);
                        if (pessoa != null)
                            reuniaoAtual.Presidente = pessoa;
                        linhaAtual++;
                        continue;
                    }
                    else if (parte == "Oração Inicial")
                    {
                        var pessoa = _pessoasRepository.BuscarPessoa(designado);
                        if (pessoa != null)
                            reuniaoAtual.OracaoInicial = pessoa;
                        linhaAtual++;
                        continue;

                    }
                    else if (parte == "Oração Final")
                    {
                        var pessoa = _pessoasRepository.BuscarPessoa(designado);
                        if (pessoa != null)
                            reuniaoAtual.OracaoFinal = pessoa;
                        linhaAtual++;
                        continue;
                    }

                    Parte parteAtual = new Parte(indiceParte, parte, tempoMinutos);
                    if (!string.IsNullOrEmpty(designado))
                    {
                        var pessoa = _pessoasRepository.BuscarPessoa(designado);
                        if (pessoa != null)
                            parteAtual.AdicionarDesignado(pessoa);
                    }
                    if (!string.IsNullOrEmpty(ajudante))
                    {
                        var pessoa = _pessoasRepository.BuscarPessoa(ajudante);
                        if (pessoa != null)
                            parteAtual.AdicionarDesignado(pessoa);
                    }

                    // Adiciona a parte à sessão
                    sessaoAtual.AdicionarParte(parteAtual);



                    linhaAtual++;
                }

                // Adiciona a última reunião
                if (reuniaoAtual != null)
                {
                    reunioes.Add(reuniaoAtual);
                }
            }

            return reunioes;
        }

        public List<ReuniaoDesignacoes> ImportarReunioesExcel(string caminhoarquivo, int mes)
        {
            var listaReunioes = ImportarDesignacoesPorMes(mes, caminhoarquivo);
            return listaReunioes;
        }

        public List<ReuniaoDesignacoes> ImportarDesignacoesPorMes(int mes, string caminhoarquivo)
        {
            var designacoes = new List<ReuniaoDesignacoes>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(caminhoarquivo)))
            {
                // Processar worksheets para Anciãos/Servos e Estudantes
                var worksheets = package.Workbook.Worksheets;
                foreach (var worksheet in worksheets)
                {
                    int startRow = 0;
                    int startColumn = 0;

                    // Identificar o início da tabela com base no nome da worksheet
                    if (worksheet.Name.Contains("Anciãos, Servos", StringComparison.OrdinalIgnoreCase))
                    {
                        startRow = 5; // Começa na linha 5
                        startColumn = 2; // Começa na coluna B
                    }
                    else if (worksheet.Name.Contains("Estudantes", StringComparison.OrdinalIgnoreCase))
                    {
                        startRow = 4; // Começa na linha 4
                        startColumn = 3; // Começa na coluna C
                    }

                    if (startRow == 0 || startColumn == 0)
                        continue; // Pular se não encontrar a tabela

                    var monthColumn = FindMonthColumn(worksheet, startRow, startColumn, mes);
                    if (monthColumn == 0)
                        continue; // Pular se não encontrar o mês

                    // Identificar a coluna final do mês
                    var monthColumnEnd = FindMonthColumnEnd(worksheet, startRow, startColumn, mes);
                    if (monthColumnEnd == 0)
                        continue; // Pular se não encontrar o mês

                    // Processar as partes das semanas
                    int semanaCounter = 1; // Contador para as semanas
                    int colCount = worksheet.Dimension.Columns;

                    for (int col = monthColumn; col <= monthColumnEnd; col++) // Itera sobre as colunas do mês
                    {
                        // Para cada linha da planilha, verificar as partes preenchidas
                        int rowCount = worksheet.Dimension.Rows;
                        for (int row = startRow + 1; row <= rowCount; row++)
                        {
                            var nomeCell = worksheet.Cells[row, startColumn]; // Primeira coluna contém os nomes
                            if (nomeCell == null || string.IsNullOrWhiteSpace(nomeCell.Text))
                                continue;

                            var parteCell = worksheet.Cells[row, col];
                            if (parteCell == null || string.IsNullOrWhiteSpace(parteCell.Text))
                                continue;

                            var designacao = new ReuniaoDesignacoes
                            {
                                Semana = semanaCounter, // Data não necessária
                                Designado = _pessoasRepository.BuscarPessoa(nomeCell.Text),
                                Tipo = worksheet.Name, // Identificar pelo nome da worksheet
                                Parte = parteCell.Text, // Parte designada 
                                Cor = string.IsNullOrEmpty(parteCell.Style.Fill.BackgroundColor.Rgb) ? parteCell.Style.Fill.BackgroundColor.Theme.ToString() : parteCell.Style.Fill.BackgroundColor.Rgb // Cor da designação
                            };

                            designacoes.Add(designacao);
                        }

                        semanaCounter++; // Incrementa o contador da semana
                    }
                }
            }

            return designacoes
                .OrderBy(d => d.Tipo) // Ordenar por Tipo (Anciãos, Servos, Estudantes)
                .ThenBy(d => d.Designado.NomeResumido) // Ordenar por Nome
                .ToList();
        }


        private (int StartRow, int StartColumn) FindStartRowAndColumn(ExcelWorksheet worksheet)
        {
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // Valores a serem ignorados
            var ignoredValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "NOSSA VIDA E MINISTERIO CRISTÃO",
                "PARTES - ANDORINHA"
            };

            // Buscar na diagonal
            for (int i = 1; i <= Math.Min(rowCount, colCount); i++)
            {
                var cell = worksheet.Cells[i, i];
                if (!string.IsNullOrWhiteSpace(cell.Text) && !ignoredValues.Contains(cell.Text))
                {
                    // Expandir a partir da célula diagonal encontrada
                    int startRow = i;
                    int startColumn = i;

                    // Procurar a primeira linha com conteúdo acima
                    for (int r = i; r > 0; r--)
                    {
                        if (!string.IsNullOrWhiteSpace(worksheet.Cells[r, i].Text) && !ignoredValues.Contains(worksheet.Cells[r, i].Text))
                        {
                            startRow = r;
                        }
                        else
                        {
                            break;
                        }
                    }

                    // Procurar a primeira coluna com conteúdo à esquerda
                    for (int c = i; c > 0; c--)
                    {
                        if (!string.IsNullOrWhiteSpace(worksheet.Cells[i, c].Text) && !ignoredValues.Contains(worksheet.Cells[i, c].Text))
                        {
                            startColumn = c;
                        }
                        else
                        {
                            break;
                        }
                    }

                    return (startRow, startColumn);
                }
            }

            return (0, 0); // Retorna 0, 0 se não encontrar nada
        }



        private int FindMonthColumn(ExcelWorksheet worksheet, int startRow, int startColumn, int mes)
        {
            int colCount = worksheet.Dimension.Columns;
            string monthName = new DateTime(DateTime.Now.Year, mes, 1).ToString("MMMM", new System.Globalization.CultureInfo("pt-BR"));

            for (int col = startColumn; col <= colCount; col++) // Começa na coluna detectada
            {
                var cellValue = worksheet.Cells[startRow, col].Text;
                if (cellValue.Contains(monthName, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }
            return 0;
        }

        private int FindMonthColumnEnd(ExcelWorksheet worksheet, int startRow, int startColumn, int mes)
        {
            int colCount = worksheet.Dimension.Columns;
            string monthName = new DateTime(DateTime.Now.Year, mes, 1).ToString("MMMM", new System.Globalization.CultureInfo("pt-BR"));
            string nextMonthName = new DateTime(DateTime.Now.Year, mes + 1, 1).ToString("MMMM", new System.Globalization.CultureInfo("pt-BR"));

            int monthColumnStart = 0;
            int monthColumnEnd = 0;

            // Buscar a coluna do mês atual
            for (int col = startColumn; col <= colCount; col++)
            {
                var cellValue = worksheet.Cells[startRow, col].Text;
                if (cellValue.Contains(monthName, StringComparison.OrdinalIgnoreCase))
                {
                    monthColumnStart = col;
                    break;
                }
            }

            // Se não encontrar o mês atual, retorna 0
            if (monthColumnStart == 0)
                return 0;

            // Buscar a coluna do próximo mês
            for (int col = monthColumnStart + 1; col <= colCount; col++)
            {
                var cellValue = worksheet.Cells[startRow, col].Text;
                if (cellValue.Contains(nextMonthName, StringComparison.OrdinalIgnoreCase))
                {
                    monthColumnEnd = col - 1; // A coluna antes do próximo mês
                    break;
                }
            }

            // Caso não encontre o próximo mês, limite a 5 colunas a partir do mês atual
            if (monthColumnEnd == 0)
                monthColumnEnd = Math.Min(monthColumnStart + 5, colCount);

            return monthColumnEnd;
        }
    }
}