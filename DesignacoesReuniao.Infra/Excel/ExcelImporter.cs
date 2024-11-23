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

                    Parte parteAtual = new Parte(indiceParte,parte, tempoMinutos);
                    if (!string.IsNullOrEmpty(designado))
                    {
                        var pessoa = _pessoasRepository.BuscarPessoa(designado);
                        if(pessoa != null)
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
    }
}