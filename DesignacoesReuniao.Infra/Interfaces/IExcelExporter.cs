using DesignacoesReuniao.Domain.Models;

namespace DesignacoesReuniao.Infra.Interfaces
{
    public interface IExcelExporter
    {
        string BuscarArquivo(int month, int year);
        string ExportarReunioesParaExcel(int month, int year, List<Reuniao> reunioes);
    }
}
