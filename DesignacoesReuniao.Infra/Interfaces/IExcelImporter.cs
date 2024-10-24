using DesignacoesReuniao.Domain.Models;

namespace DesignacoesReuniao.Infra.Interfaces
{
    public interface IExcelImporter
    {
        List<Reuniao> ImportarReunioesDeExcel(string path);
    }
}
