using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Excel;

namespace DesignacoesReuniao.Infra.Interfaces
{
    public interface IExcelImporter
    {
        List<Reuniao> ImportarReunioesDeExcel(string path);
        List<ReuniaoDesignacoes> ImportarReunioesExcel(string caminhoarquivo, int mes);
    }
}
