using DesignacoesReuniao.Domain.Models;

namespace DesignacoesReuniao.Infra.Interfaces
{
    public interface IPdfEditor
    {
        string EditPdfForm(int month, int year, List<Reuniao> reunioes);
    }
}
