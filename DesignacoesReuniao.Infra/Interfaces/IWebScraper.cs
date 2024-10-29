using DesignacoesReuniao.Domain.Models;

namespace DesignacoesReuniao.Infra.Interfaces
{
    public interface IWebScraper
    {
        List<Reuniao> GetReunioes(int year, int month);
    }
}
