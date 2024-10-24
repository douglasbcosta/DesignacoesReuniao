using DesignacoesReuniao.Domain.Models;

namespace DesignacoesReuniao.Infra.Interfaces
{
    public interface IWordReplacer
    {
        string PreencherReunioesEmModelo(int month, int year, List<Reuniao> reunioes);
    }
}
