using DesignacoesReuniao.Domain.Models;

namespace DesignacoesReuniao.Infra.Repostories.Interface
{
    public interface IPessoaRepository
    {
        Pessoa? BuscarPessoa(string nome);
    }
}
