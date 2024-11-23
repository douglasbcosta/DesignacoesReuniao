using DesignacoesReuniao.Shared.Extensions;

namespace DesignacoesReuniao.Domain.Models
{
    public class Pessoa
    {
        public Pessoa()
        {
               
        }
        public Pessoa(string nomeCompleto)
        {
            NomeCompleto = nomeCompleto.FormatarTextoComPrimeiraLetraMaiuscula();
        }

        public Pessoa(string nomeCompleto, string nomeResumido)
        {
            NomeCompleto = nomeCompleto.FormatarTextoComPrimeiraLetraMaiuscula();
            NomeResumido = nomeResumido.FormatarTextoComPrimeiraLetraMaiuscula();
        }

        public string NomeCompleto { get; set; }
        public string NomeResumido { get; set; }

        public override string ToString()
        {
            return !string.IsNullOrEmpty(NomeResumido) ? NomeResumido : string.Empty;
        }
    }
}
