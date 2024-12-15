
namespace DesignacoesReuniao.Domain.Models
{
    public class Parte
    {
        private static readonly string[] PartesEstudantes =
        {
            "Leitura da Bíblia",
            "Iniciando conversas",
            "Cultivando o interesse",
            "Fazendo discípulos",
            "Explicando suas crenças",
            "Discurso"
        };
        public Parte(int indiceParte, string tituloParte, int tempoMinutos)
        {
            IndiceParte = indiceParte;
            TituloParte = tituloParte;
            TempoMinutos = tempoMinutos;
        }

        public int IndiceParte { get; set; }
        public string TituloParte { get; set; }        
        public int TempoMinutos { get; set; }
        public Pessoa Designado { get; set; }
        public Pessoa Ajudante { get; set; }

        public bool ContemDesignado()
        {
            return !string.IsNullOrEmpty(Designado?.NomeCompleto);
        }
        public bool ContemAjudante()
        {
            return !string.IsNullOrEmpty(Ajudante?.NomeCompleto);
        }

        public static bool ParteDeEstudante(string tituloParte)
        {
            return PartesEstudantes.Any(pe => tituloParte.Contains(pe));
        }


        public void AdicionarDesignado(string designado)
        {
            if (Designado != null)
            {
                Designado = new Pessoa(designado.Trim());
            }
            else if (Ajudante != null)
            {
                Ajudante = new Pessoa(designado.Trim());
            }
            else
            {
                Console.WriteLine("Não é possível adicionar mais de 2 designados para esta parte.");
            }
        }
        public string ObterNomesDesignadoEAjudante()
        {
            string nomes = string.Empty;
            if (!string.IsNullOrEmpty(Designado?.NomeResumido))
            {
                nomes += Designado.NomeResumido;
            }
            if (!string.IsNullOrEmpty(Ajudante?.NomeResumido))
            {
                nomes += $"/ {Ajudante.NomeResumido}";
            }
            return nomes;
        }

        public void AdicionarDesignado(Pessoa designado)
        {
            if (Designado == null)
            {
                Designado = designado;
            }
            else if (Ajudante == null)
            {
                Ajudante = designado;
            }
            else
            {
                Console.WriteLine("Não é possível adicionar mais de 2 designados para esta parte.");
            }
        }
    }
}
