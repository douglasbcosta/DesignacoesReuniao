namespace DesignacoesReuniao.Domain.Models
{
    public class Reuniao
    {
        public string Semana { get; set; }
        public DateOnly InicioSemana { get; set; }
        public string LeituraDaSemana { get; set; }
        public Pessoa Presidente { get; set; } 
        public Pessoa OracaoInicial { get; set; } 
        public Pessoa OracaoFinal { get; set; } 
        public List<string> Canticos { get; set; } = new List<string>();
        public List<Sessao> Sessoes { get; set; }
        public Reuniao() { 
            Sessoes = new List<Sessao>(); 
            Presidente = new Pessoa();
            OracaoInicial = new Pessoa();
            OracaoFinal = new Pessoa();
        }
        public void AdicionarSessao(Sessao sessao) { 
            Sessoes.Add(sessao); 
        }
        public void AdicionarCantico(string cantico)
        {
            Canticos.Add(cantico);
        }

        public static string[] GetSessoesReunioes()
        {
            return new string[] { "TESOUROS DA PALAVRA DE DEUS", "FAÇA SEU MELHOR NO MINISTÉRIO", "NOSSA VIDA CRISTÃ" };
        }
    }
}
