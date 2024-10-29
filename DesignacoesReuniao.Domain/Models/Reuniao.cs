namespace DesignacoesReuniao.Domain.Models
{
    public class Reuniao
    {
        public string Semana { get; set; }
        public DateOnly InicioSemana { get; set; }
        public string LeituraDaSemana { get; set; }
        public string Presidente { get; set; } 
        public string OracaoInicial { get; set; } 
        public string OracaoFinal { get; set; } 
        public List<string> Canticos { get; set; } = new List<string>();
        public List<Sessao> Sessoes { get; set; }
        public Reuniao() { 
            Sessoes = new List<Sessao>(); 
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
