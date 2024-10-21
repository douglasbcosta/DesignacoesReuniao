namespace DesignacoesReuniao.Domain.Models
{
    public class Sessao
    {
        public string TituloSessao { get; set; }
        public List<Parte> Partes { get; set; }
        public Sessao(string tituloSessao) { 
            TituloSessao = tituloSessao; 
            Partes = new List<Parte>(); 
        }
        public void AdicionarParte(Parte parte) {
            Partes.Add(parte); 
        }
    }
}
