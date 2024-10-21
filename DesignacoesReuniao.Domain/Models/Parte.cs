namespace DesignacoesReuniao.Domain.Models
{
    public class Parte
    {
        public Parte(int indiceParte, string tituloParte, int tempoMinutos)
        {
            IndiceParte = indiceParte;
            TituloParte = tituloParte;
            TempoMinutos = tempoMinutos;
            Designados = new List<string>();
        }

        public int IndiceParte { get; set; }
        public string TituloParte { get; set; }        
        public int TempoMinutos { get; set; }
        public List<string> Designados { get; set; } 

        public void AdicionarDesignado(string designado)
        {
            if (Designados.Count < 2)
            {
                Designados.Add(designado);
            }
            else
            {
                Console.WriteLine("Não é possível adicionar mais de 2 designados para esta parte.");
            }
        }
    }
}
