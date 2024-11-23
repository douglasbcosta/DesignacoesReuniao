namespace DesignacoesReuniao.Infra.Word
{
    public class Substituicao
    {
        public Substituicao(string valorOriginal, string valorSubstituicao, string sessao = "", string tema = "")
        {
            Sessao = sessao;
            ValorOriginal = valorOriginal ?? "";
            ValorSubstituicao = valorOriginal.Contains("Nome") ? valorSubstituicao : valorSubstituicao ?? "";
            Tema = tema;
        }
        public string Tema { get; set; }
        public string Sessao { get; set; }
        public string ValorOriginal { get; set; }
        public string ValorSubstituicao { get; set; }

        public static Dictionary<string, string> GetSubstituicoesPadrao()
        {
            Dictionary<string, string> substiticoes = new Dictionary<string, string>();

            substiticoes.Add("[", "");
            substiticoes.Add("]", "");
            substiticoes.Add("Chairman", "Presidente");            
            substiticoes.Add("NOME DA CONGREGAÇÃO", "ANDORINHA DA MATA");
            substiticoes.Add("Conselheiro da sala B", "");
            substiticoes.Add("Sala B", "");
            substiticoes.Add("Dirigente/leitor", "Dirigente");
            substiticoes.Add("TREASURES FROM GOD’S WORD", "TESOUROS DA PALAVRA DE DEUS");
            substiticoes.Add("APPLY YOURSELF TO THE FIELD MINISTRY", "FAÇA SEU MELHOR NO MINISTÉRIO");
            substiticoes.Add("LIVING AS CHRISTIANS", "NOSSA VIDA CRISTÃ");

            return substiticoes;
        }
    }
}
