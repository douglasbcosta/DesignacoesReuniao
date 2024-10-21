namespace DesignacoesReuniao.Infra.Word
{
    public class Substituicao
    {
        public Substituicao(string valorOriginal, string valorSubstituicao, string sessao = "", string tema = "")
        {
            Sessao = sessao;
            ValorOriginal = valorOriginal ?? "";
            ValorSubstituicao = valorOriginal.Contains("Nome") ? FormatarTextoComPrimeiraLetraMaiuscula(valorSubstituicao ?? "") : valorSubstituicao ?? "";
            Tema = tema;
        }
        public string Tema { get; set; }
        public string Sessao { get; set; }
        public string ValorOriginal { get; set; }
        public string ValorSubstituicao { get; set; }

        private string FormatarTextoComPrimeiraLetraMaiuscula(string texto)
        {
            if (string.IsNullOrEmpty(texto))
            {
                return texto;
            }
            return System.Text.RegularExpressions.Regex.Replace(texto.ToLower(), @"\b\w", m => m.Value.ToUpper());
        }
    }
}
