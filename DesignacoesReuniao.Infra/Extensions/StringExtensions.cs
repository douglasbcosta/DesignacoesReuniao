using System.Text.RegularExpressions;

namespace DesignacoesReuniao.Infra.Extensions
{
    public static class StringExtensions
    {
        public static string FormatarTextoComPrimeiraLetraMaiuscula(this string texto)
        {
            if (string.IsNullOrEmpty(texto))
            {
                return texto;
            }
            return Regex.Replace(texto.ToLower(), @"\b\w", m => m.Value.ToUpper());
        }
        public static int ExtrairTempo(this string texto)
        {
            var match = Regex.Match(texto, @"\((\d+)\s*min\)");
            return match.Success ? int.Parse(match.Groups[1].Value) : 0;
        }
    }
}
