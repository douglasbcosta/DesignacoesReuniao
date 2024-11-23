using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Interfaces;
using DesignacoesReuniao.Infra.Word;
using iText.Forms;
using iText.Kernel.Pdf;
using System.Text;

namespace DesignacoesReuniao.Infra.Pdf
{
    public class PdfEditor : IPdfEditor
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

        string modelo = "S-89-T.pdf";

        public string EditPdfForm(int month, int year, List<Reuniao> reunioes)
        {
            Environment.SetEnvironmentVariable("ITEXT_BOUNCY_CASTLE_FACTORY_NAME", "bouncy-castle");

            string caminhoDestinho = $"PartesEstudantes/{year}/{month}/PartesEstudantes_{year}_{month}.pdf";

            var fileInfo = new FileInfo(caminhoDestinho);

            // Verifica se o diretório existe, se não, cria o diretório
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            caminhoDestinho = fileInfo.FullName;

            // Abre o PDF existente
            using (var reader = new PdfReader(modelo))
            using (var writer = new PdfWriter(caminhoDestinho))
            using (var pdfDoc = new PdfDocument(reader, writer))
            {
                // Obtém o formulário do PDF
                var form = PdfAcroForm.GetAcroForm(pdfDoc, true);

                // Obtém todos os campos do formulário
                var fields = form.GetAllFormFields();
                int indiceField = 1;
                int indicePagina = 0;

                var substituicoes = new List<Substituicao>();

                foreach (var reuniao in reunioes)
                {
                    var dataReuniao = GetProximaTerca(reuniao.InicioSemana);

                    var partes = reuniao.Sessoes
                        .SelectMany(s => s.Partes.Where(p => PartesEstudantes.Any(pe => p.TituloParte.Contains(pe))))
                        .ToList();

                    foreach (var parte in partes)
                    {
                        if (indiceField == 29)
                        {
                            indiceField = 1;
                            indicePagina++;
                        }

                        // Adiciona substituições para Estudante, Ajudante, Data, Título e Salão Principal
                        if (parte.Designado != null && !string.IsNullOrEmpty(parte.Designado.NomeCompleto))
                        {
                            AdicionarSubstituicoes(substituicoes, parte, dataReuniao, indicePagina, ref indiceField);
                        }
                    }
                }

                // Aplica as substituições nos campos do PDF
                foreach (var substituicao in substituicoes)
                {
                    fields[substituicao.ValorOriginal].SetValue(substituicao.ValorSubstituicao);
                }
            }
            return caminhoDestinho;
        }

        private static DateOnly GetProximaTerca(DateOnly data)
        {
            while (data.DayOfWeek != DayOfWeek.Tuesday)
            {
                data = data.AddDays(1);
            }
            return data;
        }

        private void AdicionarSubstituicoes(List<Substituicao> substituicoes, Parte parte, DateOnly dataReuniao, int indicePagina, ref int indiceField)
        {
            // Estudante

            var nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
            substituicoes.Add(new Substituicao(nomeField, parte.Designado !=null ? parte.Designado.NomeCompleto: ""));
            indiceField++;

            // Ajudante
            var ajudante = parte.Ajudante != null ? parte.Ajudante.NomeCompleto : "";
            nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
            substituicoes.Add(new Substituicao(nomeField, ajudante));
            indiceField++;

            // Data
            nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
            substituicoes.Add(new Substituicao(nomeField, dataReuniao.ToString()));
            indiceField++;

            // Título parte
            nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
            substituicoes.Add(new Substituicao(nomeField, $"{parte.IndiceParte}. {parte.TituloParte}"));
            indiceField++;

            // Salão Principal
            nomeField = GetNomeField(indicePagina, TipoInput.Checkbox, indiceField);
            substituicoes.Add(new Substituicao(nomeField, "Yes"));
            indiceField++;

            // Checkboxes vazios
            for (int i = 0; i < 2; i++)
            {
                nomeField = GetNomeField(indicePagina, TipoInput.Checkbox, indiceField);
                substituicoes.Add(new Substituicao(nomeField, ""));
                indiceField++;
            }
        }

        private string GetNomeField(int indiceReuniao, TipoInput tipoInput, int indiceField)
        {
            var nomeField = new StringBuilder();
            if (indiceReuniao > 0)
            {
                nomeField.Append($"{indiceReuniao}.");
            }
            nomeField.Append($"900_{indiceField}_");
            nomeField.Append(tipoInput == TipoInput.Text ? "Text_SanSerif" : "CheckBox");
            return nomeField.ToString();
        }
    }
}