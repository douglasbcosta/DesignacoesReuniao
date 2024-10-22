using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Word;
using iText.Forms;
using iText.Kernel.Pdf;

namespace DesignacoesReuniao.Infra.Pdf
{
    public class PdfEditor
    {

        public void EditPdfForm(string inputPdfPath, string outputPdfPath, List<Reuniao> reunioes)
        {
            string[] partesEstudantes = ["Leitura da Bíblia", "Iniciando conversas", "Cultivando o interesse", "Fazendo discípulos", "Explicando suas crenças", "Discurso"];
            FileInfo fileInfo = new FileInfo(outputPdfPath);

            // Verifica se o diretório existe, se não, cria o diretório
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            // Abre o PDF existente
            using (var reader = new PdfReader(inputPdfPath))
            using (var writer = new PdfWriter(outputPdfPath))
            using (var pdfDoc = new PdfDocument(reader, writer))
            {
                // Obtém o formulário do PDF
                // Obtém o formulário do PDF
                var form = PdfAcroForm.GetAcroForm(pdfDoc, true);

                // Obtém todos os campos do formulário
                var fields = form.GetAllFormFields();
                int qtdeInputs = 4;
                int qtdeCheckbox = 3;
                int indiceField = 0;

                List<Substituicao> substituicoes = new List<Substituicao>();
                for (int i = 0; i < reunioes.Count; i++)
                {
                    string nomeField = "";
                    var partes = reunioes[i].Sessoes
                        .SelectMany(s => s.Partes.Where(p => partesEstudantes.Contains(p.TituloParte)))
                        .ToList();

                    //Estudante
                    nomeField = GetNomeField(i,TipoInput.Text, indiceField);
                    substituicoes.Add(nomeField, partes)
                    //Ajudante

                    //Data

                    //Título parte
                }


                fields["nomeCampo1"].SetValue("Novo Valor 1");
                fields["nomeCampo2"].SetValue("Novo Valor 2");

            }
        }

        private string GetNomeField(int indiceReuniao, TipoInput tipoInput, int indiceField)
        {
            string nomeField = indiceReuniao > 0 ? indiceReuniao.ToString() : "";
            nomeField += $"900_{indiceField}_";
            nomeField += tipoInput == TipoInput.Text ? "Text_SanSerif" : "CheckBox";
            return nomeField;
        }
    }
}
