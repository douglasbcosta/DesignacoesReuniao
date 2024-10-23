using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Extensions;
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
                int indiceField = 1;
                int indicePagina = 0;

                List<Substituicao> substituicoes = new List<Substituicao>();
                foreach (var reuniao in reunioes)
                {
                    string nomeField = "";
                    DateOnly dataReuniao = reuniao.InicioSemana;
                    while (dataReuniao.DayOfWeek != DayOfWeek.Tuesday)
                    {
                        dataReuniao = dataReuniao.AddDays(1);
                    }

                    var partes = reuniao.Sessoes
                        .SelectMany(s => s.Partes.Where(p => p.TituloParte.Contains(partesEstudantes[0])
                        || p.TituloParte.Contains(partesEstudantes[1])
                        || p.TituloParte.Contains(partesEstudantes[2])
                        || p.TituloParte.Contains(partesEstudantes[3])
                        || p.TituloParte.Contains(partesEstudantes[4])
                        || p.TituloParte.Contains(partesEstudantes[5])))
                        .ToList();
                    foreach (var parte in partes) 
                    {
                        if (indiceField == 29)
                        {
                            indiceField = 1;
                            indicePagina++;
                        }
                        //Estudante
                        nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
                        substituicoes.Add(new Substituicao(nomeField, parte.Designados.Any() ? parte.Designados[0].FormatarTextoComPrimeiraLetraMaiuscula() :""));
                        indiceField++;
                        //Ajudante
                        string ajudante = parte.Designados.Count > 1 ? parte.Designados[1].FormatarTextoComPrimeiraLetraMaiuscula() : "";
                        nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
                        substituicoes.Add(new Substituicao(nomeField, ajudante));
                        indiceField++;
                        //Data                        
                        nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
                        substituicoes.Add(new Substituicao(nomeField, dataReuniao.ToString()));
                        indiceField++;
                        //Título parte
                        nomeField = GetNomeField(indicePagina, TipoInput.Text, indiceField);
                        substituicoes.Add(new Substituicao(nomeField,$"{parte.IndiceParte}. {parte.TituloParte}"));
                        indiceField++;
                        //Salao Principal
                        nomeField = GetNomeField(indicePagina, TipoInput.Checkbox, indiceField);
                        substituicoes.Add(new Substituicao(nomeField, "Yes"));
                        indiceField++;
                        nomeField = GetNomeField(indicePagina, TipoInput.Checkbox, indiceField);
                        substituicoes.Add(new Substituicao(nomeField, ""));
                        indiceField++;
                        nomeField = GetNomeField(indicePagina, TipoInput.Checkbox, indiceField);
                        substituicoes.Add(new Substituicao(nomeField, ""));
                        indiceField++;
                        
                    }
                }

                foreach(var substituicao in substituicoes)
                {
                    fields[substituicao.ValorOriginal].SetValue(substituicao.ValorSubstituicao);
                }

            }
        }

        private string GetNomeField(int indiceReuniao, TipoInput tipoInput, int indiceField)
        {
            string nomeField = indiceReuniao > 0 ? indiceReuniao.ToString()+"." : "";
            nomeField += $"900_{indiceField}_";
            nomeField += tipoInput == TipoInput.Text ? "Text_SanSerif" : "CheckBox";
            return nomeField;
        }
    }
}
