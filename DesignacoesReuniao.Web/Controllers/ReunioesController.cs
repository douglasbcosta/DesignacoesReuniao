using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace DesignacoesReuniao.Web.Controllers
{
    public class ReunioesController : Controller
    {
        private readonly IWebScraper _scraper;
        private readonly IExcelExporter _excelExporter;
        private readonly IWordReplacer _wordReplacer;
        private readonly IPdfEditor _pdfEditor;
        private readonly IExcelImporter _excelImporter;

        public ReunioesController(IWebScraper scraper, IExcelExporter excelExporter, IWordReplacer wordReplacer, IPdfEditor pdfEditor, IExcelImporter excelImporter)
        {
            _scraper = scraper;
            _excelExporter = excelExporter;
            _wordReplacer = wordReplacer;
            _pdfEditor = pdfEditor;
            _excelImporter = excelImporter;
        }

        // Exibe a página inicial com as opções
        public IActionResult Index()
        {
            return View();
        }

        // Exportar programação da reunião de um mês específico em excel, word e pdf
        [HttpGet]
        public IActionResult ExportarMesEspecifico(int month, int year)
        {
            var caminhoExcel = _excelExporter.BuscarArquivo(month, year);
            if (!string.IsNullOrEmpty(caminhoExcel))
            {
                return Ok(new
                {
                    excelPath = caminhoExcel
                });
            }

            var reunioes = _scraper.GetReunioes(year, month);
            caminhoExcel = _excelExporter.ExportarReunioesParaExcel(month, year, reunioes);

            // Retorna os caminhos dos arquivos para o front-end habilitar os botões de download
            return Ok(new
            {
                excelPath = caminhoExcel
            });
        }

        // Preencher designados das reuniões de um mês específico com base em arquivo excel
        [HttpPost]
        public IActionResult PreencherDesignacoes(int month, int year, string tipoExcel, IFormFile excelFile)
        {
            if (tipoExcel == "template")
            {
                return PreencherDesignacoesModelo1(month, year, excelFile);
            }
            else
            {
                return PreencherDesignacoesModelo2(month, year, excelFile);
            }
        }

        private IActionResult PreencherDesignacoesModelo1(int month, int year, IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
                return BadRequest("Arquivo Excel não fornecido.");

            // Definir o caminho onde o arquivo será salvo no servidor
            var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "uploads");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            // Gerar um nome de arquivo único para evitar conflitos
            var filePath = Path.Combine(uploadsFolder, $"{year}_{month}_{excelFile.FileName}");

            // Salvar o arquivo no servidor
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                excelFile.CopyTo(stream);
            }

            // Agora que o arquivo foi salvo, você pode passar o caminho completo para o método de importação
            var reunioesImportadas = _excelImporter.ImportarReunioesDeExcel(filePath);
            var reunioesProgramacao = _scraper.GetReunioes(year, month);

            reunioesProgramacao = Reuniao.PreencherReunioes(reunioesProgramacao, reunioesImportadas);

            

            var caminhoWord = _wordReplacer.PreencherReunioesEmModelo(month, year, reunioesProgramacao);
            var caminhoPdf = PreencherPartesEstudantes(month, year, reunioesProgramacao);

            // Retorna os caminhos dos arquivos para o front-end habilitar os botões de download
            return Ok(new
            {
                wordPath = caminhoWord,
                pdfPath = caminhoPdf
            });
        }

        private IActionResult PreencherDesignacoesModelo2(int month, int year, IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
                return BadRequest("Arquivo Excel não fornecido.");

            // Definir o caminho onde o arquivo será salvo no servidor
            var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "uploads");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            // Gerar um nome de arquivo único para evitar conflitos
            var filePath = Path.Combine(uploadsFolder, $"{year}_{month}_{excelFile.FileName}");

            // Salvar o arquivo no servidor
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                excelFile.CopyTo(stream);
            }

            // Agora que o arquivo foi salvo, você pode passar o caminho completo para o método de importação
            
            var designacoesImportadas = _excelImporter.ImportarReunioesExcel(filePath, month);
            var reunioesProgramacao = _scraper.GetReunioes(year, month);

            

            var caminhoWord = _wordReplacer.PreencherReunioesEmModelo(month, year, reunioesProgramacao);
            var caminhoPdf = PreencherPartesEstudantes(month, year, reunioesProgramacao);

            // Retorna os caminhos dos arquivos para o front-end habilitar os botões de download
            return Ok(new
            {
                wordPath = caminhoWord,
                pdfPath = caminhoPdf
            });
        }

        // Exportar todas as programações de reuniões disponíveis a partir do mês atual
        [HttpPost]
        public IActionResult ExportarAutomaticamente()
        {
            DateTime currentDate = DateTime.Now;
            int year = currentDate.Year;
            int month = currentDate.Month;

            string caminhoArquivo = _excelExporter.BuscarArquivo(month,year);
            List<string> caminhosExcel = new List<string>();
            while (caminhoArquivo != "")
            {
                caminhosExcel.Add(caminhoArquivo);
                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
                caminhoArquivo = _excelExporter.BuscarArquivo(month, year);
            }
            List<Reuniao> reunioes = _scraper.GetReunioes(year, month);

            while (reunioes.Any())
            {
                caminhosExcel.Add(_excelExporter.ExportarReunioesParaExcel(month, year, reunioes));

                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
                reunioes = _scraper.GetReunioes(year, month);
            }

            // Retorna os caminhos dos arquivos Excel gerados para o front-end habilitar os botões de download
            return Ok(new { excelPaths = caminhosExcel });
        }


        private string PreencherPartesEstudantes(int month, int year, List<Reuniao> reunioes)
        {
            return _pdfEditor.EditPdfForm(month, year, reunioes);
        }

        // Método auxiliar para realizar o download do arquivo
        [HttpGet]
        public IActionResult DownloadFile(string filePath, string contentType)
        {
            if (!System.IO.File.Exists(filePath))
                return NotFound("Arquivo não encontrado.");

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            var fileName = Path.GetFileName(filePath);
            return File(fileBytes, contentType, fileName);
        }
    }
}