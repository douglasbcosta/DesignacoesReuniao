using DesignacoesReuniao.Infra.Excel;
using DesignacoesReuniao.Infra.Interfaces;
using DesignacoesReuniao.Infra.Pdf;
using DesignacoesReuniao.Infra.Repostories;
using DesignacoesReuniao.Infra.Repostories.Interface;
using DesignacoesReuniao.Infra.Scraper;
using DesignacoesReuniao.Infra.Word;
using Microsoft.Extensions.DependencyInjection;

namespace DesignacoesReuniao.CrossCutting.Injections
{
    public static class DependencesInjection
    {
        public static IServiceCollection ConfigureDependences(this IServiceCollection services)
        {
            // Registrar as interfaces e suas implementações
            services.AddTransient<IWebScraper, WebScraper>(provider => new WebScraper("https://wol.jw.org/pt/wol/meetings/r5/lp-t"));
            services.AddTransient<IExcelExporter, ExcelExporter>();
            services.AddTransient<IWordReplacer, WordReplacer>();
            services.AddTransient<IPdfEditor, PdfEditor>();
            services.AddTransient<IExcelImporter, ExcelImporter>();
            services.AddScoped<IPessoaRepository, PessoaRepository>();

            return services;
        }
    }
}
