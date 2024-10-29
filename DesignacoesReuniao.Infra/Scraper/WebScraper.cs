using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Extensions;
using DesignacoesReuniao.Infra.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Globalization;

namespace DesignacoesReuniao.Infra.Scraper;
public class WebScraper : IWebScraper
{
    private readonly string _baseUrl;

    public WebScraper(string baseUrl)
    {
        _baseUrl = baseUrl;
    }
    public List<Reuniao> GetReunioes(int year, int month)
    {
        var options = new ChromeOptions();
        options.AddArgument("--headless");
        options.AddArgument("--disable-dev-shm-usage");
        options.AddArgument("--headless");
        options.AddArgument("--no-sandbox");
        // Simular um navegador real
        options.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
        options.AddArgument("--disable-blink-features=AutomationControlled");

        List<Reuniao> reunioes = new List<Reuniao>();

        using (IWebDriver driver = new ChromeDriver(options))
        {
            var (firstWeek, lastWeek) = GetWeekRange(year, month);
            bool programacaoDisponivel = ProgramacaoDisponivel(year, driver, firstWeek);
            if (!programacaoDisponivel)
            {
                return reunioes;
            }
            for (int week = firstWeek; week <= lastWeek; week++)
            {
                string url = $"{_baseUrl}/{year}/{week}";
                driver.Navigate().GoToUrl(url);
                bool pageIsReady = false;
                while (!pageIsReady)
                {
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    pageIsReady = (bool)js.ExecuteScript("return document.readyState == 'complete'");
                }
                Reuniao reuniao = ProcessarReuniao(driver, month, year, week);
                if (reuniao != null)
                {
                    reunioes.Add(reuniao);
                }
            }

            driver.Quit();
        }

        return reunioes;
    }

    private bool ProgramacaoDisponivel(int year, IWebDriver driver, int week)
    {
        string url = $"{_baseUrl}/{year}/{week}";
        driver.Navigate().GoToUrl(url);
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        bool pageIsReady = false;
        while (!pageIsReady)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            pageIsReady = (bool)js.ExecuteScript("return document.readyState == 'complete'");
        }
        var bannerErro = driver.FindElements(By.CssSelector(".bannerErro"));
        bool programacaoDisponivel = bannerErro.Count() == 0;
        return programacaoDisponivel;
    }


    private (int firstWeek, int lastWeek) GetWeekRange(int year, int month)
    {
        DateTime firstDayOfMonth = new DateTime(year, month, 1);
        DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

        Calendar calendar = CultureInfo.CurrentCulture.Calendar;
        int firstWeek = calendar.GetWeekOfYear(firstDayOfMonth, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
        int lastWeek = calendar.GetWeekOfYear(lastDayOfMonth, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

        if (firstDayOfMonth.DayOfWeek != DayOfWeek.Monday)
        {
            firstWeek++;
        }

        return (firstWeek, lastWeek);
    }

    private Reuniao ProcessarReuniao(IWebDriver driver, int month, int year, int week)
    {
        DateOnly dataSemana = ObterDataDaSemana(year, week);
        string tituloSemana = ObterTituloSemana(driver);
        string leituraSemana = ObterLeituraDaSemana(driver);
        if (string.IsNullOrEmpty(tituloSemana)) return null;

        Reuniao reuniao = new Reuniao { 
            Semana = tituloSemana, 
            InicioSemana = dataSemana,
            LeituraDaSemana = leituraSemana
        };

        AdicionarCanticos(driver, reuniao);
        AdicionarSessoes(driver, reuniao, month, year);

        return reuniao;
    }

    public static DateOnly ObterDataDaSemana(int ano, int semanaDoAno)
    {
        DateTime primeiroDiaDoAno = new DateTime(ano, 1, 1);
        var cultura = CultureInfo.CurrentCulture;
        int semanaPrimeiroDia = cultura.Calendar.GetWeekOfYear(primeiroDiaDoAno, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

        int diasAteSemana = (semanaDoAno - semanaPrimeiroDia) * 7;
        DateTime dataDaSemana = primeiroDiaDoAno.AddDays(diasAteSemana);

        return DateOnly.FromDateTime(dataDaSemana);
    }

    private string ObterTituloSemana(IWebDriver driver)
    {
        try
        {
            var weekTitleElement = driver.FindElement(By.XPath("//h1"));
            string tituloSemana = weekTitleElement.Text;
            return tituloSemana;
        }
        catch (NoSuchElementException)
        {
            Console.WriteLine("Título da semana não encontrado.");
            return null;
        }
    }

    private string ObterLeituraDaSemana(IWebDriver driver)
    {
        try
        {
            var weeklyReadingElement = driver.FindElement(By.CssSelector("h2 > a"));
            string leituraDaSemana = weeklyReadingElement.Text;
            return leituraDaSemana;
        }
        catch (NoSuchElementException)
        {
            Console.WriteLine("Leitura da semana não encontrada.");
            return null;
        }
    }

    private void AdicionarCanticos(IWebDriver driver, Reuniao reuniao)
    {
        // Busca todos os elementos que contêm a palavra "Cântico" no texto
        var canticoElements = driver.FindElements(By.XPath("//*[contains(text(), 'Cântico')]"));
        foreach (var canticoElement in canticoElements)
        {
            string canticoTexto = canticoElement.Text;
            reuniao.AdicionarCantico(canticoTexto);
        }
    }

    private void AdicionarSessoes(IWebDriver driver, Reuniao reuniao, int month, int year)
    {
        var sessionTitleElements = driver.FindElements(By.XPath("//h2[contains(@class, 'color')]"));
        if (sessionTitleElements.Count == 0)
        {
            Console.WriteLine("Nenhuma sessão encontrada para o mês selecionado.");
            return;
        }

        foreach (var element in sessionTitleElements)
        {
            string tituloSessao = element.Text;

            Sessao sessao = new Sessao(tituloSessao);
            string corClasse = ExtrairClasseCor(element.GetAttribute("class"));

            AdicionarPartes(driver, sessao, corClasse);
            reuniao.AdicionarSessao(sessao);
        }
    }

    private void AdicionarPartes(IWebDriver driver, Sessao sessao, string corClasse)
    {
        var partesElements = driver.FindElements(By.XPath($"//h3[contains(@class, '{corClasse}') and not(ancestor::div[contains(@class, 'boxContent')])]"));
        foreach (var parteElement in partesElements)
        {
            string[] parte = parteElement.Text.Split('.');
            int indice = int.Parse(parte[0]);
            string tituloParte = parte[1];

            var tempoElement = parteElement.FindElement(By.XPath("following-sibling::div//p"));
            string tempoTexto = tempoElement.Text;
            int tempoMinutos = tempoTexto.ExtrairTempo();

            sessao.AdicionarParte(new Parte(indice, tituloParte, tempoMinutos));
        }
    }

    private string ExtrairClasseCor(string classes)
    {
        var classList = classes.Split(' ');
        foreach (var className in classList)
        {
            if (className.Contains("color"))
            {
                return className;
            }
        }
        return string.Empty;
    }

}