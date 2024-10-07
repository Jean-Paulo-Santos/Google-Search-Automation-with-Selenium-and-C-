using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace WorkerServiceForResearch
{
public class Worker : BackgroundService
{
private readonly ILogger<Worker> _logger;

public Worker(ILogger<Worker> logger)
{
_logger = logger;
}

protected override Task ExecuteAsync(CancellationToken stoppingToken)
{
// Log da hora que o trabalhador começa
_logger.LogInformation("Worker started at: {time}", DateTimeOffset.Now);

// Caminho do arquivo Excel de entrada
var inputFilePath = @"C:\Users\Cliente\source\repos\WorkerServiceForResearch\input.xlsx";
// Caminho do arquivo Excel de saída
var outputFilePath = @"C:\Users\Cliente\source\repos\WorkerServiceForResearch\output.xlsx";

// Lê os termos de busca do arquivo Excel
var searchTerms = ReadSearchTermsFromExcel(inputFilePath);
var searchResults = new List<SearchResult>();

// Loop para pesquisar cada termo
foreach (var term in searchTerms)
{
    var results = SearchGoogle(term);
    searchResults.AddRange(results); // Adiciona os resultados encontrados à lista total
}

// Salva os resultados no Excel
SaveResultsToExcel(searchResults, outputFilePath);

// Log da hora que o trabalhador termina
_logger.LogInformation("Worker completed at: {time}", DateTimeOffset.Now);
return Task.CompletedTask;
}

// Método para ler os termos de busca do arquivo Excel
private List<string> ReadSearchTermsFromExcel(string filePath)
{
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var searchTerms = new List<string>();

using (var package = new ExcelPackage(new FileInfo(filePath)))
{
    var worksheet = package.Workbook.Worksheets[0]; // Supondo que seja a primeira planilha

    // Lê as palavras na primeira coluna
    for (int row = 1; worksheet.Cells[row, 1].Value != null; row++)
    {
        searchTerms.Add(worksheet.Cells[row, 1].Value.ToString());
    }
}

return searchTerms; // Retorna a lista de termos de busca
}

// Método para buscar no Google usando Selenium
private List<SearchResult> SearchGoogle(string searchTerm)
{
var searchResults = new List<SearchResult>();

var options = new ChromeOptions();
options.AddArgument("--start-maximized"); // Abre o navegador em tela cheia

// Configura o ChromeDriver
using (IWebDriver driver = new ChromeDriver(options))
{
    driver.Navigate().GoToUrl("https://www.google.com"); // Acessa o Google

    // Aguarda um pouco para garantir que a página carregou
    Thread.Sleep(1000); 

            
        // Cria uma espera explícita de 10 segundos
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1));
        var acceptButton = wait.Until(drv =>
        {
                
                // Tenta encontrar o botão de aceitar cookies
                var button = drv.FindElement(By.XPath("//*[@id=\"L2AGLb\"]"));
                return button; // Retorna o botão se encontrado
                
        });

        if (acceptButton != null) // Se o botão foi encontrado
        {
            acceptButton.Click(); // Clica no botão
        }
            

    // Encontra a barra de pesquisa e digita o termo de busca
    var searchBox = driver.FindElement(By.Name("q")); // Encontra a barra de pesquisa
    searchBox.SendKeys(searchTerm); // Digita o termo de busca
    searchBox.Submit(); // Submete a busca

    // Espera carregar os resultados
    WebDriverWait resultsWait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
    resultsWait.Until(drv => drv.FindElement(By.CssSelector("div.g"))); // Espera os resultados aparecerem

    // Pega os 5 primeiros resultados (títulos, URLs e descrições)
    var resultElements = driver.FindElements(By.CssSelector("div.g")).Take(5); // Encontra os elementos de resultado

    foreach (var resultElement in resultElements)
    {
        try
        {
            // Extrai título, URL e descrição dos resultados
            var titleElement = resultElement.FindElement(By.TagName("h3"));
            var urlElement = resultElement.FindElement(By.CssSelector("a"));
            var descriptionElement = resultElement.FindElement(By.CssSelector("span.st"));

            // Cria um novo resultado de pesquisa
            var result = new SearchResult
            {
                SearchTerm = searchTerm,
                Title = titleElement.Text,
                Url = urlElement.GetAttribute("href"),
                Description = descriptionElement.Text
            };

            searchResults.Add(result); // Adiciona o resultado à lista
        }
        catch (NoSuchElementException)
        {
            // Ignora se algum resultado não tiver título ou URL
        }
    }
}

return searchResults; // Retorna a lista de resultados encontrados
}





// Método para salvar os resultados no Excel
private void SaveResultsToExcel(List<SearchResult> searchResults, string filePath)
{
using (var package = new ExcelPackage())
{
    var worksheet = package.Workbook.Worksheets.Add("Resultados"); // Cria uma nova planilha

    // Cabeçalho
    worksheet.Cells[1, 1].Value = "Termo de Busca";
    worksheet.Cells[1, 2].Value = "Título";
    worksheet.Cells[1, 3].Value = "URL";
    worksheet.Cells[1, 4].Value = "Descrição";

    // Escreve os dados
    int row = 2;
    foreach (var result in searchResults)
    {
        worksheet.Cells[row, 1].Value = result.SearchTerm;
        worksheet.Cells[row, 2].Value = result.Title;
        worksheet.Cells[row, 3].Value = result.Url;
        worksheet.Cells[row, 4].Value = result.Description;
        row++; // Incrementa a linha
    }

    // Salva o arquivo
    package.SaveAs(new FileInfo(filePath));
}
}
}

// Classe para mapear os resultados da pesquisa
public class SearchResult
{
public string SearchTerm { get; set; }
public string Title { get; set; }
public string Url { get; set; }
public string Description { get; set; }
}
}
