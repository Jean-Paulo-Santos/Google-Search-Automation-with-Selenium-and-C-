using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace WorkerServiceForResearch
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly GoogleSearch _googleSearch;

        private readonly string _inputFilePath = @"C:\Users\Cliente\source\repos\WorkerServiceForResearch\input.xlsx";
        private readonly string _outputFilePath = @"C:\Users\Cliente\source\repos\WorkerServiceForResearch\output.xlsx";

        public Worker(ILogger<Worker> logger, GoogleSearch googleSearch)
        {
            _logger = logger;
            _googleSearch = googleSearch;
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("Worker started at: {time}", DateTimeOffset.Now);

            // Lê os termos de busca do arquivo Excel
            List<string> searchTerms = ReadSearchTermsFromExcel(_inputFilePath);
            var searchResults = new List<SearchResult>();

            // Executa as pesquisas e salva os resultados
            foreach (var term in searchTerms)
            {
                var results = _googleSearch.Search(term);  // Utiliza GoogleSearch.cs
                searchResults.AddRange(results);
            }

            SaveResultsToExcel(searchResults, _outputFilePath);
            _logger.LogInformation("Worker completed at: {time}", DateTimeOffset.Now);
            return Task.CompletedTask;
        }

        private List<string> ReadSearchTermsFromExcel(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var searchTerms = new List<string>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int row = 1; worksheet.Cells[row, 1].Value != null; row++)
                {
                    searchTerms.Add(worksheet.Cells[row, 1].Value.ToString());
                }
            }

            return searchTerms;
        }

        private void SaveResultsToExcel(List<SearchResult> searchResults, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Resultados");
                worksheet.Cells[1, 1].Value = "Termo de Busca";
                worksheet.Cells[1, 2].Value = "Título";
                worksheet.Cells[1, 3].Value = "URL";
                worksheet.Cells[1, 4].Value = "Descrição";

                int row = 2;
                foreach (var result in searchResults)
                {
                    worksheet.Cells[row, 1].Value = result.SearchTerm;
                    worksheet.Cells[row, 2].Value = result.Title;
                    worksheet.Cells[row, 3].Value = result.Url;
                    worksheet.Cells[row, 4].Value = result.Description;
                    row++;
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}
