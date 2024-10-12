using System;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace WorkerServiceForResearch
{
    public class GoogleSearch
    {
        public List<SearchResult> Search(string searchTerm)
        {
            var listSearchResults = new List<SearchResult>();
            var options = new ChromeOptions();
            options.AddArgument("--start-maximized");

            using (IWebDriver driver = new ChromeDriver(options))
            {
                driver.Navigate().GoToUrl("https://www.google.com");
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));

                // Aceita cookies se o botão aparecer
                try
                {
                    var acceptButton = wait.Until(drv => drv.FindElement(By.Id("L2AGLb")));
                    acceptButton.Click();
                }
                catch (NoSuchElementException) { }
                
                // Realiza a busca
                var searchBox = driver.FindElement(By.Name("q"));
                searchBox.SendKeys(searchTerm);
                searchBox.Submit();

                // Clica na aba "Web" (ajuste o seletor conforme necessário)
                var webTab = wait.Until(drv => drv.FindElement(By.XPath("//a[@class='nPDzT T3FoJb' and contains(., 'Web')]")));
                webTab.Click();

                // Espera os resultados aparecerem
                wait.Until(drv => drv.FindElement(By.CssSelector("div.g")));

                // Pega os 5 primeiros resultados
                var resultElements = driver.FindElements(By.CssSelector("div.g")).Take(5);
                foreach (var resultElement in resultElements)
                {
                    try
                    {
                        var titleElement = resultElement.FindElement(By.CssSelector("h3"));
                        var urlElement = resultElement.FindElement(By.CssSelector("a"));
                        var descriptionElement = resultElement.FindElement(By.CssSelector(".VwiC3b span")); ;

                        var result = new SearchResult
                        {
                            SearchTerm = searchTerm,
                            Title = titleElement.Text,
                            Url = urlElement.GetAttribute("href"),
                            Description = descriptionElement.Text
                        };

                        listSearchResults.Add(result);
                    }
                    catch (NoSuchElementException) { }
                }
            }

            return listSearchResults;
        }
    }
}
