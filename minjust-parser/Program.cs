using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using minjust_parser.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace minjust_parser
{
    class Program
    {
        private static IWebDriver driver;
        static void Main(string[] args)
        {

            /*Helpers.ExcelHelper.Read();*/

            driver = new ChromeDriver();

            driver.Navigate().GoToUrl(ParserSettings.Url);
            Thread.Sleep(300);

            var radioButtons = driver.FindElements(By.ClassName("form-check-input"));
            radioButtons[0].Click();

            var searchField = driver.FindElement(By.XPath("/html/body/app-root/div/content-wrapper/div/div/free-search-main/div/search-information/api-dynamic-form-builder/div/form/div[2]/div/div[1]/app-input-field/div/div/input"));

            searchField.SendKeys(ParserSettings.TestValue);
            
            var searchButton=driver.FindElement(By.ClassName("button-primary"));

            searchButton.Click();

            Thread.Sleep(100000);
        }
    }
}