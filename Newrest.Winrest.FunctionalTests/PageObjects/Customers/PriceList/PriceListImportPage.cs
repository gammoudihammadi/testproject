using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList
{
    public class PriceListImportPage : PageBase
    {

        public PriceListImportPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________ Constantes _______________________________________

        private const string FILE_TO_IMPORT = "fileSent";
        private const string CHECK_FILE = "//*[@id=\"ImportFileForm\"]/div[3]/button[2]";
        private const string IMPORT = "//*[@id=\"ImportFileForm\"]/div[3]/input";
        private const string CONFIRM = "//*[@id=\"ImportFileForm\"]/div[3]/button";

        // ____________________________________ Variables ________________________________________

        [FindsBy(How = How.Id, Using = FILE_TO_IMPORT)]
        private IWebElement _fileToImport;

        [FindsBy(How = How.XPath, Using = CHECK_FILE)]
        private IWebElement _checkFile;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = CONFIRM)]
        private IWebElement _confirm;

        // ____________________________________ Méthodes ___________________________________________

        public PriceListDetailsPage ImportFile(string fullPath)
        {
            _fileToImport = WaitForElementIsVisible(By.Id(FILE_TO_IMPORT));             
            _fileToImport.SendKeys(fullPath);

            _checkFile = WaitForElementIsVisible(By.XPath(CHECK_FILE));            
            _checkFile.Click();
            WaitForLoad();

            _import = WaitForElementIsVisible(By.XPath(IMPORT));           
            _import.Click();
            WaitForLoad();

            _confirm = WaitForElementIsVisible(By.XPath(CONFIRM));

            try
            {
                _confirm.Click();
                WaitPageLoading();
            }
            catch
            {
                _webDriver.Navigate().Refresh();
            }

            return new PriceListDetailsPage(_webDriver, _testContext);
        }
    }
}
