using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SupplyOrderImport : PageBase
    {
        // ________________________________ Constantes ________________________________________

        private const string IMPORT_BUTTON = "//*[@id=\"ImportFileForm\"]/div[3]/button[2]";
        private const string CHOOSE_FILE = "fileSent";
        private const string CHECK_FILE = "//*[@id=\"ImportFileForm\"]/div[3]/button[2]";
        private const string IMPORT = "//*[@id=\"ImportFileForm\"]/div[3]/button[2]";

        // ________________________________ Variables _________________________________________

        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = CHECK_FILE)]
        private IWebElement _checkFile;

        [FindsBy(How = How.XPath, Using = IMPORT_BUTTON)]
        private IWebElement _import;

        // _________________________________ Méthodes ________________________________________

        public SupplyOrderImport(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }                

        public void ImportFile(string fullPath)
        {
            CheckFile(fullPath);
            Import();
        }

        public void CheckFile(string fullPath)
        {
            // Selection d'un fichier à importer
            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(fullPath);

            // Vérificaton du fichier
            _checkFile = WaitForElementToBeClickable(By.XPath(CHECK_FILE));
            _checkFile.Click();
            WaitForLoad();
        }

        public void Import()
        {
            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();
            WaitForLoad();
        }
    }
}
