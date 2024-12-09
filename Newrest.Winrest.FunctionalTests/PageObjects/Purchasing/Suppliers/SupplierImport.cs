using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers
{
    public class SupplierImport : PageBase
    {

        public SupplierImport(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ___________________________________________ Constantes ______________________________________________

        private const string CHOOSE_FILE = "fileSent";
        private const string CHECK_FILE_BTN = "//*[@id=\"ImportFileForm\"]/div[3]/button[2]";
        private const string IMPORT = "//*[@id=\"CopyPriceForm\"]/div[3]/button[2]";
        private const string OK_BTN = "//*[@id=\"modal-1\"]/div/div/div[3]/button";
        private const string IMPORT_DONE = "//*[@id=\"modal-1\"]/div/div/div[2]/div/h4[@style='color:Green']";

        // ___________________________________________ Variables _______________________________________________

        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = CHECK_FILE_BTN)]
        private IWebElement _checkFileBtn;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = OK_BTN)]
        private IWebElement _OKBtn;



        // __________________________________________ METHODES ___________________________________________

        public bool ImportFile(string fullPath)
        {
            bool valueBool = true;
            try
            {
                CheckFile(fullPath);
                Import();

                // On attend que l'import soit terminé avant de cliquer sur OK
                    WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/h4[@style='color:Green']"));
                
                WaitForLoad();

                    _OKBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/button"));
               
                _OKBtn.Click();
                WaitForLoad();

            }
            catch (Exception)
            {
                valueBool = false;
            }
            return valueBool;
        }

        public void CheckFile(string fullPath)
        {
            // Selection d'un fichier à importer
            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(fullPath);
            WaitForLoad();
            // Vérificaton du fichier
            _checkFileBtn = WaitForElementToBeClickable(By.XPath(CHECK_FILE_BTN));
            _checkFileBtn.Click();
            WaitForLoad();
        }

        public void Import()
        {
            // Import du fichier
            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();
            WaitForLoad();
        }

    }
}
