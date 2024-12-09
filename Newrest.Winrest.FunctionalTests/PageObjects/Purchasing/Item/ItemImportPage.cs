using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class ItemImportPage : PageBase
    {

        public ItemImportPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ________________________________________ Constantes __________________________________

        private const string IMPORT_BUTTON = "//*[@id=\"ImportFileForm\"]/div[3]/input";
        private const string CHOOSE_FILE = "fileSent";
        private const string CHECK_FILE = "//*[@id=\"ImportFileForm\"]/div[3]/button[2]";
        private const string IMPORT_DONE = "//*[@id=\"ImportFileForm\"]/div[2]/h4";
        private const string OK_BUTTON = "//*[@id=\"ImportFileForm\"]/div[3]/button";
        private const string CANCEL_BUTTON = "//*[@id=\"ImportFileForm\"]/div[3]/button";

        // Messages erreur
        private const string ERR_MESSAGE = "//*[@id=\"ImportFileForm\"]/div[2]/div[1]/div/div[1]/div/div[2]/p/b";

        // ________________________________________ Variables
        
        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = CHECK_FILE)]
        private IWebElement _checkFile;

        [FindsBy(How = How.XPath, Using = IMPORT_BUTTON)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = OK_BUTTON)]
        private IWebElement _okBtn;

        [FindsBy(How = How.XPath, Using = CANCEL_BUTTON)]
        private IWebElement _cancelBtn;

        // Messages erreur
        [FindsBy(How = How.XPath, Using = ERR_MESSAGE)]
        private IWebElement _errorMessage;


        // ____________________________________ Méthodes ___________________________________________

        public void ImportFile(string fullPath, bool hasRigths = true)
        {
            CheckFile(fullPath);
            Import();

            // On attend que l'import soit terminé avant de cliquer sur OK
            WaitForElementIsVisible(By.XPath(IMPORT_DONE));
            WaitForLoad();

            _okBtn = WaitForElementIsVisible(By.XPath(OK_BUTTON));
            _okBtn.Click();

            WaitPageLoading();
        }
        public void ImportExcel(string fullPath, bool hasRigths = true)
        {
            CheckFile(fullPath);
       
;

            WaitPageLoading();
        }

        public void CheckFile(string fullPath)
        {
            // Selection d'un fichier à importer
            WaitPageLoading();
            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(fullPath);
            WaitPageLoading();
            WaitForLoad();
            // Vérificaton du fichier
            _checkFile = WaitForElementToBeClickable(By.XPath(CHECK_FILE));
            _checkFile.Click();
            WaitPageLoading();
          //  WaitForLoad();           
        }

        public void Import()
        {
            // Import du fichier
            _import = WaitForElementIsVisible(By.XPath(IMPORT_BUTTON));
            _import.Click();
        }

        // Messages d'erreur
        public bool GetErrorMessageNoRights(string siteName)
        {
            _errorMessage = WaitForElementToBeClickable(By.XPath(ERR_MESSAGE));

            if (_errorMessage.Text.Contains("have permission to work for site '" + siteName + "'") || _errorMessage.Text.Contains("have rights to work for site '" + siteName + "'"))
            {
                return true;
            }

            return false;
        }

        public bool GetErrorMessagePackagingAlreadyExisting()
        {
            _errorMessage = WaitForElementToBeClickable(By.XPath(ERR_MESSAGE));

            string expectedMessage = "An active/inactive packaging with the same site, supplier, packaging, storage value, storage and production quantity already exists for this item. (1 line)";
            if (_errorMessage.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }

        public bool GetErrorMessagePackagingUsedInRecipes()
        {
            _errorMessage = WaitForElementIsVisible(By.XPath(ERR_MESSAGE));

            if (_errorMessage.Text.Contains("inactive this packaging") && _errorMessage.Text.Contains("the last packaging and the item is used in recipes"))
            {
                return true;
            }

            return false;
        }

        public bool GetErrorMessagePackagingWithStock()
        {
            _errorMessage = WaitForElementToBeClickable(By.XPath(ERR_MESSAGE));

            if (_errorMessage.Text.Contains("inactive this packaging because there is a stock with it"))
            {
                return true;
            }

            return false;
        }

        public bool GetErrorMessagePackagingUnit(string newPackagingUnit)
        {
            _errorMessage = WaitForElementToBeClickable(By.XPath(ERR_MESSAGE));

            if (_errorMessage.Text.Contains("The item unit packaging '" + newPackagingUnit + "'"))
            {
                return true;
            }

            return false;
        }

        public bool GetMessageErrorProductionUnit()
        {
            if(isElementVisible(By.XPath(ERR_MESSAGE)))
            {
                WaitForElementIsVisible(By.XPath(ERR_MESSAGE));

                return true;
            }
            else
            {
                return false;
            }
        }

        public void ClosePopup()
        {
            _cancelBtn = WaitForElementIsVisible(By.XPath(CANCEL_BUTTON));
            _cancelBtn.Click();
            WaitForLoad();
        }

    }
}
