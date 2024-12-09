using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceImportPage : PageBase
    {
        public ServiceImportPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ________________________________________ Constantes __________________________________

        private const string IMPORT_BUTTON = "//*[@id=\"form-import\"]/div[3]/button[2]";
        private const string CHOOSE_FILE = "fileSent";
        private const string CHECK_FILE = "//*[@id=\"form-import\"]/div[3]/button[2]";
        private const string IMPORT_OF_SERVICES_PRICES_ERROR_MSG = "//*[@id=\"form-import\"]/div[2]/div/div/p[1]/b";
        private const string IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON = "//*[@id=\"form-import\"]/div[3]/button";
        private const string OK_BUTTON = "//*[@id=\"ImportFileForm\"]/div[3]/button";
        private const string IMPORT_BTN = "//*[@id=\"ImportFileForm\"]/div[2]/button[2]";
        private const string OK_IMPORT_BUTTON = "//*[@id=\"modal-1\"]/div/div/div[3]/button";
        private const string CHECKFILE_POPUP_IS_MESSAGE = "//*[@id=\"form-import\"]/div[2]/p";
        private const string CHECKFILE_POPUP_MESSAGE_ERROR = "//*[@id=\"form-import\"]/div[2]/div/div/p[1]/b";
        private const string CHECKFILE_POPUP_MESSAGE = "//*[@id=\"form-import\"]/div[2]/p/span[1]";
        private const string CHECKFILE_POPUP_TOADD = "//*[@id=\"form-import\"]/div[2]/p/span[2]";
        private const string CHECKFILE_POPUP_TOUPDATE = "//*[@id=\"form-import\"]/div[2]/p/span[3]";
        private const string CHECKFILE_POPUP_TODELETE = "//*[@id=\"form-import\"]/div[2]/p/span[4]";
        private const string CHECKFILE_POPUP_UPDATE_GATES = "//*[@id=\"isGatesToUpdate\"]";
        // ________________________________________ Variables

        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = IMPORT_BTN)]
        private IWebElement _importBtn;

        [FindsBy(How = How.XPath, Using = CHECK_FILE)]
        private IWebElement _checkFile;

        [FindsBy(How = How.XPath, Using = IMPORT_BUTTON)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON)]
        private IWebElement _closeButton;

        [FindsBy(How = How.XPath, Using = OK_BUTTON)]
        private IWebElement _okBtn;

        [FindsBy(How = How.XPath, Using = OK_IMPORT_BUTTON)]
        private IWebElement _okImportButton;

        [FindsBy(How = How.XPath, Using = CHECKFILE_POPUP_UPDATE_GATES)]
        private IWebElement _checkfile_popup_update_gates;

        // ____________________________________ Méthodes ___________________________________________

        public void CheckFile(string fullPath)
        {
            // Selection d'un fichier à importer
            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(fullPath);
            WaitForLoad();

            // Vérificaton du fichier
            _checkFile = WaitForElementToBeClickable(By.XPath(CHECK_FILE));
            _checkFile.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public void ImportFlightsExcelFile(string fullPath, bool hasRigths = true)
        {
            //selection du fichier a importer
            CheckFileImportFlight(fullPath);

            // On attend que l'import soit terminé avant de cliquer sur OK
             _okImportButton = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/button"));
            _okImportButton.Click();
            WaitForLoad();
        }
        public void CheckFileImportFlight(string fullPath)
        {
            // Selection d'un fichier à importer
            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(fullPath);
            WaitForLoad();

            //check update gates
            _checkfile_popup_update_gates=_webDriver.FindElement(By.XPath(CHECKFILE_POPUP_UPDATE_GATES));
            _checkfile_popup_update_gates.Click();
            WaitForLoad();

            // Vérificaton que le bouton d'import est visible
            _importBtn = WaitForElementIsVisible(By.XPath(IMPORT_BTN));
            _importBtn.Click();
            WaitForLoad();
        }
        

        public bool ImportWithFail()
        {
            // Import du fichier avec erreur
            if (isElementVisible(By.XPath(IMPORT_BUTTON)))
            {
                _import = WaitForElementIsVisible(By.XPath(IMPORT_BUTTON));
                _import.Click();
                WaitForLoad();
            }

            if (isElementVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_ERROR_MSG)))
            {
                _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                _closeButton.Click();
                WaitForLoad();
                return false;
            }
            else
            {
                // raté c'est un Success popup
                _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                _closeButton.Click();
                WaitForLoad();
                return true;
            }
        }

        public bool ImportWithFail2()
        {
            if (isElementVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_ERROR_MSG)))
            {
                _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                _closeButton.Click();
                WaitForLoad();
                return false;
            }
            else
            {
                // raté c'est un Success popup
                _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                _closeButton.Click();
                WaitForLoad();
                return true;
            }
        }


        public bool Import()
        {
            // Import du fichier avec succès
            if (isElementVisible(By.XPath(IMPORT_BUTTON)))
            {
                _import = WaitForElementIsVisible(By.XPath(IMPORT_BUTTON));
                _import.Click();
                WaitForLoad();
            }

            if (isElementVisible(By.XPath(CHECKFILE_POPUP_IS_MESSAGE)))
            {
                var message = WaitForElementIsVisible(By.XPath(CHECKFILE_POPUP_IS_MESSAGE));
                bool isMessageSuccess = ("Import of prices done successfully." == message.Text);
                _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                _closeButton.Click();
                WaitForLoad();
                return isMessageSuccess;
            } 
            else
            {
                // raté c'est un Fail popup
                _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                _closeButton.Click();
                WaitForLoad();
                return false;
            }
        }

        public bool ImportModal()
        {
            if (isElementVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_ERROR_MSG)))
            {
                WaitForLoad();
                return false;
            }
            else
            {
                if (isElementVisible(By.XPath(CHECKFILE_POPUP_MESSAGE)))
                {
                    var done = WaitForElementIsVisible(By.XPath(CHECKFILE_POPUP_MESSAGE));
                    if (done.Text == "Verification step done successfully.")
                    {
                        _import = WaitForElementIsVisible(By.XPath(IMPORT_BUTTON));
                        _import.Click();
                        WaitForLoad();

                        _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                        _closeButton.Click();
                    }
                }
                else
                {
                    _closeButton = WaitForElementIsVisible(By.XPath(IMPORT_OF_SERVICES_PRICES_CLOSE_BUTTON));
                    _closeButton.Click();
                    WaitForLoad();
                }
                return true;
            }
        }

        public string CheckFilePopupMessage()
        {
            var Verification = WaitForElementIsVisible(By.XPath(CHECKFILE_POPUP_MESSAGE), "CHECKFILE_POPUP_MESSAGE");
            return Verification.Text;
        }

        public string CheckFilePopupMessageError()
        {
            var Verification = WaitForElementIsVisible(By.XPath(CHECKFILE_POPUP_MESSAGE_ERROR), "CHECKFILE_POPUP_MESSAGE_ERROR");
            return Verification.Text;
        }
        

        public string CheckFilePopupToAdd()
        {
            var ToAdd = WaitForElementIsVisible(By.XPath(CHECKFILE_POPUP_TOADD), "CHECKFILE_POPUP_TOADD");
            return ToAdd.Text;
        }

        public string CheckFilePopupToUpdate()
        {
            var ToUpdate = WaitForElementIsVisible(By.XPath(CHECKFILE_POPUP_TOUPDATE), "CHECKFILE_POPUP_TOUPDATE");
            return ToUpdate.Text;
        }

        public string CheckFilePopupToDelete()
        {
            var ToDelete = WaitForElementIsVisible(By.XPath(CHECKFILE_POPUP_TODELETE), "CHECKFILE_POPUP_TODELETE");
            return ToDelete.Text;
        }

        public void CheckFileClickImportFile()
        {
            var buttonImportFile = WaitForElementIsVisible(By.XPath("//*/button[text()='Import File']"));
            buttonImportFile.Click();
            WaitForLoad();
        }
    }
}
