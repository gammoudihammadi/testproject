using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet
{
    public class DatasheetCreateModalPage : PageBase
    {
        public DatasheetCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes ____________________________________________

        private const string DATASHEET_NAME = "datasheet-name";
        private const string CUSTOMER_CODE = "Datasheet_CustomerCode";
        private const string GUEST_TYPE = "Datasheet_GuestTypeId";
        private const string SITE_FILTER = "//*[@id=\"createFormMenu\"]/table/tbody/tr/td[2]/input";
        private const string SITE_SELECTED = "//*[@id=\"createFormMenu\"]/div[6]/table/tbody/tr[*]/td[2]";
        private const string CREATE = "//*[@id=\"form-createdit-menu\"]/div[3]/div/button[2]";
        private const string SELECT_ALL = "//*[@id=\"datasheet-editSitesTable\"]/tbody/tr/td[3]/a";

        private const string MODAL_ERREUR = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/p";
        private const string MSG_ERR_NAME = "//*[@id=\"createFormMenu\"]/div[1]/div/span";
        private const string SITE_CHECK_BOX = "//*/td[@class='site-name-col' and contains(text(),'{0}')]";
        private const string DATASHEET_NAMEVERIF = "//*[@id=\"div-body\"]/div/div[2]/div[1]";

        
        // _________________________________________ Variables _____________________________________________

        [FindsBy(How = How.Id, Using = DATASHEET_NAME)]
        private IWebElement _datasheetName;

        [FindsBy(How = How.Id, Using = CUSTOMER_CODE)]
        private IWebElement _customerCode;

        [FindsBy(How = How.Id, Using = GUEST_TYPE)]
        private IWebElement _guestType;

        [FindsBy(How = How.XPath, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _create;

        [FindsBy(How = How.XPath, Using = MODAL_ERREUR)]
        private IWebElement _modaleErreur;

        [FindsBy(How = How.XPath, Using = MSG_ERR_NAME)]
        private IWebElement _msgErrName;

        // _________________________________________ Méthodes ______________________________________________

        public DatasheetDetailsPage FillField_CreateNewDatasheet(string datasheetName, string guestType, string site, string customerCode = null)
        {
            _datasheetName = WaitForElementIsVisible(By.Id(DATASHEET_NAME));
            _datasheetName.SetValue(ControlType.TextBox, datasheetName);

            if (customerCode != null)
            {
                _customerCode = WaitForElementIsVisible(By.Id(CUSTOMER_CODE));
                _customerCode.SetValue(ControlType.TextBox, customerCode);
            }

            _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
            _guestType.SetValue(ControlType.DropDownList, guestType);

            _siteFilter = WaitForElementIsVisible(By.XPath(SITE_FILTER));
            _siteFilter.SetValue(ControlType.TextBox, site);

            // On sélectionne le site désiré
            SelectSite(site);

            //_siteFilter = WaitForElementIsVisible(By.XPath(SITE_FILTER));
            //_siteFilter.SetValue(ControlType.TextBox, "ACE");

            // On sélectionne le site désiré
            //SelectSite("ACE");

            // Click sur le bouton "Create"
            _create = WaitForElementToBeClickable(By.XPath(CREATE));
            _create.Click();
            WaitPageLoading();
            // secouse de la page web
            Thread.Sleep(2000);
            WaitForLoad();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }
        public DatasheetDetailsPage FillField_CreateNewDatasheetWidh2Sites(string datasheetName, string guestType, string site1, string site2, string customerCode = null)
        {
            _datasheetName = WaitForElementIsVisible(By.Id(DATASHEET_NAME));
            _datasheetName.SetValue(ControlType.TextBox, datasheetName);

            if (customerCode != null)
            {
                _customerCode = WaitForElementIsVisible(By.Id(CUSTOMER_CODE));
                _customerCode.SetValue(ControlType.TextBox, customerCode);
            }

            _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
            _guestType.SetValue(ControlType.DropDownList, guestType);

            _siteFilter = WaitForElementIsVisible(By.XPath(SITE_FILTER));
            _siteFilter.SetValue(ControlType.TextBox, site1);

            // On sélectionne le 1er site désiré
            SelectSite(site1);

            _siteFilter = WaitForElementIsVisible(By.XPath(SITE_FILTER));
            _siteFilter.SetValue(ControlType.TextBox, site2);

            // On sélectionne le 2 éme site désiré
            SelectSite(site2);

            // Click sur le bouton "Create"
            _create = WaitForElementToBeClickable(By.XPath(CREATE));
            _create.Click();
            WaitPageLoading();
            // secouse de la page web
            Thread.Sleep(2000);
            WaitForLoad();

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public void FillField_CreateNewDatasheetWithoutSite(string datasheetName, string guestType)
        {
            _datasheetName = WaitForElementIsVisible(By.Id(DATASHEET_NAME));
            _datasheetName.SetValue(ControlType.TextBox, datasheetName);

            _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
            _guestType.SetValue(ControlType.DropDownList, guestType);

            // Click sur le bouton "Create"
            _create = WaitForElementToBeClickable(By.XPath(CREATE));
            _create.Click();
            WaitPageLoading();
            // secouse de la page web
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public void SelectSite(string site)
        {
            var siteCheckBox = WaitForElementExists(By.XPath(string.Format(SITE_CHECK_BOX, site)));
            siteCheckBox.Click();
        }

        public bool IsErrorMessageSiteDisplayed()
        {
            if(isElementVisible(By.XPath(MODAL_ERREUR)))
            {
                _modaleErreur = _webDriver.FindElement(By.XPath(MODAL_ERREUR));
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsErrorMessageNameDisplayed()
        {
            if(isElementVisible(By.XPath(MSG_ERR_NAME)))
            {
                _msgErrName = _webDriver.FindElement(By.XPath(MSG_ERR_NAME));
                return true;
            }
            else
            {
                return false;
            }
        }
        public void FillName_NewDatasheet(string datasheetName)
        {
            // Remplir le champ "Name" avec le nom de la datasheet
            WaitPageLoading();
            _webDriver.FindElement(By.Id(DATASHEET_NAME)).SendKeys(datasheetName);

     
        }
        public void SelectAllSites()
        {
            // Trouver et cocher la case "Select All" pour les sites
            var selectAllCheckbox = _webDriver.FindElement(By.XPath(SELECT_ALL));
            if (!selectAllCheckbox.Selected)
            {
                selectAllCheckbox.Click();
            }
        }
        public void ClickCreateButton()
        {
            WaitPageLoading();
            var createButton = _webDriver.FindElement(By.XPath(CREATE));
            createButton.Click();
        }
        public bool IsDatasheetCreatedSuccessfully(string expectedDatasheetName)
        {
            WaitPageLoading();
            try
            {
                WaitPageLoading();
                WaitForLoad();
                // Localiser l'élément dans la page qui contient le nom de la datasheet
                var datasheetNameElement = _webDriver.FindElement(By.XPath(DATASHEET_NAMEVERIF));
                WaitPageLoading();
                WaitForLoad();
                // Récupérer le texte affiché dans cet élément
                string actualDatasheetName = datasheetNameElement.Text;
                WaitPageLoading();
                WaitForLoad();
                // Vérifier si le texte récupéré contient la chaîne de caractères attendue (expectedDatasheetName)
                return actualDatasheetName.Contains(expectedDatasheetName);
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }


    }
}
