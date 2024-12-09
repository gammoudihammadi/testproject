using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerReinvoicePage : PageBase
    {

        public CustomerReinvoicePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________________ Constantes ______________________________________________

        // General
        private const string NEW_CONFIG = "//*[@id=\"tabContentDetails\"]/div/div[2]/div[1]/div/div/a[2]";
        private const string REINVOICE_FROM_SITE = "FromSiteId";
        private const string TO_SITE = "ToSiteId";
        private const string CREATE = "//*[@id=\"form-add-sitetosite\"]/div[3]/button[2]";
        private const string ERROR_MESSAGE = "//*[@id=\"form-add-sitetosite\"]/div[2]/div/div[3]";

        // Tableau
        private const string EDIT_REINVOICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[4]/div/a[2]";
        private const string DELETE_REINVOICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[4]/div/a[1]";
        private const string CONFIRM_DELETE = "/html/body/div[11]/div/div/div[3]/a[1]";

        private const string FIRST_REINVOICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div";
        private const string FIRST_REINVOICE_FROM_SITE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[2]";
        private const string FIRST_REINVOICE_TO_SITE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[3]";
        private const string FROM_SITE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div[2]";
        private const string NB_REINVOICE = "item-tab-row";

        // Filtre
        private const string SEARCH_FILTER = "SearchPattern";

        //__________________________________________ Variables ______________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_CONFIG)]
        private IWebElement _newReinvoice;

        [FindsBy(How = How.Id, Using = REINVOICE_FROM_SITE)]
        private IWebElement _reinvoiceFromSite;

        [FindsBy(How = How.Id, Using = TO_SITE)]
        private IWebElement _reinvoiceToSite;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;

        // Tableau
        [FindsBy(How = How.XPath, Using = EDIT_REINVOICE)]
        private IWebElement _editReinvoice;

        [FindsBy(How = How.XPath, Using = DELETE_REINVOICE)]
        private IWebElement _deleteReinvoice;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = FIRST_REINVOICE_FROM_SITE)]
        private IWebElement _firstReinvoiceFromSite;

        [FindsBy(How = How.XPath, Using = FIRST_REINVOICE_TO_SITE)]
        private IWebElement _firstReinvoiceToSite;

        //__________________________________________ Filtre __________________________________________________

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchInput;

        public enum ReinvoiceFilterType
        {
            Search
        }

        public void Filter(ReinvoiceFilterType filterType, object value)
        {
            switch (filterType)
            {
                case ReinvoiceFilterType.Search:
                    _searchInput = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchInput.SetValue(ControlType.TextBox, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(ReinvoiceFilterType), filterType, null);
            }

            WaitPageLoading();
            Thread.Sleep(2500);
        }

        //___________________________________________ Méthodes ___________________________________________________

        // General
        public void CreateReinvoice(string reinvoiceFromSite, string reinvoiceToSite)
        {
            _newReinvoice = WaitForElementIsVisible(By.XPath(NEW_CONFIG));
            _newReinvoice.Click();
            WaitForLoad();

            _reinvoiceFromSite = WaitForElementIsVisible(By.Id(REINVOICE_FROM_SITE));
            _reinvoiceFromSite.SetValue(ControlType.DropDownList, reinvoiceFromSite);
            WaitForLoad();

            // Renseigner le site d'arrivée
            _reinvoiceToSite = WaitForElementIsVisible(By.Id(TO_SITE));
            _reinvoiceToSite.SetValue(ControlType.DropDownList, reinvoiceToSite);
            WaitForLoad();

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitPageLoading();
        }

        public bool IsErrorMessageDisplayed()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(ERROR_MESSAGE)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        // Tableau
        public void EditReinvoice(string reinvoiceToSite)
        {
            _editReinvoice = WaitForElementIsVisible(By.XPath(EDIT_REINVOICE));
            _editReinvoice.Click();
            WaitForLoad();

            _reinvoiceToSite = WaitForElementIsVisible(By.Id(TO_SITE));
            _reinvoiceToSite.SetValue(ControlType.DropDownList, reinvoiceToSite);
            _reinvoiceToSite.SendKeys(Keys.Tab);
            WaitForLoad();

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitPageLoading();
        }

        public string GetFirstReinvoiceFromSite()
        {
            _firstReinvoiceFromSite = WaitForElementIsVisible(By.XPath(FIRST_REINVOICE_FROM_SITE));
            return _firstReinvoiceFromSite.Text;
        }

        public string GetFirstReinvoiceToSite()
        {
            WaitForLoad();
            _firstReinvoiceToSite = WaitForElementIsVisible(By.XPath(FIRST_REINVOICE_TO_SITE));
            return _firstReinvoiceToSite.Text;
        }

        public int GetReinvoiceNumber()
        {
            return _webDriver.FindElements(By.ClassName(NB_REINVOICE)).Count;
        }

        public void DeleteReinvoice()
        {
            _deleteReinvoice = WaitForElementIsVisible(By.XPath(DELETE_REINVOICE));
            _deleteReinvoice.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public bool IsReinvoiceDisplayed()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(FIRST_REINVOICE)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
