using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerCardexNotifPage : PageBase
    {

        public CustomerCardexNotifPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________________________________Constantes____________________________________________________________

        // General
        private const string NEW_CARDEX = "//a[text()='New CARDEX notification']";
        private const string REGISTRATION = "drop-down-registrations";
        private const string OPERATION = "Operation";
        private const string INVOICING = "Invoicing";

        private const string CREATE = "//*[@id=\"form-add-cardex\"]/div[3]/button[2]";

        // Tableau
        private const string EDIT_CARDEX = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[5]/div/a[2]";
        private const string DELETE_CARDEX = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[5]/div/a[1]";
        private const string CONFIRM_DELETE = "/html/body/div[11]/div/div/div[3]/a[1]";
        private const string FIRST_INVOICING = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[4]";
        private const string FIRST_REGISTRATION = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[2]";

        private const string SEARCH_FILTER = "OrderIndex_SearchPattern";

        //_______________________________________________________Variables_____________________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_CARDEX)]
        private IWebElement _newCardex;

        [FindsBy(How = How.Id, Using = REGISTRATION)]
        private IWebElement _cardexRegistration;

        [FindsBy(How = How.Id, Using = OPERATION)]
        private IWebElement _cardexOperation;

        [FindsBy(How = How.Id, Using = INVOICING)]
        private IWebElement _cardexInvoicing;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;


        // Tableau
        [FindsBy(How = How.XPath, Using = EDIT_CARDEX)]
        private IWebElement _editCardex;

        [FindsBy(How = How.XPath, Using = DELETE_CARDEX)]
        private IWebElement _deleteCardex;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = FIRST_INVOICING)]
        private IWebElement _firstInvoicing;

        [FindsBy(How = How.XPath, Using = FIRST_REGISTRATION)]
        private IWebElement _firstRegistration;


        //_________________________________________ Filtres ____________________________________________

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchInput;

        public enum CardexFilterType
        {
            Search
        }

        public void Filter(CardexFilterType filterType, object value)
        {
            switch (filterType)
            {
                case CardexFilterType.Search:
                    _searchInput = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchInput.SetValue(ControlType.TextBox, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(CardexFilterType), filterType, null);

            }

            WaitPageLoading();
            Thread.Sleep(2500);
        }

        // __________________________________________ Méthodes _______________________________________

        // General
        public void CreateCardex(string cardexRegistration, string cardexOperation, string cardexInvoicing)
        {
            _newCardex = WaitForElementIsVisible(By.XPath(NEW_CARDEX));
            _newCardex.Click();
            WaitForLoad();

            _cardexRegistration = WaitForElementIsVisible(By.Id(REGISTRATION));
            _cardexRegistration.Click();
            _cardexRegistration.SendKeys(cardexRegistration);
            WaitForLoad();

            var selectRecipeRegistration = WaitForElementIsVisible(By.XPath("//option[text()='" + cardexRegistration + "']"));
            selectRecipeRegistration.Click();
            WaitForLoad();

            // Renseigner le cardex operation
            _cardexOperation = WaitForElementIsVisible(By.Id(OPERATION));
            _cardexOperation.SetValue(ControlType.TextBox, cardexOperation);

            // Renseigner le cardex invoicing
            _cardexInvoicing = WaitForElementIsVisible(By.Id(INVOICING));
            _cardexInvoicing.SetValue(ControlType.TextBox, cardexInvoicing);

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitPageLoading();
        }

        // Tableau
        public void EditCardex(string invoicing)
        {
            _editCardex = WaitForElementIsVisible(By.XPath(EDIT_CARDEX));
            _editCardex.Click();
            WaitForLoad();

            _cardexInvoicing = WaitForElementIsVisible(By.Id(INVOICING));
            _cardexInvoicing.SetValue(ControlType.TextBox, invoicing);

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitForLoad();
            Thread.Sleep(3000);//WaitForLoad not suffisant to wait creation cardex
        }

        public string GetFirstInvoicing()
        {
            _firstInvoicing = WaitForElementIsVisible(By.XPath(FIRST_INVOICING));
            return _firstInvoicing.Text;
        }

        public string GetFirstRegistration()
        {
            _firstRegistration = WaitForElementIsVisible(By.XPath(FIRST_REGISTRATION));
            return _firstRegistration.Text;
        }

        public bool IsCardexDisplayed()
        {
            if(isElementVisible(By.XPath(FIRST_REGISTRATION)))
            {
                _firstRegistration = _webDriver.FindElement(By.XPath(FIRST_REGISTRATION));
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DeleteCardex()
        {
            _deleteCardex = WaitForElementIsVisible(By.XPath(DELETE_CARDEX));
            _deleteCardex.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

    }
}
