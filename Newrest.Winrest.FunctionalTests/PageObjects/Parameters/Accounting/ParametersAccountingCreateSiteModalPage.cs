using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Accounting
{
    public class ParametersAccountingCreateSiteModalPage : PageBase
    {

        // _______________________________________ Constantes ____________________________________________

        private const string SITE = "AirportId";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";
        private const string AIRPORT_TAX_VALUE = "Value";
        private const string SAVE = "last";
        private const string ADD_NEW = "SaveAndNew";
        private const string SELECT_ALL = "/html/body/div[11]/div/ul/li[1]/a";

        // _______________________________________ Variables _____________________________________________

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.Id, Using = AIRPORT_TAX_VALUE)]
        private IWebElement _airportTaxValue;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = ADD_NEW)]
        private IWebElement _addNew;

        [FindsBy(How = How.XPath, Using = SELECT_ALL)]
        private IWebElement _selectAll;

        public ParametersAccountingCreateSiteModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public void SetSite(string site)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            _site.SendKeys(Keys.Tab);
        }

        public void SetCustomer(string customer)
        {
            _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
            _customerFilter.Click();

            _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
            _searchCustomer.SetValue(ControlType.TextBox, customer);

            var searchCustomerValue = WaitForElementIsVisible(By.XPath("//span[text()='" + customer + "']"));
            searchCustomerValue.SetValue(ControlType.CheckBox, true);

            _customerFilter.Click();
        }

        public void SetAllCustomer()
        {
            _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
            _customerFilter.Click();

            _selectAll = _webDriver.FindElement(By.XPath(SELECT_ALL));
            _selectAll.Click();

            _customerFilter.Click();
        }

        public void SetValue(string value)
        {
            _airportTaxValue = WaitForElementIsVisible(By.Id(AIRPORT_TAX_VALUE));
            _airportTaxValue.SetValue(ControlType.TextBox, value);
        }

        public ParametersAccounting Save()
        {
            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();

            return new ParametersAccounting(_webDriver, _testContext);
        }

        public ParametersAccounting AddNew()
        {
            _addNew = WaitForElementIsVisible(By.Id(ADD_NEW));
            _addNew.Click();
            WaitForLoad();

            return new ParametersAccounting(_webDriver, _testContext);
        }

    }
}
