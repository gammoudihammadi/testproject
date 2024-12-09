using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerPortalCustomerOrdersPageModal : PageBase
    {
        public CustomerPortalCustomerOrdersPageModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _______________________________________________

        private const string SELECT_SITE = "//*[@id=\"SelectedSite\"]";
        private const string SELECT_CUSTOMER = "//*[@id=\"SelectedCustomer\"]";
        private const string SELECT_AIRCRAFT = "//*[@id=\"Aircraftt\"]";

        private const string CREATE_BUTTON = "//*[@id=\"btn-valid-claim\"]";

        private const string RESULT_CUSTOMER_ORDER = "/html/body/div[4]/div/div[1]/h2";

        public void SelectSite(string site)
        {
            var siteField = WaitForElementIsVisible(By.XPath(SELECT_SITE));
            var selectSite = new SelectElement(siteField);
            selectSite.SelectByText(site);
        }

        public void SelectCustomer(string customer)
        {
            var customerField = WaitForElementIsVisible(By.XPath(SELECT_CUSTOMER));
            var selectCustomer = new SelectElement(customerField);
            selectCustomer.SelectByText(customer);
        }

        public void SelectAircraft(string aircraft)
        {
            var airCraftField = WaitForElementIsVisible(By.XPath(SELECT_AIRCRAFT));
            var selectAirCraft = new SelectElement(airCraftField);
            selectAirCraft.SelectByText(aircraft);
        }

        public void Create()
        {
            var button = WaitForElementIsVisible(By.XPath(CREATE_BUTTON));
            button.Click();
        }

        public string ResultCustomerOrder()
        {
            var customerOrderField = WaitForElementIsVisible(By.XPath(RESULT_CUSTOMER_ORDER));
            //Customer order 32323 - EMIRATES AIRLINES
            return customerOrderField.Text;
        }
    }
}
