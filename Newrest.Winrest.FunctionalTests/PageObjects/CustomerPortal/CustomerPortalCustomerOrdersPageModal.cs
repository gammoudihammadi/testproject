using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
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

        public void SelectSite(string site)
        {
            var siteField = WaitForElementExists(By.Id("drop-down-sites"));
            var selectSite = new SelectElement(siteField);
            selectSite.SelectByText(site);
        }

        public void SelectCustomer(string customer)
        {
            // dev & patch
            // multiple select jquery (SelectedCustomer is hidden)
            var customerField = WaitForElementExists(By.Id("drop-down-customers"));
            var selectCustomer = new SelectElement(customerField);
            WaitForLoad();
            selectCustomer.SelectByText(customer);
        }

        public void SelectDeliveryName(string deliveryName)
        {
            var deliveryNameField = WaitForElementExists(By.Id("DeliveryName"));
            deliveryNameField.SetValue(ControlType.TextBox, deliveryName);
            WaitForLoad();
        }

        public void SelectAircraft(string aircraft)
        {
            // flight-detail est hidden
            var flightDetail = WaitForElementExists(By.Id("flight-detail"));
            if(!flightDetail.GetAttribute("style").Contains("display: none;"))
            {
                var aircraftElement = WaitForElementExists(By.Id("Aircraftt"));
                var selectAircraft = new SelectElement(aircraftElement);
                selectAircraft.SelectByText(aircraft);
                WaitForLoad();
            }
            return;
          
        }

        public CustomerPortalCustomerOrdersPageResult Create()
        {
            var button = WaitForElementIsVisible(By.XPath(CREATE_BUTTON));
            button.Click();
            WaitForLoad();
            return new CustomerPortalCustomerOrdersPageResult(_webDriver, _testContext);
        }
    }
}
