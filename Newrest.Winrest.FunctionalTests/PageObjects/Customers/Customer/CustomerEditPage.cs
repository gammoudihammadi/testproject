using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerEditPage : PageBase
    {

        ////Number
        //[FindsBy(How = How.XPath, Using = CLAIM_NUMBER_XPATH)]
        //private IWebElement _inventaryID;

        //Date
        [FindsBy(How = How.XPath, Using = "//*[@id=\"dropdown-aircraft-details-selectized\"]")]
        private IWebElement _aircraftNumber;

        public CustomerEditPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void ActivateAirportTax()
        {
            throw new NotImplementedException();
        }
    }
}
