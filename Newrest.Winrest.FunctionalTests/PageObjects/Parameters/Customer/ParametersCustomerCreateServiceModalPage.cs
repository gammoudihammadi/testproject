using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObject.Parameters.Customer;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Customer
{
    public class ParametersCustomerCreateServiceModalPage : PageBase
    {

        // _________________________________________ Constantes _______________________________________

        private const string SAVE = "last";
        private const string ADD_NEW = "SaveAndNew";
        private const string SERVICE_NAME = "first";
        private const string AIRPORT_TAX = "IsAirportTax";
        private const string SERVICE_CATEGORY = "ServiceCategoryFamilyId";

        // _________________________________________ Variables ________________________________________

        [FindsBy(How = How.Id, Using = SERVICE_NAME)]
        private IWebElement _servicename;

        [FindsBy(How = How.Id, Using = AIRPORT_TAX)]
        private IWebElement _airportTax;

        [FindsBy(How = How.Id, Using = SERVICE_CATEGORY)]
        private IWebElement _serviceCategory;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = ADD_NEW)]
        private IWebElement _addNew;

        // _________________________________________ Méthodes __________________________________________

        public ParametersCustomerCreateServiceModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public void SetServiceName(string serviceName)
        {
            _servicename = WaitForElementIsVisible(By.Id(SERVICE_NAME));
            _servicename.SetValue(ControlType.TextBox, serviceName);
        }

        public void SetFamily(string family)
        {
            _serviceCategory = WaitForElementIsVisible(By.Id(SERVICE_CATEGORY));
            _serviceCategory.SetValue(ControlType.DropDownList, family);
            _serviceCategory.SendKeys(Keys.Tab);
        }

        public void SetAirportTax()
        {
            _airportTax = WaitForElementExists(By.Id(AIRPORT_TAX));

            if (_airportTax.GetAttribute("checked") != "true")
                _airportTax.Click();

            _airportTax.SendKeys(Keys.PageDown);
        }

        public ParametersCustomer Save()
        {
            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();

            return new ParametersCustomer(_webDriver, _testContext);
        }

        public ParametersCustomer AddNew()
        {
            _addNew = WaitForElementIsVisible(By.Id(ADD_NEW));
            _addNew.Click();
            WaitForLoad();

            return new ParametersCustomer(_webDriver, _testContext);
        }

    }
}
