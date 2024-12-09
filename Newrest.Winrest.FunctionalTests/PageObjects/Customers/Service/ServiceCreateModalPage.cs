using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceCreateModalPage : PageBase
    {

        public ServiceCreateModalPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        //________________________________________ Constantes ____________________________________________________________

        private const string SERVICE_NAME = "first";
        private const string SERVICE_CODE = "Code";
        private const string SERVICE_PRODUCTION = "ProductionName";
        private const string CATEGORY = "CategoryId";
        private const string GUEST_TYPE = "GuestTypeId";
        private const string SERVICE_TYPE = "ServiceTypeId";
        private const string CREATE = "last";

        private const string ERROR_NAME_REQUIRED = "//*[@id=\"service-filter-form\"]/div/div[1]/div[1]/div[1]/div[1]/div/span";
        private const string ACTIVATED = "//*[@id=\"IsActive\"]";

        private const string ISSPML = "//*[@id=\"IsSPML\"][@value='true']";
        private const string SPECIAL_MEAL = "SpecialMealId";
        //________________________________________ Variables _____________________________________________________________

        [FindsBy(How = How.Id, Using = SERVICE_NAME)]
        private IWebElement _serviceName;

        [FindsBy(How = How.Id, Using = SERVICE_CODE)]
        private IWebElement _serviceCode;
        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;
        [FindsBy(How = How.Id, Using = SERVICE_PRODUCTION)]
        private IWebElement _serviceProduction;

        [FindsBy(How = How.Id, Using = CATEGORY)]
        private IWebElement _category;

        [FindsBy(How = How.Id, Using = GUEST_TYPE)]
        private IWebElement _guestType;

        [FindsBy(How = How.Id, Using = SERVICE_TYPE)]
        private IWebElement _serviceType;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _create;

        [FindsBy(How = How.XPath, Using = ISSPML)]
        private IWebElement _isspml;

        [FindsBy(How = How.Id, Using = SPECIAL_MEAL)]
        private IWebElement _specialmeal;
        //_________________________________________ Méthodes _____________________________________________________________

        public void FillFields_CreateServiceModalPage(string serviceName, string serviceCode = null, string serviceProduction = null, string category = null, string guestType = null, string serviceType = null, string specialmeals = null)
        {
            WaitForLoad();
            // Renseigner le nom
            _serviceName = WaitForElementIsVisibleNew(By.Id(SERVICE_NAME));
            _serviceName.SetValue(ControlType.TextBox, serviceName);

            // Renseigner le code
            if (serviceCode != null)
            {
                _serviceCode = WaitForElementIsVisibleNew(By.Id(SERVICE_CODE));
                _serviceCode.SetValue(ControlType.TextBox, serviceCode);
            }

            if (serviceProduction != null)
            {
                // Renseigner le nom de production
                _serviceProduction = WaitForElementIsVisibleNew(By.Id(SERVICE_PRODUCTION));
                _serviceProduction.SetValue(ControlType.TextBox, serviceProduction);
            }

            // Renseigner la categorie
            if (category != null)
            {
                _category = WaitForElementIsVisibleNew(By.Id(CATEGORY));
                _category.SetValue(ControlType.DropDownList, category);
            }

            if(guestType != null)
            {
                _guestType = WaitForElementIsVisibleNew(By.Id(GUEST_TYPE));
                _guestType.SetValue(ControlType.DropDownList, guestType);
            }

            if(serviceType != null)
            {
                _serviceType = WaitForElementIsVisibleNew(By.Id(SERVICE_TYPE));
                _serviceType.SetValue(ControlType.DropDownList, serviceType);
            }

            WaitForLoad();
        }
        public void SetSPML(string specialmeals = null)
        {
            _isspml = WaitForElementExists(By.XPath(ISSPML));
            _isspml.SetValue(ControlType.RadioButton, true);
            WaitForLoad();
            if (specialmeals != null)
            {
                _specialmeal = WaitForElementIsVisible(By.Id(SPECIAL_MEAL));
                _specialmeal.SetValue(ControlType.DropDownList, specialmeals);
            }
        }


        public void FillFields_CreateServiceModalPageWithDesactivatedMode(string serviceName, string serviceCode = null, string serviceProduction = null, string category = null, string guestType = null, string serviceType = null)
        {
            WaitForLoad();
            // Renseigner le nom
            _serviceName = WaitForElementIsVisible(By.Id(SERVICE_NAME));
            _serviceName.SetValue(ControlType.TextBox, serviceName);

            // Renseigner le code
            if (serviceCode != null)
            {
                _serviceCode = WaitForElementIsVisible(By.Id(SERVICE_CODE));
                _serviceCode.SetValue(ControlType.TextBox, serviceCode);
            }

            if (serviceProduction != null)
            {
                // Renseigner le nom de production
                _serviceProduction = WaitForElementIsVisible(By.Id(SERVICE_PRODUCTION));
                _serviceProduction.SetValue(ControlType.TextBox, serviceProduction);
            }

            // Renseigner la categorie
            if (category != null)
            {
                _category = WaitForElementIsVisible(By.Id(CATEGORY));
                _category.SetValue(ControlType.DropDownList, category);
            }

            if (guestType != null)
            {
                _guestType = WaitForElementIsVisible(By.Id(GUEST_TYPE));
                _guestType.SetValue(ControlType.DropDownList, guestType);
            }

            if (serviceType != null)
            {
                _serviceType = WaitForElementIsVisible(By.Id(SERVICE_TYPE));
                _serviceType.SetValue(ControlType.DropDownList, serviceType);
            }
            var valueToSite = WaitForElementIsVisible(By.XPath(ACTIVATED));
            valueToSite.SetValue(ControlType.CheckBox, false);
            WaitForLoad();

            WaitForLoad();
        }

        public ServiceGeneralInformationPage Create()
        {
            _create = WaitForElementToBeClickable(By.Id(CREATE));
            _create.Click();
            WaitPageLoading();
            WaitForLoad();

            return new ServiceGeneralInformationPage(_webDriver, _testContext);
        }

        public bool GetError()
        {
            if(isElementVisible(By.XPath(ERROR_NAME_REQUIRED)))
            {
                WaitForElementIsVisible(By.XPath(ERROR_NAME_REQUIRED));
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
