using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerCreateModalPage : PageBase
    {

        public CustomerCreateModalPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        //_______________________________________________________Constantes____________________________________________________________

        private const string CUSTOMER_NAME = "Customer_Name";
        private const string CUSTOMER_CODE = "Customer_Code";
        private const string CUSTOMER_TYPE = "Customer_CustomerTypeId";
        private const string CUSTOMER_CURRENCY = "Customer_CurrencyId";
        private const string CUSTOMER_SHIPPING = "shipping-duration";
        private const string CUSTOMER_PROCESSING = "processing-duration";
        private const string CUSTOMER_LAST_UPDATE = "last-update";
        private const string OK_BUTTON = "dataAlertCancel";




        private const string SAVE_BTN = "//*[@id=\"customer-filter-form\"]/div[2]/div/button[2]";
        private const string NAME_ERROR = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[1]/div[1]/div/span/span";
        private const string ICAO_ERROR = "//*[@id=\"customer-filter-form\"]/div[1]/div/div[1]/div[1]/div[2]/div/span/span";
        private const string INACTIVE = "/html/body/div[3]/div/div/div[2]/div/form/div[1]/div/div[1]/div[2]/div[1]/div/div";

        //_______________________________________________________Variables_____________________________________________________________

        [FindsBy(How = How.Id, Using = CUSTOMER_NAME)]
        private IWebElement _customerName;

        [FindsBy(How = How.Id, Using = CUSTOMER_CODE)]
        private IWebElement _icao;

        [FindsBy(How = How.Id, Using = CUSTOMER_TYPE)]
        private IWebElement _customerType;

        [FindsBy(How = How.Id, Using = CUSTOMER_CURRENCY)]
        private IWebElement _customerCurrency;

        [FindsBy(How = How.Id, Using = OK_BUTTON)]
        private IWebElement _OkButton;

        [FindsBy(How = How.XPath, Using = SAVE_BTN)]
        private IWebElement _create;


        //_______________________________________________________Pages_________________________________________________________________

        public void FillFields_CreateCustomerModalPage(string customerName, string customerCode, string customerType)
        {
            // Renseigner le nom
            _customerName = WaitForElementIsVisible(By.XPath("//input[@id=\"Customer_Name\"]"));
            _customerName.SetValue(ControlType.TextBox, customerName);

            // Renseigner le ICAO
            _icao = WaitForElementIsVisible(By.XPath("//input[@id=\"Customer_Code\"]"));
            _icao.SetValue(ControlType.TextBox, customerCode);

            // Renseigner le type de client
            _customerType = WaitForElementIsVisible(By.Id(CUSTOMER_TYPE));
            _customerType.Click();
            _customerType.SetValue(ControlType.DropDownList, customerType);

            var shippingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_SHIPPING));
            shippingDuration.SetValue(ControlType.TextBox, "3");
            var processingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_PROCESSING));
            processingDuration.SetValue(ControlType.TextBox, "3");
            var lastUpdate = WaitForElementIsVisible(By.Id(CUSTOMER_LAST_UPDATE));
            lastUpdate.SetValue(ControlType.TextBox, "3");
        }

        public void FillFieldsWithCurrency_CreateCustomerModalPage(string customerName, string customerCode, string customerType , string customerCurrency)
        {
            // Renseigner le nom
            _customerName = WaitForElementIsVisible(By.Id(CUSTOMER_NAME));
            _customerName.SetValue(ControlType.TextBox, customerName);

            // Renseigner le ICAO
            _icao = WaitForElementIsVisible(By.Id(CUSTOMER_CODE));
            _icao.SetValue(ControlType.TextBox, customerCode);

            // Renseigner le type de client
            _customerType = WaitForElementIsVisible(By.Id(CUSTOMER_TYPE));
            _customerType.Click();
            _customerType.SetValue(ControlType.DropDownList, customerType);

            _customerCurrency = WaitForElementIsVisible(By.Id(CUSTOMER_CURRENCY));
            _customerCurrency.Click();
            _customerCurrency.SetValue(ControlType.DropDownList, customerCurrency);

            _OkButton = WaitForElementIsVisible(By.Id(OK_BUTTON));
            _OkButton.Click();

            var shippingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_SHIPPING));
            shippingDuration.SetValue(ControlType.TextBox, "3");
            var processingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_PROCESSING));
            processingDuration.SetValue(ControlType.TextBox, "3");
            var lastUpdate = WaitForElementIsVisible(By.Id(CUSTOMER_LAST_UPDATE));
            lastUpdate.SetValue(ControlType.TextBox, "3");
        }


        public CustomerGeneralInformationPage Create()
        {
            _create = WaitForElementIsVisible(By.XPath(SAVE_BTN));
            _create.Click();
            WaitPageLoading();

            return new CustomerGeneralInformationPage(_webDriver, _testContext);
        }

        public bool IsNameErrorDisplayed()
        {
            if(isElementVisible(By.XPath(NAME_ERROR)))
            {
                _webDriver.FindElement(By.XPath(NAME_ERROR));
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsIcaoErrorDisplayed()
        {
            if (isElementVisible(By.XPath(ICAO_ERROR)))
            {
                _webDriver.FindElement(By.XPath(ICAO_ERROR));
                return true;
            }
            else
            {
                return false;
            }
        }
        public void FillFields_CreateCustomerInactiveModalPage(string customerName, string customerCode, string customerType)
        {
            // Renseigner le nom
            _customerName = WaitForElementIsVisible(By.XPath("//input[@id=\"Customer_Name\"]"));
            _customerName.SetValue(ControlType.TextBox, customerName);

            // Renseigner le ICAO
            _icao = WaitForElementIsVisible(By.XPath("//input[@id=\"Customer_Code\"]"));
            _icao.SetValue(ControlType.TextBox, customerCode);

            // Renseigner le type de client
            _customerType = WaitForElementIsVisible(By.Id(CUSTOMER_TYPE));
            _customerType.Click();
            _customerType.SetValue(ControlType.DropDownList, customerType);

            var shippingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_SHIPPING));
            shippingDuration.SetValue(ControlType.TextBox, "3");
            var processingDuration = WaitForElementIsVisible(By.Id(CUSTOMER_PROCESSING));
            processingDuration.SetValue(ControlType.TextBox, "3");
            var lastUpdate = WaitForElementIsVisible(By.Id(CUSTOMER_LAST_UPDATE));
            lastUpdate.SetValue(ControlType.TextBox, "3");

            var inactive = WaitForElementIsVisible(By.XPath(INACTIVE));
            inactive.Click();

        }
    }
}
