using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Reinvoice
{
    public class ReinvoiceCreateModalPage : PageBase
    {
        public ReinvoiceCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ________________________________ Constantes ______________________________________________

        private const string CUSTOMER_FROM = "customer-from-input";
        private const string FIRST_CUSTOMER_FROM = "//div[contains(text(), '{0}')]";
        private const string SERVICE_FROM = "ServiceFromId";
        private const string CUSTOMER_TO = "customer-to-input";
        private const string FIRST_CUSTOMER_TO = "//div[contains(text(), '{0}')]";
        private const string SITE_TO = "SiteToId";
        private const string SERVICE_TO = "ServiceToId";
        private const string CREATE = "create";

        // ________________________________ Variables _________________________________________________

        [FindsBy(How = How.Id, Using = CUSTOMER_FROM)]
        private IWebElement _customerFrom;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_FROM)]
        private IWebElement _firstCustomerFrom;

        [FindsBy(How = How.Id, Using = SERVICE_FROM)]
        private IWebElement _serviceFrom;

        [FindsBy(How = How.Id, Using = CUSTOMER_TO)]
        private IWebElement _customerTo;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_TO)]
        private IWebElement _firstCustomerTo;

        [FindsBy(How = How.Id, Using = SITE_TO)]
        private IWebElement _siteTo;

        [FindsBy(How = How.Id, Using = SERVICE_TO)]
        private IWebElement _serviceTo;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _create;

        // ________________________________ Méthodes _____________________________________________________



        public void FillFields_CreateReinvoiceModalPage(string customerFromName, string customerToName, string siteTo = null)
        {
            // Définition du nom customerFrom
            _customerFrom = WaitForElementIsVisible(By.Id(CUSTOMER_FROM));
            _customerFrom.SetValue(ControlType.TextBox, customerFromName);
            _firstCustomerFrom = WaitForElementIsVisible(By.XPath(string.Format(FIRST_CUSTOMER_FROM, customerFromName)));
            _firstCustomerFrom.Click();
            WaitForLoad();

            // Définition du nom customerTO
            _customerTo = WaitForElementIsVisible(By.Id(CUSTOMER_TO));
            _customerTo.SetValue(ControlType.TextBox, customerToName);
            _firstCustomerTo = WaitForElementIsVisible(By.XPath(string.Format(FIRST_CUSTOMER_TO, customerToName)));
            _firstCustomerTo.Click();
            WaitForLoad();

            // Définition du nom site TO
            if (siteTo != null)
            {
                _siteTo = WaitForElementIsVisible(By.Id(SITE_TO));
                _siteTo.SetValue(ControlType.TextBox, siteTo);
                _siteTo.SendKeys(Keys.Tab);
                WaitForLoad();
            }
        }

        public ReinvoicePage Create()
        {
            // Click sur le bouton "Create"
            _create = WaitForElementIsVisible(By.Id(CREATE));
            _create.Click();
            WaitPageLoading();

            return new ReinvoicePage(_webDriver, _testContext);
        }


        public bool IsServiceFromSelected()
        {
            _serviceFrom = WaitForElementExists(By.Id(SERVICE_FROM));

            if (_serviceFrom.Text.Equals(""))
            {
                return false;
            }

            return true;
        }

        public bool IsServiceToSelected()
        {
            _serviceTo = WaitForElementExists(By.Id(SERVICE_TO));
            if (_serviceTo.Text.Equals(""))
            {
                return false;
            }

            return true;
        }


        public bool IsOkToCreate()
        {
            _create = WaitForElementExists(By.Id(CREATE));
            return _create.Enabled;
        }

        public void CloseReinvoiceCreateModal()
        {
            var close = WaitForElementIsVisible(By.XPath("//button[contains(text(), 'Close')]"));
            close.Click();
        }
    }
}