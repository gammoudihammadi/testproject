using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery
{
    public class DeliveryCreateModalPage : PageBase
    {

        public DeliveryCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________________ Constantes____________________________________________________________

        private const string DELIVERY_NAME = "first";
        private const string DELIVERY_CUSTOMER = "CustomerId";
        private const string DELIVERY_SITE = "SiteId";
        private const string IS_ACTIVATED = "check-box-isactive";

        private const string SAVE_BTN = "last";

        private const string ERROR_NAME_REQUIRED = "//*[@id=\"flightdelivery-filter-form\"]/div/div/div[1]/div/div[1]/div/span";

        //_______________________________________ Variables_____________________________________________________________

        [FindsBy(How = How.Id, Using = DELIVERY_NAME)]
        private IWebElement _deliveryName;

        [FindsBy(How = How.Id, Using = DELIVERY_CUSTOMER)]
        private IWebElement _deliveryCustomer;

        [FindsBy(How = How.Id, Using = DELIVERY_SITE)]
        private IWebElement _deliverySite;

        [FindsBy(How = How.Id, Using = IS_ACTIVATED)]
        private IWebElement _isActivatedBtn;

        [FindsBy(How = How.XPath, Using = SAVE_BTN)]
        private IWebElement _create;

        [FindsBy(How = How.XPath, Using = ERROR_NAME_REQUIRED)]
        private IWebElement _error;

        //_________________________________________ Méthodes _______________________________________________________________

        public void FillFields_CreateDeliveryModalPage(string deliveryName, string deliveryCustomer, string deliverySite, bool isActivated)
        {
            // Renseigner le nom
            _deliveryName = WaitForElementIsVisible(By.Id(DELIVERY_NAME));
            _deliveryName.SetValue(ControlType.TextBox, deliveryName);
            WaitForLoad();

            // Renseigner le  site
            _deliverySite = WaitForElementIsVisible(By.Id(DELIVERY_SITE));
            _deliverySite.SetValue(ControlType.DropDownList, deliverySite);
            WaitForLoad();

            // Renseigner le customer
            _deliveryCustomer = WaitForElementIsVisible(By.Id(DELIVERY_CUSTOMER));
            _deliveryCustomer.SetValue(ControlType.DropDownList, deliveryCustomer);
            WaitForLoad();

            if (!isActivated)
            {
                _isActivatedBtn = WaitForElementExists(By.Id(IS_ACTIVATED));
                _isActivatedBtn.Click();
                WaitForLoad();
            }
        }

        public DeliveryLoadingPage Create()
        {
            _create = WaitForElementIsVisible(By.Id(SAVE_BTN));
            _create.Click();
            WaitForLoad();

            return new DeliveryLoadingPage(_webDriver, _testContext);
        }

        public string GetErrorMessage()
        {
            _error = WaitForElementIsVisible(By.XPath(ERROR_NAME_REQUIRED));
            return _error.Text;
        }
    }
}
