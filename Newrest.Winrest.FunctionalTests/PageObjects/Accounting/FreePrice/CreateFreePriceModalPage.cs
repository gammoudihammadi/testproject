using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Accounting.Invoice
{
    public class CreateFreePriceModalPage : PageBase
    {
        // ___________________________________ Constantes _______________________________________

        public const string NAME = "Name";
        public const string QUANTITY = "Quantity";
        public const string SELLING_PRICE = "SellingPrice";
        public const string WORKSHOP = "dropdown-workshops";
        public const string VALIDATE = "//*[@id=\"freePriceForm\"]/div[2]/button[2]";

        // ___________________________________ Variables ________________________________________

        [FindsBy(How = How.XPath, Using = VALIDATE)]
        private IWebElement _validate;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = QUANTITY)]
        private IWebElement _quantity;

        [FindsBy(How = How.Id, Using = SELLING_PRICE)]
        private IWebElement _sellingPrice;

        [FindsBy(How = How.Id, Using = WORKSHOP)]
        private IWebElement _workshop;

        // ___________________________________ Methodes _________________________________________

        public CreateFreePriceModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void FillField_CreatNewFreePrice(string serviceName, string quantity, string sellingPrice, string workshop)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            //_name.Click();
            _name.SetValue(ControlType.TextBox, serviceName);

            _quantity = WaitForElementIsVisible(By.Id(QUANTITY));
            //_quantity.Click();
            _quantity.SetValue(ControlType.TextBox, quantity);

            _sellingPrice = WaitForElementIsVisible(By.Id(SELLING_PRICE));
            //_sellingPrice.Click();
            _sellingPrice.SetValue(ControlType.TextBox, sellingPrice);

            _workshop = WaitForElementIsVisible(By.Id(WORKSHOP));
            //_workshop.Click();
            _workshop.SetValue(ControlType.DropDownList, workshop);
            _workshop.SendKeys(Keys.Tab);
        }

        public InvoiceDetailsPage ValidateForInvoice()
        {
            _validate = WaitForElementIsVisible(By.XPath(VALIDATE));
            _validate.Click();
            WaitForLoad();

            return new InvoiceDetailsPage(_webDriver, _testContext);
        }

        public CustomerOrderItem ValidateForCustomerOrder()
        {
            _validate = WaitForElementIsVisible(By.XPath(VALIDATE));
            _validate.Click();
            WaitForLoad();

            return new CustomerOrderItem(_webDriver, _testContext);
        }
    }
}
