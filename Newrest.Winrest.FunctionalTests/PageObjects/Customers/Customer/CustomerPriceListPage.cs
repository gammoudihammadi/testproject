using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerPriceListPage : PageBase
    {

        public CustomerPriceListPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _______________________________________________

        // General
        private const string NEW_PRICE = "//*[@id=\"tabContentDetails\"]/div/div[2]/div/div/a[2]";
        private const string PRICE_NAME = "drop-down-vip-prices";
        private const string PRICE_SITE = "//*[@id=\"drop-down-sites\"]/option[text()='{0}']";
        private const string CREATE = "//*[@id=\"form-add-vip-price\"]/div[3]/button[2]";

        // Tableau
        private const string FIRST_PRICE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div";
        private const string FIRST_PRICE_SITE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[2]/span";
        private const string FIRST_PRICE_NAME = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div/div[3]/span";
        private const string EDIT_PRICE = "//span[@class=\"glyphicon glyphicon-pencil\"]";
        private const string DELETE_PRICE = "//span[@class=\"glyphicon glyphicon-trash glyph-span\"]";
        private const string CONFIRM_DELETE = "first";

        //__________________________________ Variables _________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_PRICE)]
        private IWebElement _newPrice;

        [FindsBy(How = How.Id, Using = PRICE_NAME)]
        private IWebElement _priceName;

        [FindsBy(How = How.XPath, Using = PRICE_SITE)]
        private IWebElement _priceSite;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;

        // Tableau
        [FindsBy(How = How.XPath, Using = EDIT_PRICE)]
        private IWebElement _editPrice;

        [FindsBy(How = How.XPath, Using = DELETE_PRICE)]
        private IWebElement _deletePrice;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = FIRST_PRICE)]
        private IWebElement _firstPrice;

        [FindsBy(How = How.XPath, Using = FIRST_PRICE_SITE)]
        private IWebElement _firstPriceSite;

        [FindsBy(How = How.XPath, Using = FIRST_PRICE_NAME)]
        private IWebElement _firstPriceName;

        //___________________________________ Methodes ___________________________________________________

        public void CreateNewPrice(string priceName, string site)
        {
            _newPrice = WaitForElementIsVisible(By.XPath(NEW_PRICE));
            _newPrice.Click();
            WaitForLoad();

            // Renseigner le site
            _priceSite = WaitForElementIsVisible(By.XPath(String.Format(PRICE_SITE, site)));
            _priceSite.Click();
            WaitForLoad();

            // Renseigner le nom
            _priceName = WaitForElementIsVisible(By.Id(PRICE_NAME));
            _priceName.Click();
            _priceName.SetValue(ControlType.DropDownList, priceName);
            WaitForLoad();

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            //lenteur dans la création
            WaitPageLoading();
        }

        public PriceListDetailsPage EditPriceList()
        {
            _editPrice = WaitForElementIsVisible(By.XPath("//span[contains(@class,'pencil')]"));
            _editPrice.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PriceListDetailsPage(_webDriver, _testContext);
        }

        public string GetFirstPriceSite()
        {
            _firstPriceSite = WaitForElementIsVisible(By.XPath(FIRST_PRICE_SITE));
            return _firstPriceSite.Text;
        }

        public string GetFirstPriceName()
        {
            _firstPriceName = WaitForElementIsVisible(By.XPath(FIRST_PRICE_NAME));
            return _firstPriceName.Text;
        }

        public bool IsPriceDisplayed()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath(FIRST_PRICE)))
            {
                _firstPrice = _webDriver.FindElement(By.XPath(FIRST_PRICE));
                return _firstPrice.Displayed;
            }
            else
            {
                return false;
            }
        }

        public void DeletePriceList()
        {
            _deletePrice = WaitForElementIsVisible(By.XPath("//span[contains(@class,'trash')]"));
            _deletePrice.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }
    }
}
