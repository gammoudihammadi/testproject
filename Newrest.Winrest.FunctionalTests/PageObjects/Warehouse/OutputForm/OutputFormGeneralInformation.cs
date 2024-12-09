using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.OutputForm
{
    public class OutputFormGeneralInformation : PageBase
    {
        public OutputFormGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________ Constantes ______________________________________

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";

        // Informations
        private const string DATE = "datapicker-output-form-date";
        private const string SITE = "OutputForm_FromSitePlace_Site_Name";
        private const string PLACE_FROM = "OutputForm_FromSitePlace_Title";
        private const string PLACE_TO = "OutputForm_ToSitePlace_Title";
        private const string COMMENT = "OutputForm_Comment";
        public const string INVENTORY_NUMBER = "tb-new-outputform-number";
        public const string SUPPLY_ORDER_NUMBER = "//*[@id='form-create-output-form']/div[2]/div/div/div/span/a";
        public const string SUPPLY_ORDER_MESSAGE = "//*[@id='form-create-output-form']/div[2]/div/div/div/span";

        //_______________________________ Variables _______________________________________

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        // Informations
        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = PLACE_FROM)]
        private IWebElement _placeFrom;

        [FindsBy(How = How.Id, Using = PLACE_TO)]
        private IWebElement _placeTo;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comments;

        [FindsBy(How = How.Id, Using = INVENTORY_NUMBER)]
        private IWebElement _inventoryNumber;

        [FindsBy(How = How.Id, Using = SUPPLY_ORDER_NUMBER)]
        private IWebElement _supplyOrderNumber;

        [FindsBy(How = How.Id, Using = SUPPLY_ORDER_MESSAGE)]
        private IWebElement _supplyOrderMessage;
        


        //_______________________________ Méthodes ________________________________________

        // Onglets
        public OutputFormItem ClickOnItemsTab()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new OutputFormItem(_webDriver, _testContext);
        }

        // Informations
        public void UpdateInformations(string comments, DateTime date)
        {
            _comments = WaitForElementIsVisible(By.Id(COMMENT));
            _comments.SetValue(ControlType.TextBox, comments);
            _comments.SendKeys(Keys.Tab);

            _date = WaitForElementIsVisible(By.Id(DATE));
            _date.SetValue(ControlType.DateTime, date);
            _comments.SendKeys(Keys.Tab);

            WaitPageLoading();
        }

        public string GetDate()
        {
            _date = WaitForElementIsVisible(By.Id(DATE));
            return _date.GetAttribute("value");
        }

        public string GetComments()
        {
            _comments = WaitForElementIsVisible(By.Id(COMMENT));
            return _comments.GetAttribute("value");
        }

        public string GetInventoryNumber()
        {
            _inventoryNumber = WaitForElementIsVisible(By.Id(INVENTORY_NUMBER));
            return _inventoryNumber.GetAttribute("value");
        }

        public string GetSupplyOrderNumber()
        {
            _supplyOrderNumber = WaitForElementIsVisible(By.XPath(SUPPLY_ORDER_NUMBER));
            return _supplyOrderNumber.Text.Trim();
        }

        public string GetSupplyOrderMessage()
        {
            _supplyOrderMessage = WaitForElementIsVisibleNew(By.XPath(SUPPLY_ORDER_MESSAGE));
            return _supplyOrderMessage.Text.Trim();
        }

        public string GetSite()
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            return _site.GetAttribute("value");
        }

        public string GetPlaceFrom()
        {
            _placeFrom = WaitForElementIsVisible(By.Id(PLACE_FROM));
            return _placeFrom.GetAttribute("value");
        }

        public string GetPlaceTo()
        {
            _placeTo = WaitForElementIsVisible(By.Id(PLACE_TO));
            return _placeTo.GetAttribute("value");
        }
        public ProductionCOPage ClickNumber()
        {
            _supplyOrderNumber = WaitForElementIsVisible(By.XPath(SUPPLY_ORDER_NUMBER));
            _supplyOrderNumber.Click();
            WaitPageLoading();
            WaitForLoad();//pas suffisant
            return new ProductionCOPage(_webDriver, _testContext); ;
        }
    }
}
