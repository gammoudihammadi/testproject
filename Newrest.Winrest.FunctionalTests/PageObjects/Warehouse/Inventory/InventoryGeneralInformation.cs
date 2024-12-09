using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory
{
    public class InventoryGeneralInformation : PageBase
    {
        public InventoryGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_____________________________ Constantes ______________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string VALIDATE = "btn-validate-inventory";

        // Onglets
        private const string ITEMS = "hrefTabContentItems";

        // Informations
        private const string INVENTORY_NUMBER = "tb-new-inventory-number";
        private const string INVENTORY_SITE = "Inventory_SitePlace_Site_Code";
        private const string INVENTORY_PLACE = "Inventory_SitePlace_Title";
        private const string ACTIVATED = "Inventory_IsActive";

        //_____________________________ Variables ______________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = VALIDATE)]
        private IWebElement _validate;


        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS)]
        private IWebElement _items;

        // Informations
        [FindsBy(How = How.Id, Using = INVENTORY_NUMBER)]
        private IWebElement _inventoryNumber;

        [FindsBy(How = How.Id, Using = INVENTORY_SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = INVENTORY_PLACE)]
        private IWebElement _place;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        //____________________________ Méthodes ________________________________________

        // General
        public InventoriesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new InventoriesPage(_webDriver, _testContext);
        }

        // Onglets
        public InventoryItem ClickOnItems()
        {
            _items = WaitForElementIsVisible(By.Id(ITEMS));
            _items.Click();
            WaitForLoad();

            return new InventoryItem(_webDriver, _testContext);
        }

        // Informations
        public string GetInventoryNumber()
        {
            _inventoryNumber = WaitForElementExists(By.Id(INVENTORY_NUMBER));
            return _inventoryNumber.GetAttribute("value");
        }

        public string GetInventorySite()
        {
            _site = WaitForElementExists(By.Id(INVENTORY_SITE));
            return _site.GetAttribute("value");
        }

        public string GetInventoryPlace()
        {
            _place = WaitForElementExists(By.Id(INVENTORY_PLACE));
            return _place.GetAttribute("value");
        }

        public bool CanValidate()
        {
            ShowValidationMenu();

            _validate = WaitForElementExists(By.Id(VALIDATE));
            return _validate.GetAttribute("disabled") == null;
        }

        public void ChangeStatusInventory(bool activated)
        {
            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, activated);
            WaitForLoad();

            Thread.Sleep(2000);
        }

        public void SetComment(string comment)
        {
            var _commentInput = WaitForElementIsVisible(By.Id("Inventory_Comment"));
            _commentInput.Clear();
            _commentInput.SendKeys(comment);
            Thread.Sleep(2000);
        }

        public string GetComment()
        {
            var _commentInput = WaitForElementIsVisible(By.Id("Inventory_Comment"));
            return _commentInput.Text;
        }
    }
}
