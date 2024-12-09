using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System; 

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory
{
    public class InventoryCreateModalPage : PageBase
    {
        public InventoryCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //________________________________________________Constantes___________________________________________________


        private const string INVENTORY_NUMBER = "tb-new-inventory-number";
        private const string DATE = "datepicker-new-inventory";
        private const string SITE = "SelectedSiteId";
        private const string PLACE = "SelectedSitePlaceId";
        private const string ACTIVATED = "Inventory_IsActive";

        private const string SUBMIT = "btn-submit-create-inventory";
        private const string SUBMIT_MOBILE = "btn-submit-mobile-inventory";

        //________________________________________________Variables_____________________________________________________

        [FindsBy(How = How.Id, Using = INVENTORY_NUMBER)]
        private IWebElement _inventoryNumber;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = PLACE)]
        private IWebElement _place;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        [FindsBy(How = How.Id, Using = SUBMIT_MOBILE)]
        private IWebElement _submitMobile;

        //________________________________________________ Méthodes _________________________________________________________

        public void FillField_CreateNewInventory(DateTime date, string site, string place, bool isActive = true)
        {

            _date = WaitForElementExists(By.Id(DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _site = WaitForElementExists(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            //Attente du populate de la dropdown des place en fonction du site qu'on vient de choisir
            _place = WaitForElementExists(By.Id(PLACE));
            _place.SetValue(ControlType.DropDownList, place);

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActive);
            WaitForLoad();
        }

        public string GetInventoryNumber()
        {
            _inventoryNumber = WaitForElementExists(By.Id(INVENTORY_NUMBER));
            return _inventoryNumber.GetAttribute("value");
        }

        public InventoryItem Submit()
        {
            _submit = WaitForElementIsVisibleNew(By.Id(SUBMIT));
            _submit.Click();
            WaitPageLoading();

            return new InventoryItem(_webDriver, _testContext);
        }

        public InventoryGeneralInformation SubmitMobile()
        {
            _submitMobile = WaitForElementIsVisible(By.Id(SUBMIT_MOBILE));
            _submitMobile.Click();
            WaitForLoad();

            return new InventoryGeneralInformation(_webDriver, _testContext);
        }

    }
}
