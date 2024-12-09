using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Suppliers
{
    public class SupplierGeneralInfoTab : ItemGeneralInformationPage
    {
        public SupplierGeneralInfoTab(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        private const string SELECT_SITES = "SelectedSites_ms";
        private const string SELECT_ALL_SITES = "/html/body/div[11]/div/ul/li[1]/a/span[2]";
        private const string SUPPLIER_TYPE = "drop-down-suppliertypes";
        private const string GO_TO_DELIVERIES = "hrefTabContentDelivery";
        private const string AUTO_CLOSE_CHECKBOX = "/html/body/div[3]/div/div/div/div/div/form/div/div[8]/div[2]/div/div[2]/div/input"; //"IsAutoClose";
        private const string TAB_ITEM = "hrefTabContentArticles"; 





        //__________________________________Variables________________________________________

        [FindsBy(How = How.XPath, Using = AUTO_CLOSE_CHECKBOX)]
        private IWebElement _autoCloseCheckbox;

        [FindsBy(How = How.Id, Using = SELECT_SITES)]
        private IWebElement _selectSite;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_SITES)]
        private IWebElement _selectAllSites;

        [FindsBy(How = How.Id, Using = SUPPLIER_TYPE)]
        private IWebElement _supplierType;

        [FindsBy(How = How.Id, Using = GO_TO_DELIVERIES)]
        private IWebElement _goToDeliveries;

        public void SetName(string random_supplier_name)
        {
            var _supplierName = WaitForElementIsVisible(By.Id("Name"));
            _supplierName.SetValue(ControlType.TextBox, random_supplier_name);
            WaitPageLoading();
        }
        public IEnumerable<string> GetSites()
        {
            List<string> sites = new List<string>();

            var selectSitesBtn = WaitForElementIsVisible(By.Id(SELECT_SITES));
            selectSitesBtn.Click();

            var listSites = _webDriver.FindElements(By.XPath("//input[@type=\"checkbox\" and @name=\"multiselect_SelectedSites\" and @aria-selected]/../span"));

            foreach (var site in listSites)
            {
                sites.Add(site.Text.Substring(0, 3));
            }
            return sites;
        }

        public void SetSite()
        {
             _selectSite = WaitForElementIsVisible(By.Id(SELECT_SITES));
            _selectSite.Click();
            WaitPageLoading();
            _selectAllSites = WaitForElementIsVisible(By.XPath(SELECT_ALL_SITES));
            _selectAllSites.Click();
            WaitPageLoading();


        }

        public void SetSupplierType(string supplirType)
        {
            _supplierType = WaitForElementIsVisible(By.Id(SUPPLIER_TYPE));
            _supplierType.SetValue(ControlType.DropDownList, supplirType);
            WaitPageLoading();
        }

        public SupplierDeliveriesTab GoToSupplierDeliveries()
        {
            _goToDeliveries = WaitForElementIsVisible(By.Id(GO_TO_DELIVERIES));
            _goToDeliveries.Click();
            WaitPageLoading();
            return new SupplierDeliveriesTab(_webDriver, _testContext);
        }
        public void CheckAutoCloseRecipeNote()
        {
            _autoCloseCheckbox = WaitForElementIsVisible(By.XPath(AUTO_CLOSE_CHECKBOX));
            _autoCloseCheckbox.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }
        public SupplierItem ClickOnItemTab()
        {
            var itemTab = WaitForElementIsVisible(By.Id(TAB_ITEM));
            itemTab.Click();
            WaitForLoad();
            return new SupplierItem(_webDriver, _testContext);
        }
        public void SetSupplierInactive()
        {
            try
            {
                ShowExtendedMenu(); 
                var desactivateSupplier = WaitForElementExists(By.Id("btn-deactivate"));
                desactivateSupplier.Click();
                WaitPageLoading();
            }
            catch
            {
                // déjà à No Active
            }
            WaitForLoad();
        }
    }
}
