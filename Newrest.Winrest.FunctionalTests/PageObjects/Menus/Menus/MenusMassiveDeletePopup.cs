using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes.RecipeMassiveDeletePopup;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenusMassiveDeletePopup : PageBase
    {
        public MenusMassiveDeletePopup(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        private const string SITE_FILTER_ID = "siteFilter";
        private const string SITE_FILTER_SHOW_INACTIVE_ID = "ShowInactiveSites";
        private const string VARIANT_FILTER_ID = "SelectedVariants_ms";
        private const string RECIPETYPE_FILTER_ID = "/html/body/div[16]/div/ul/li[1]/a/span[2]";
        private const string STATUS_FILTER_ID = "SelectedStatus_ms";
        private const string SEARCH_BTN_ID = "SearchMenusBtn";
        private const string LIST_VARIANT = "//*[@id=\"tableMenus\"]/tbody/tr[*]/td[4]";

        public void SelectAllSites()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(SITE_FILTER_ID, "", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }

        public void SelectAllVariant()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(VARIANT_FILTER_ID, "", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }
        public void SelectAllStatus()
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(STATUS_FILTER_ID, "", false)
            { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true };
            ComboBoxSelectById(cbOpt);
        }
        public void SelectStatus(string statusLabel)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions()
            {
                XpathId = "collapseSelectedStatusFilter",
                SelectionValue = statusLabel,
                ClickCheckAllAtStart = false,
                ClickUncheckAllAtStart = false,
                IsUsedInFilter = false
            };
            ComboBoxSelectById(cbOpt);
        }       
        public bool VerifyIfNewWindowIsOpened()
        {

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            if (_webDriver.WindowHandles.Count == 1)
            {
                return false;
            }
            return true;
        }
        public void ClickOnSearch()
        {
            var searchBtn = WaitForElementExists(By.Id(SEARCH_BTN_ID));
            searchBtn.Click();
            WaitPageLoading();
        }
        public List<string> GetAllNamesResultPaged()
        {
            var names = new List<string>();

            var elements = _webDriver.FindElements(By.XPath(LIST_VARIANT));
            if (elements == null)
            {
                return new List<string>();
            }
            foreach (var element in elements)
            {
                names.Add(element.Text.Trim());
            }
            return names;
        }
        public MenusGeneralInformationPage ClickMenuLinkFromRow(int rowNumber)
        {
            var menuDetail = WaitForElementIsVisible(By.XPath("//*[@id=\"tableMenus\"]/tbody/tr["+ rowNumber + "]/td[7]/a"));
            menuDetail.Click();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitForLoad();
            return new MenusGeneralInformationPage(_webDriver, _testContext);
        }


        public void SelectVariantByName(string variantName, bool ignoreUncheckAll)
        {
            ComboBoxOptions cbOpt = new ComboBoxOptions(VARIANT_FILTER_ID, variantName, false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = !ignoreUncheckAll };
            ComboBoxSelectById(cbOpt);
        }      public class MenuMassiveDeleteRowResult
        {
            public string MenuName { get; set; }
            public bool IsMenuInactive { get; set; }
            public string SiteName { get; set; }
            public bool IsSiteInactive { get; set; }
            public string VariantName { get; set; }
            public double Weight { get; set; }
            public int UseCase { get; set; }
        }
 
    }
}
