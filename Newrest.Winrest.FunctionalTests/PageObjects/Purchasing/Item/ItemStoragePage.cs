using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Caching;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemStoragePage : PageBase
    {
        public ItemStoragePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string CHECK_SELECT = "//*/input[@class='select-item-detail-stock']";
        private const string SELECT_ALL = "//*/a[contains(@class,'usecase-select-all')]";
        private const string SELECT_NONE = "//*/a[contains(@class,'usecase-select-none')]";


        private const string RESET_FILTER = "reset-button";

        private const string SITE_FILTER = "SelectedSites_ms";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";
        private const string UNCHECK_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string FILTER_SITES = "//*[@id=\"table-itemDetailsStorage\"]/tbody/tr[*]/td[2]";

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL_SITES)]
        private IWebElement _uncheckAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        public enum FilterType
        {
            Site,
            StorageUnits
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Site:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.Click();

                    _uncheckAllSites = WaitForElementIsVisible(By.XPath(UNCHECK_ALL_SITES));
                    _uncheckAllSites.Click();

                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value);

                    var siteToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    siteToCheck.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    break;
                case FilterType.StorageUnits:

                    _siteFilter = WaitForElementIsVisible(By.Id("SelectedStorageUnits_ms"));
                    _siteFilter.Click();
                    _uncheckAllSites = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/ul/li[2]/a"));
                    _uncheckAllSites.Click();
                    _searchSite = WaitForElementIsVisible(By.XPath("/html/body/div[12]/div/div/label/input"));
                    _searchSite.SetValue(ControlType.TextBox, value);
                    var siteToCheck2 = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    siteToCheck2.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }
            WaitPageLoading();
        }

        public bool SelectAll()
        {
            var selectAll = WaitForElementIsVisible(By.XPath(SELECT_ALL));
            selectAll.Click();
            var coches = _webDriver.FindElements(By.XPath(CHECK_SELECT));
            foreach (var coche in coches)
            {
                if (!coche.Selected)
                {
                    return false;
                }
            }
            return true;
        }

        public bool SelectNone()
        {
            var selectNone = WaitForElementIsVisible(By.XPath(SELECT_NONE));
            selectNone.Click();
            var coches = _webDriver.FindElements(By.XPath(CHECK_SELECT));
            foreach (var coche in coches)
            {
                if (coche.Selected)
                {
                    return false;
                }
            }
            return true;
        }

        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool CheckAllSites()
        {
            var check = WaitForElementIsVisible(By.XPath("//*[@id=\"SelectedSites_ms\"]/span[2]"));
            //28 of 28 selected
            string [] text = check.Text.Split(' ');
            if (text.Length < 3)
            {
                return false;
            }
            if (text[0] != text[2])
            {
                return false;
            }
            return true;
        }

        public bool CheckAllStorageUnits()
        {
            var check = WaitForElementIsVisible(By.XPath("//*[@id=\"SelectedStorageUnits_ms\"]/span[2]"));
            //15 of 15 selected
            string[] text = check.Text.Split(' ');
            if (text.Length < 3)
            {
                return false;
            }
            if (text[0] != text[2])
            {
                return false;
            }
            return true;
        }
        public string GetFirstUnitFromLine()
        {
            var unitFirstLine = WaitForElementIsVisible(By.XPath("//*[@id=\"table-itemDetailsStorage\"]/tbody/tr[2]/td[5]"));
            return unitFirstLine.Text;
        }
        public List<string> GetSitesListFiltred()
        {
            return _webDriver.FindElements(By.XPath(FILTER_SITES)).Select(element => element.Text.Trim()).ToList();

        }
        public bool IsLinkToStorageOK()
        {
            if (isElementVisible(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div/div[2]/table/tbody/tr[2]")))
            {
                WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div[2]/div[2]/div/div/div[2]/table/tbody/tr[2]"));
            }
            else
            {
                return false;
            }
            return true;
        }
    }
}
