using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList
{
    public class PriceListPage : PageBase
    {
        public PriceListPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _______________________________________

        // Général
        private const string PLUS_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/button";
        private const string CREATE_PRICING = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a";

        // Tableau
        private const string PRICING = "//*[@id=\"list-item-with-action\"]/div/div[*]/div[1]/div/div[2]/table/tbody/tr/td[contains(text(),'{0}')]";
        private const string PRICING_NAME = "//*[@id=\"list-item-with-action\"]/div/div[*]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string PRICING_SITE = "//*[@id=\"list-item-with-action\"]/div/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]";

        private const string DELETE_PRICING = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[1]/div/div[2]/table/tbody/tr/td[contains(text(), '{0}')]/../../../../../div[3]/a/span[@class='glyphicon glyphicon-trash']";
        private const string DELETE_PRICING_DEV = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[1]/div/div[2]/table/tbody/tr/td[contains(text(), '{0}')]/../../../../../div[3]/a";
        private const string DELETE_PERIOD = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[4]/a";
        private const string CONFIRM_DELETE = "dataConfirmOK";

        private const string UNFOLD = "//*[@id=\"list-item-with-action\"]/div/div[*]/div[1]/div/div[1]/span";
        private const string CONTENT = "//*[starts-with(@id,\"content_\")]";
        private const string PERIOD = "//*[@id=\"tableListMenu\"]/tbody/tr[*]/td[2]";

        // Filtres
        private const string RESET_FILTER = "//a[contains(text(), 'Reset Filter')]";

        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string SITE_FILTER = "cbSites";
        private const string PRLIST_IS_LOCAL = "//span[text()='Local' and (contains(@class, \"primary\"))]";

        //__________________________________Variables_______________________________________

        // Général
        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;

        [FindsBy(How = How.XPath, Using = CREATE_PRICING)]
        private IWebElement _createNewPricing;

        // Tableau
        [FindsBy(How = How.XPath, Using = PRICING)]
        private IWebElement _pricing;

        [FindsBy(How = How.XPath, Using = PRICING_NAME)]
        private IWebElement _firstPricingName;

        [FindsBy(How = How.XPath, Using = DELETE_PRICING)]
        private IWebElement _deleteBtn;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDeleteBtn;

        [FindsBy(How = How.XPath, Using = UNFOLD)]
        private IWebElement _unfold;

        [FindsBy(How = How.XPath, Using = DELETE_PERIOD)]
        private IWebElement _deletePeriod;

        // _________________________________ Filtres _______________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByFilter;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;


        public enum FilterType
        {
            Search,
            SortBy,
            Site
        }

        public void Filter(FilterType filterType, object value)
        {

            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortByFilter = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.SetValue(ControlType.DropDownList, value);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            Thread.Sleep(2500);
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            
            WaitForLoad();

            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool VerifySite(string site)
        {
            var sites = _webDriver.FindElements(By.XPath(PRICING_SITE));

            foreach (var elm in sites)
            {
                if (!elm.Text.Equals(site))
                    return false;
            }

            return true;
        }

        public bool IsSortedBySite()
        {
            var ancienSite = "";
            int compteur = 1;

            var sites = _webDriver.FindElements(By.XPath(PRICING_SITE));

            if (sites.Count == 0)
                return false;

            foreach (var elm in sites)
            {
                if (compteur == 1)
                    ancienSite = elm.Text;

                if (String.Compare(ancienSite, elm.Text) > 0)
                    return false;

                ancienSite = elm.Text;
                compteur++;
            }

            return true;
        }

        public bool IsSortedByName()
        {
            var ancienName = "";
            int compteur = 1;

            var names = _webDriver.FindElements(By.XPath(PRICING_NAME));

            if (names.Count == 0)
                return false;

            foreach (var elm in names)
            {
                if (compteur == 1)
                    ancienName = elm.Text;

                if (String.Compare(ancienName, elm.Text) > 0)
                    return false;

                ancienName = elm.Text;
                compteur++;
            }

            return true;
        }


        // _____________________________________ Méthodes ________________________________________________

        // Général
        public override void ShowPlusMenu()
        {
            var actions = new Actions(_webDriver);
            _plusBtn = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            actions.MoveToElement(_plusBtn).Perform();
        }

        public PricingCreateModalpage CreateNewPricing()
        {
            ShowPlusMenu();

            _createNewPricing = WaitForElementIsVisible(By.XPath(CREATE_PRICING));
            _createNewPricing.Click();
            WaitForLoad();

            return new PricingCreateModalpage(_webDriver, _testContext);
        }

        // Tableau

        public PriceListDetailsPage ClickOnFirstPricing()
        {
            _firstPricingName = WaitForElementIsVisible(By.XPath(PRICING_NAME));
            _firstPricingName.Click();
            WaitForLoad();

            return new PriceListDetailsPage(_webDriver, _testContext);
        }

        public string GetFirstPricingName()
        {
            _firstPricingName = WaitForElementIsVisible(By.XPath(PRICING_NAME));
            return _firstPricingName.Text;
        }

        public void DeletePricing(string pricingName)
        {
            Actions action = new Actions(_webDriver);
            _pricing = WaitForElementIsVisible(By.XPath(string.Format(PRICING, pricingName)));
            action.MoveToElement(_pricing).Perform();

            _deleteBtn = WaitForElementIsVisible(By.XPath(string.Format(DELETE_PRICING_DEV, pricingName)));
            _deleteBtn.Click();
            WaitForLoad();

            _confirmDeleteBtn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDeleteBtn.Click();
            WaitForLoad();
        }

        public void DeleteAllPriceList()
        {
            if (!isPageSizeEqualsTo100())
            {
                PageSize("100");
            }

            var priceNumber = CheckTotalNumber();

            for (int i = 0; i < priceNumber; i++)
            {
                var priceName = GetFirstPricingName();
                DeletePricing(priceName);
            }
        }

        public void UnfoldFold()
        {
            _unfold = WaitForElementIsVisible(By.XPath(UNFOLD));
            _unfold.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public bool IsUnfold()
        {
            var elements = _webDriver.FindElements(By.XPath(CONTENT));

            foreach (var elm in elements)
            {
                if (!elm.GetAttribute("class").Equals("panel-collapse collapse show"))
                    return false;
            }

            return true;
        }

        public bool IsFold()
        {
            var elements = _webDriver.FindElements(By.XPath(CONTENT));

            foreach (var elm in elements)
            {
                if (!elm.GetAttribute("class").Equals("panel-collapse collapse"))
                    return false;
            }

            return true;
        }

        public int CountPeriod()
        {
            return _webDriver.FindElements(By.XPath(PERIOD)).Count;
        }

        public void DeletePeriod()
        {
            Actions action = new Actions(_webDriver);

            var firstPeriod = WaitForElementIsVisible(By.XPath(PERIOD));
            action.MoveToElement(firstPeriod).Perform();

            _deletePeriod = WaitForElementIsVisible(By.XPath(DELETE_PERIOD));
            _deletePeriod.Click();
            WaitForLoad();

            _confirmDeleteBtn = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDeleteBtn.Click();
            WaitForLoad();
        }

        public bool isPriceListLocal()
        {
            WaitForLoad();
            return isElementVisible(By.XPath(PRLIST_IS_LOCAL));

        }
    }
}
