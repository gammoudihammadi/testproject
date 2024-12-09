using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys
{
    public class TrolleyAdjustingPage : PageBase
    {
        public TrolleyAdjustingPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes _____________________________________

        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_SERVICE_FILTER = "SearchPattern";
        private const string SEARCH_FLIGHT_FILTER = "FlightNoSearchPattern";
        private const string SORT_BY_FILTER = "cbSortBy";
        private const string SITES_FILTER = "SelectedSite";
        private const string UNSELECT_SITES = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_SITE = "/html/body/div[10]/div/div/label/input";
        private const string GUEST_TYPE_FILTER = "/html/body/div[10]/div/div/label/input";
        private const string UNCHECK_GUEST_TYPE = "/html/body/div[10]/div/div/label/input";
        private const string GUEST_TYPE_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string FILTER_FLIGHT_KEY = "FlightKeySearch";
        private const string FILTER_TROLLEY_CODE = "cbAirports";
        private const string SERVICE_INVALID_POSITION = "cbAirports";
        private const string SERVICE_NO_POSITION_QTY = "cbAirports";
        private const string PROD_DATE_FILTER = "ProdDate";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string DUPLICATE_LPCART = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[2]";

        private const string VERIFY_FLIGHT = "//*[@id=\"services-table\"]/tbody/tr[*]/td[2]";


        // __________________________________ Filtres ________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_SERVICE_FILTER)]
        private IWebElement _searchService;

        [FindsBy(How = How.Id, Using = SEARCH_FLIGHT_FILTER)]
        private IWebElement _searchFlight;

        [FindsBy(How = How.Id, Using = SORT_BY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SITES_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.Id, Using = PROD_DATE_FILTER)]
        private IWebElement _prodDate;

        [FindsBy(How = How.Id, Using = FILTER_FLIGHT_KEY)]
        private IWebElement _searchFlightKey;

        [FindsBy(How = How.XPath, Using = FILTER_TROLLEY_CODE)]
        private IWebElement _serviceInvalidPosition;

        [FindsBy(How = How.XPath, Using = SERVICE_NO_POSITION_QTY)]
        private IWebElement _serviceNoPositionOrQty;

        [FindsBy(How = How.Id, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;

        [FindsBy(How = How.Id, Using = GUEST_TYPE_FILTER)]
        private IWebElement _guestTypesFilter;

        [FindsBy(How = How.XPath, Using = UNCHECK_GUEST_TYPE)]
        private IWebElement _uncheckGuestTypes;

        [FindsBy(How = How.XPath, Using = GUEST_TYPE_SEARCH)]
        private IWebElement _guestTypesSearch;
        public enum FilterType
        {
            SearchService,
            SearchFlight,
            Site,
            ProdDate,
            SortBy,
            ShowNotCountedOnly,
            SearchFlightKey,
            SearchTrolleyCode,
            SearchPosition,
            CountingSite,

        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.SearchService:
                    _searchService = WaitForElementIsVisible(By.Id(SEARCH_SERVICE_FILTER));
                    _searchService.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SearchFlight:
                    _searchFlight = WaitForElementIsVisible(By.Id(SEARCH_FLIGHT_FILTER));
                    _searchFlight.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;

                case FilterType.Site:
                    _siteFilter = WaitForElementExists(By.Id(SITES_FILTER));
                    _siteFilter.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY_FILTER));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;

                case FilterType.ProdDate:
                    _prodDate = WaitForElementIsVisible(By.Id(PROD_DATE_FILTER));
                    _prodDate.SetValue(ControlType.DateTime, value);
                    _prodDate.SendKeys(Keys.Tab);
                    break;

                case FilterType.SearchFlightKey:
                    _searchFlightKey = WaitForElementIsVisible(By.Id(FILTER_FLIGHT_KEY));
                    _searchFlightKey.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;

                case FilterType.SearchTrolleyCode:
                    _searchService = WaitForElementIsVisible(By.Id(FILTER_TROLLEY_CODE));
                    _searchService.SetValue(ControlType.TextBox, value);
                    break;

                //case FilterType.SearchPosition:
                //    _searchService = WaitForElementIsVisible(By.Id(FILTER_POSITION));
                //    _searchService.SetValue(ControlType.TextBox, value);
                //    break;

                //case FilterType.CountingSite:
                //    _siteFilter = WaitForElementExists(By.Id(FILTER_COUNTING_SITE));
                //    action.MoveToElement(_siteFilter).Perform();
                //    _siteFilter.Click();
                //    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            Thread.Sleep(2000);
        }


        public bool VerifyFlight(string value)
        {
            WaitForLoad();
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(VERIFY_FLIGHT));


            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public void ClickAjust()
        {
            var ajustBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"countTrolleyTable\"]/tbody/tr[2]/td[9]/a"));            
            ajustBtn.Click();

            WaitForLoad();

            var inputAjusted = WaitForElementIsVisible(By.Id("adjusted-input"));
            inputAjusted.SetValue(ControlType.TextBox, "4");

            Thread.Sleep(1000);
            WaitForLoad();

            var save = WaitForElementIsVisible(By.Id("button-save"));
            save.Click();

            Thread.Sleep(1500);
            WaitForLoad();
        }

        public bool IsAjusted()
        {
            var ckeckAjust = WaitForElementIsVisible(By.XPath("//*[@id=\"countTrolleyTable\"]/tbody/tr[2]/td[8]/input"));
            if (ckeckAjust.Selected)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public TrolleyCountingPage GoToTrolleyCounting()
        {
            var ajustingTab = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[2]/ul/li[4]/a"));
            ajustingTab.Click();

            WaitForLoad();
            return new TrolleyCountingPage(_webDriver, _testContext);
        }

        public TrolleyDetailedLabelPage GoToTrolleyDetailedLabel()
        {
            var ajustingTab = WaitForElementIsVisible(By.Id("hrefTabContentDetailedTrolleyContainer"));
            ajustingTab.Click();

            WaitForLoad();
            return new TrolleyDetailedLabelPage(_webDriver, _testContext);
        }

        public bool IsAllQuantity()
        {

            int qty = 0;

            int tot = CheckTotalNumber();

            for (int i = 1; i < tot + 1; i++)
            {
                Thread.Sleep(1000);
                try
                {
                    var ajustBtn = _webDriver.FindElement(By.XPath(String.Format("//*[@id=\"countTrolleyTable\"]/tbody/tr[{0}]/td[9]/a", i + 1)));
                    ajustBtn.Click();
                    WaitForLoad();
                }
                catch (StaleElementReferenceException ex) //if StaleElementReferenceException, retry operation
                {
                    var ajustBtn = _webDriver.FindElement(By.XPath(String.Format("//*[@id=\"countTrolleyTable\"]/tbody/tr[{0}]/td[9]/a", i + 1)));
                    ajustBtn.Click();
                    WaitForLoad();
                }

                var value = WaitForElementIsVisible(By.Id("adjusted-input")).GetAttribute("value");
                qty = qty + int.Parse(value);

                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"trolley-i-form\"]/div[3]/button[2]"));
                close.Click();
                WaitForLoad();

            }
            if (qty == 10)
            {
                return true;
            }


            return false;
        }
    }
}
