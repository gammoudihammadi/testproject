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
    public class TrolleyCountingPage : PageBase
    {
        public TrolleyCountingPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes _____________________________________

        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_SERVICE_FILTER = "SearchPattern";
        private const string SEARCH_FLIGHT_FILTER = "SearchPattern";
        private const string SORT_BY_FILTER = "cbSortBy";
        private const string SITES_FILTER = "SelectedSites_ms";
        private const string UNSELECT_SITES = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_SITE = "/html/body/div[10]/div/div/label/input";
        private const string GUEST_TYPE_FILTER = "/html/body/div[10]/div/div/label/input";
        private const string UNCHECK_GUEST_TYPE = "/html/body/div[10]/div/div/label/input";
        private const string GUEST_TYPE_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string SHOW_ALL_SERVICE = "cbAirports";
        private const string SERVICE_VALID_POSITION = "cbAirports";
        private const string SERVICE_INVALID_POSITION = "cbAirports";
        private const string SERVICE_NO_POSITION_QTY = "cbAirports";
        private const string PROD_DATE_FILTER = "date-from-picker";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string DUPLICATE_LPCART = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[2]";



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

        [FindsBy(How = How.Id, Using = SHOW_ALL_SERVICE)]
        private IWebElement _allServices;

        [FindsBy(How = How.Id, Using = SERVICE_VALID_POSITION)]
        private IWebElement _servicesValidPosition;

        [FindsBy(How = How.XPath, Using = SERVICE_INVALID_POSITION)]
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
            AllServices,
            ServicesValidPosition,
            ServiceInvalidPosition,
            ServiceNoPositionOrQty,
            Customers,
            GuestTypes,

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
                    break;

                case FilterType.Site:
                    _siteFilter = WaitForElementExists(By.Id(SITES_FILTER));
                    action.MoveToElement(_siteFilter).Perform();
                    _siteFilter.Click();
                    break;

                //case FilterType.SortBy:
                //    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY_FILTER));
                //    _sortBy.Click();
                //    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                //    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                //    _sortBy.Click();
                //    break;

                case FilterType.ProdDate:
                    _prodDate = WaitForElementIsVisible(By.Id(PROD_DATE_FILTER));
                    _prodDate.SetValue(ControlType.DateTime, value);
                    _prodDate.SendKeys(Keys.Tab);
                    break;

                case FilterType.AllServices:
                    _allServices = WaitForElementExists(By.XPath(SHOW_ALL_SERVICE));
                    action.MoveToElement(_allServices).Perform();
                    _allServices.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ServicesValidPosition:
                    _servicesValidPosition = WaitForElementExists(By.XPath(SERVICE_VALID_POSITION));
                    action.MoveToElement(_servicesValidPosition).Perform();
                    _servicesValidPosition.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ServiceInvalidPosition:
                    _serviceInvalidPosition = WaitForElementExists(By.XPath(SERVICE_INVALID_POSITION));
                    action.MoveToElement(_serviceInvalidPosition).Perform();
                    _serviceInvalidPosition.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ServiceNoPositionOrQty:
                    _serviceNoPositionOrQty = WaitForElementExists(By.XPath(SERVICE_NO_POSITION_QTY));
                    action.MoveToElement(_serviceNoPositionOrQty).Perform();
                    _serviceNoPositionOrQty.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Customers:
                    _customerFilter = WaitForElementExists(By.Id(CUSTOMER_FILTER));
                    action.MoveToElement(_customerFilter).Perform();
                    _customerFilter.Click();

                    // On décoche toutes les options
                    _unselectAllCustomers = _webDriver.FindElement(By.XPath(UNSELECT_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();

                    _customerFilter.Click();

                    break;

                case FilterType.GuestTypes:
                    _guestTypesFilter = WaitForElementExists(By.Id(GUEST_TYPE_FILTER));
                    action.MoveToElement(_guestTypesFilter).Perform();
                    _guestTypesFilter.Click();

                    _uncheckGuestTypes = WaitForElementIsVisible(By.XPath(UNCHECK_GUEST_TYPE));
                    _uncheckGuestTypes.Click();

                    _guestTypesSearch = WaitForElementIsVisible(By.XPath(GUEST_TYPE_SEARCH));
                    _guestTypesSearch.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(1500);

                    var guestTypeToCheck = WaitForElementIsVisible(By.XPath("/html/body/div[13]/ul"));
                    guestTypeToCheck.Click();

                    _guestTypesFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public bool ClickCount()
        {
            bool isAjustGoodQty = false;
            var countBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"countTrolleyTable\"]/tbody/tr[2]/td[10]/a"));
            countBtn.Click();

            WaitForLoad();

            var inputAjusted = WaitForElementIsVisible(By.Id("adjusted-input"));
            if (inputAjusted.GetAttribute("value") == "4")
            {
                isAjustGoodQty = true;
            }
            if(isElementVisible(By.Id("button-save")))
            {
                var save = WaitForElementIsVisible(By.Id("button-save"));
                save.Click();
            }
            else
            {
                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"trolley-i-form\"]/div[3]/button[2]"));
                close.Click();
            }

            WaitPageLoading();
            WaitForLoad();
            return isAjustGoodQty;
        }

        public bool IsCounted()
        {
            var ckeckCount = WaitForElementIsVisible(By.XPath("//*[@id=\"count-tick\"]"));
            if (ckeckCount.Selected)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsAjustQtyCorrect()
        {
            var ajustValue = WaitForElementIsVisible(By.Id("adjusted-input"));
            if (ajustValue.GetAttribute("value") == "4")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
