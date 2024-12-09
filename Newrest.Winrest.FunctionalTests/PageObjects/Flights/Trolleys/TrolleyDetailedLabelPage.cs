using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys
{
    public class TrolleyDetailedLabelPage : PageBase
    {
        public TrolleyDetailedLabelPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
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
        private const string SHOW_ALL_SERVICE = "cbAirports";
        private const string SERVICE_VALID_POSITION = "cbAirports";
        private const string SERVICE_INVALID_POSITION = "cbAirports";
        private const string SERVICE_NO_POSITION_QTY = "cbAirports";

        private const string PROD_DATE_FILTER = "ProdDate";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMERS = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_CUSTOMER = "/html/body/div[11]/div/div/label/input";
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string DUPLICATE_LPCART = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[2]";

        private const string EXTENDED_MENU = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div";
        private const string PRINT_BTN = "create-print";
        private const string FIRST_ITEM = "//*[@id=\"detailedTrolleyTable\"]/tbody/tr[2]";
        private const string SECOND_ITEM = "//*[@id=\"detailedTrolleyTable\"]/tbody/tr[3]";
        private const string ITEM = "//*[@id=\"detailedTrolleyTable\"]/tbody/tr[{0}]";
        private const string SIMPLIFIED_MODE = "//*[@id=\"print-form\"]/div[1]/div[4]/div";
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

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        public IWebElement _extendedMenu;

        [FindsBy(How = How.Id, Using = PRINT_BTN)]
        public IWebElement _print;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        public IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = SECOND_ITEM)]
        public IWebElement _secondItem;
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
        public enum PrintType
        {
            Print,
            PrintMasterLabel
        }

        public PrintReportPage PrintReport(PrintType reportType, bool versionPrint)
        {
            Actions actions = new Actions(_webDriver);

            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));

            actions.MoveToElement(_extendedMenu).Perform();

            switch (reportType)
            {
                case PrintType.Print:
                    var printL = WaitForElementIsVisible(By.Id("btn-print-pdf"));
                    printL.Click();
                    WaitForLoad();
                    break;

                case PrintType.PrintMasterLabel:
                    var printM = WaitForElementIsVisible(By.Id("btn-print-master-pdf"));
                    printM.Click();
                    WaitForLoad();
                    break;

                default:
                    break;
            }

            _print = WaitForElementIsVisible(By.Id(PRINT_BTN));
            _print.Click();
            var close = WaitForElementIsVisible(By.Id("closeButton"));
            close.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);

        }

        public void AddPositionAndQuantity(string position, string qty)
        {
            int nbLignes = CheckTotalNumber();
            for (int n = 0; n < nbLignes; n++)
            {
                // passage en mode édition
                Thread.Sleep(2000);


                _firstItem = WaitForElementIsVisible(By.XPath(string.Format(ITEM, n + 2)));
                _firstItem.Click();
                // secouse du tableau
                Thread.Sleep(4000);

                var positionInput = WaitForElementIsVisible(By.XPath("//*/tr[" + (n + 2) + "]//input[@name='item.Position']"));

                positionInput.SetValue(ControlType.TextBox, position);
                WaitForLoad();

                var qtyInput = WaitForElementIsVisible(By.XPath("//*/tr[" + (n + 2) + "]//input[@name='item.MaxQuantity']"));

                qtyInput.SetValue(ControlType.TextBox, qty);
                WaitPageLoading();
                //waiting position saved

            }

        }


        public TrolleyAdjustingPage GoToTrolleyAjusting()
        {
            var ajustingTab = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div[2]/ul/li[3]/a"));
            ajustingTab.Click();

            WaitForLoad();
            return new TrolleyAdjustingPage(_webDriver, _testContext);
        }

        public bool verifyPosition(string label)
        {
            var positionLable = WaitForElementIsVisible(By.XPath("//*[@id=\"item_Position\"]"));
            
            if (positionLable.GetAttribute("value") == label)
            {
                return true;
            }
            else
            {
                return false;

            }
       
        }
        public bool verifyQty(string qty)
        {
            var qtyinput = WaitForElementIsVisible(By.XPath("//*[@id=\"item_MaxQuantity\"]"));

            if (qtyinput.GetAttribute("value") == qty)
            {
                return true;
            }
            else
            {
                return false;

            }

        }


        public PrintReportPage PrintSimplified()
        {
            bool versionPrint = true; 
            Actions actions = new Actions(_webDriver);

            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));

            actions.MoveToElement(_extendedMenu).Perform();

            var printL = WaitForElementIsVisible(By.Id("btn-print-pdf"));
            printL.Click();
            WaitForLoad();
            var _simplifiedMode = WaitForElementIsVisible(By.XPath(SIMPLIFIED_MODE));
            _simplifiedMode.Click();

            _print = WaitForElementIsVisible(By.Id(PRINT_BTN));
            _print.Click();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClosePrintButton();
            }
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[2]);

            return new PrintReportPage(_webDriver, _testContext);

        }

    }
}
