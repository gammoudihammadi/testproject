using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Trolleys
{
    public class TrolleyLightLabelPage : PageBase
    {
        public TrolleyLightLabelPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________ Constantes _____________________________________

        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string SEARCH_SERVICE_FILTER = "ServiceSearchPattern";
        private const string SEARCH_FLIGHT_FILTER = "FlightNoSearchPattern";
        private const string SORT_BY_FILTER = "cbSortBy";
        private const string SITES_FILTER = "SelectedSite";
        private const string UNSELECT_SITES = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_SITE = "/html/body/div[10]/div/div/label/input";
        private const string GUEST_TYPE_FILTER = "SelectedGuestTypes_ms";
        private const string UNCHECK_GUEST_TYPE = "/html/body/div[11]/div/div/label/input";
        private const string GUEST_TYPE_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string SHOW_ALL_SERVICE = "cbAirports";
        private const string SERVICE_VALID_POSITION = "cbAirports";
        private const string SERVICE_INVALID_POSITION = "cbAirports";
        private const string SERVICE_NO_POSITION_QTY = "cbAirports";
        private const string PROD_DATE_FILTER = "ProdDate";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_CUSTOMER = "/html/body/div[10]/div/div/label/input";
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/button";
        private const string DUPLICATE_LPCART = "//*[@id=\"tabContentItemContainer\"]/div/div/div[1]/div/a[2]";

        private const string VERIFY_SERVICE = "//*[@id=\"services-table\"]/tbody/tr[*]/td[8]";
        private const string VERIFY_FLIGHT = "//*[@id=\"services-table\"]/tbody/tr[*]/td[2]";
        private const string FIRST_ITEM = "//*[@id=\"services-table\"]/tbody/tr[2]";
        private const string EDIT_ITEM_BTN = "//*[@id=\"services-table\"]/tbody/tr[2]/td[10]/a[2]/span";

        private const string EXTENDED_MENU_LIGHT_LABEL = "//*[@id=\"trolley-list-form\"]/div[1]/div[1]/div";
        private const string EXTENDED_MENU = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div";
        private const string PRINT_BTN = "printButton";
        private const string CONFIG = "//*[@id=\"services-table\"]/tbody/tr[2]/td[10]/a[2]/span";

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

        [FindsBy(How = How.Id, Using = PRINT_BTN)]
        public IWebElement _print;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        public IWebElement _extendedMenu;

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU_LIGHT_LABEL)]
        public IWebElement _extendedMenuLightLabel;

        [FindsBy(How = How.XPath, Using = CONFIG)]
        public IWebElement _configTrolley;

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
                    _siteFilter = WaitForElementIsVisible(By.Id(SITES_FILTER));
                    _siteFilter.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORT_BY_FILTER));
                    _sortBy.Click();
                    var element1 = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element1.Text);
                    _sortBy.Click();
                    break;

                case FilterType.ProdDate:
                    _prodDate = WaitForElementIsVisible(By.Id(PROD_DATE_FILTER));
                    _prodDate.SetValue(ControlType.DateTime, value);
                    _prodDate.SendKeys(Keys.Tab);
                    break;



                case FilterType.Customers:
                    _customerFilter = WaitForElementExists(By.Id(CUSTOMER_FILTER));
                    action.MoveToElement(_customerFilter).Perform();
                    _customerFilter.Click();

                    // On décoche toutes les options
                    _unselectAllCustomers = _webDriver.FindElement(By.XPath(UNSELECT_CUSTOMERS));///html/body/div[10]/div/ul/li[2]/a/span[2]
                    _unselectAllCustomers.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();

                    _customerFilter.Click();

                    break;

                case FilterType.GuestTypes:
                    ComboBoxSelectById(new ComboBoxOptions(GUEST_TYPE_FILTER, (string)value));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }

            WaitPageLoading();
        }


        public enum PrintType
        {
            PrintLabel
        }

        public PrintReportPage PrintReport(PrintType reportType, bool versionPrint)
        {
            Actions actions = new Actions(_webDriver);

            _extendedMenuLightLabel = WaitForElementIsVisible(By.XPath(EXTENDED_MENU_LIGHT_LABEL));

            actions.MoveToElement(_extendedMenuLightLabel).Perform();

            switch (reportType)
            {
                case PrintType.PrintLabel:
                    var printL = WaitForElementIsVisible(By.XPath("//*[@id=\"trolley-list-form\"]/div[1]/div[1]/div/div/a"));//*[@id="trolley-list-form"]/div[1]/div[1]/div/div/a
                    printL.Click();
                    WaitForLoad();
                    break;

                default:
                    break;
            }

            // modal
            Thread.Sleep(2000);
            if (isElementVisible(By.XPath("//*/h4[text()='Print trolley labels']")))
            {
                Thread.Sleep(2000);
                //Print is successful but there is some errors
                //Quantity is too high for this config : cannot put 1000 in 999 labels of max quantity 1for service: Trolley Service, flight: 1258499327
                if (isElementVisible(By.Id("printButton")))
                {
                    var printModal = WaitForElementIsVisible(By.Id("printButton"));
                    printModal.Click();
                    WaitForLoad();
                }
                Thread.Sleep(2000);
                if (reportType == PrintType.PrintLabel)
                {
                    var closeModal = WaitForElementIsVisible(By.Id("closeButton"));
                    closeModal.Click();
                    WaitForLoad();
                }
            }

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

        public bool PrintReportErrorConfig(PrintType reportType)
        {
            Actions actions = new Actions(_webDriver);

            _extendedMenuLightLabel = WaitForElementIsVisible(By.XPath(EXTENDED_MENU_LIGHT_LABEL));

            actions.MoveToElement(_extendedMenuLightLabel).Perform();

            switch (reportType)
            {
                case PrintType.PrintLabel:
                    var printL = WaitForElementIsVisible(By.XPath("//*[@id=\"trolley-list-form\"]/div[1]/div[1]/div/div/a"));
                    printL.Click();
                    WaitForLoad();
                    break;

                default:
                    break;
            }


            _print = WaitForElementIsVisible(By.Id("printButton")); 
            _print.Click();
            WaitForLoad();

            if(isElementVisible(By.Id("TrolleyErrors")))
            {
                WaitForElementIsVisible(By.Id("TrolleyErrors"));
                return true;
            }
            else
            {
                return false;
            }
               
        }



        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.ProdDate, DateUtils.Now);
            }
        }

        public bool VerifyService(string value)
        {
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

            var elements = _webDriver.FindElements(By.XPath(VERIFY_SERVICE));


            foreach (var element in elements)
            {
                if (element.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool VerifyFlight(string value)
        {
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

        public bool VerifySite(string value)
        {
            Actions action = new Actions(_webDriver);

            var firstItem = WaitForElementExists(By.XPath(FIRST_ITEM));
            action.MoveToElement(firstItem).Perform();

            var editItem = WaitForElementExists(By.XPath(EDIT_ITEM_BTN));
            editItem.Click();

            WaitForLoad();
            IWebElement siteItem;
            if (isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div[2]/h2[2]/span")))
            {
                siteItem = WaitForElementExists(By.XPath("//*[@id=\"modal-1\"]/div/div/div[2]/h2[2]/span"));
            }
            else
            {
                siteItem = WaitForElementExists(By.XPath("//*[@id=\"modal-1\"]/div[2]/h2[2]/span"));
            }
            if (siteItem.Text.Contains(value)) return true;

            return false;

        }

        public bool IsDateRespected(DateTime prodDate, string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            var dates = _webDriver.FindElements(By.XPath("//*[@id=\"services-table\"]/tbody/tr[*]/td[3]"));

            foreach (var date in dates)
            {
                DateTime dateF = DateTime.Parse(date.Text, ci).Date;

                if (dateF != prodDate.Date && dateF != prodDate.AddDays(1).Date)
                {
                    // Schedule : production 0 ou -1
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByFlight()
        {
            var listeFlight = _webDriver.FindElements(By.XPath("//*[@id=\"services-table\"]/tbody/tr[*]/td[2]"));

            if (listeFlight.Count == 0)
                return false;


            var ancientCustomer = listeFlight[0].Text;

            foreach (var elm in listeFlight)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientCustomer) < 0)
                        return false;

                    ancientCustomer = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByService()
        {
            var listeServices = _webDriver.FindElements(By.XPath("//*[@id=\"services-table\"]/tbody/tr[*]/td[8]"));

            if (listeServices.Count == 0)
                return false;

 
            var ancientCustomer = listeServices[0].Text;

            foreach (var elm in listeServices)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientCustomer) < 0)
                        return false;

                    ancientCustomer = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsCustomer(string customer)
        {
            var listeCustomers = _webDriver.FindElements(By.XPath("//*[@id=\"services-table\"]/tbody/tr[*]/td[7]"));

            if (listeCustomers.Count == 0)
                return false;

            foreach (var elm in listeCustomers)
            {
                 if (!customer.Contains(elm.Text)) 
                     return false;

            }

            return true;
        }

        public string GetFirstFlight()
        {
            var value = WaitForElementIsVisible(By.XPath("//*[@id=\"services-table\"]/tbody/tr[2]/td[2]"));
            return value.Text;
        }
        public void ClickHideLabel()
        {
            Actions action = new Actions(_webDriver);

            var firstItem = WaitForElementExists(By.XPath(FIRST_ITEM));
            action.MoveToElement(firstItem).Perform();

            var eye = WaitForElementIsVisible(By.XPath("//*[@id=\"services-table\"]/tbody/tr[2]/td[10]/a[1]"));
            eye.Click();

            WaitForLoad();
        }

        public bool IsConfigIsSet() 
        {
            var listeConfig = _webDriver.FindElements(By.XPath("//*[@id=\"services-table\"]/tbody/tr[*]/td[9]"));//*[@id="services-table"]/tbody/tr[2]/td[9]

            if (listeConfig.Count == 0)
                return false;

            foreach (var elm in listeConfig)
            {

                if(elm.Text != "Qty 1 [ 1 - 1 ]") 
                    return false;

            }
            return true;

        }

        public bool IsConfigIsDeleted()
        {
            var listeConfig = _webDriver.FindElements(By.XPath("//*[@id=\"services-table\"]/tbody/tr[*]/td[9]/span"));

            if (listeConfig.Count == 0)
                return false;

            foreach (var elm in listeConfig)
            {

                if (elm.GetAttribute("title").ToString() != "No trolley config found for this service & site !")
                    return false;

            }
            return true;

        }

        //___________________________________ Pages ____________________________________________

        public ConfigurationModalPage ConfigurationModalPage()
        {
           
            Actions action = new Actions(_webDriver);

            var firstItem = WaitForElementExists(By.XPath(FIRST_ITEM));
            action.MoveToElement(firstItem).Perform();
            _configTrolley = WaitForElementIsVisible(By.XPath(CONFIG));
            _configTrolley.Click();

            WaitForLoad();
            return new ConfigurationModalPage(_webDriver, _testContext);
        }

        public TrolleyDetailedLabelPage GoToDetailledLabel()
        {
            var detailledTab = WaitForElementIsVisible(By.Id("hrefTabContentDetailedTrolleyContainer"));
            detailledTab.Click();

            WaitForLoad();
            return new TrolleyDetailedLabelPage(_webDriver, _testContext);
        }

        public TrolleyAdjustingPage GoToTrolleyAjusting()
        {
            var ajustingTab = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/form/div[2]/ul/li[3]/a"));
            ajustingTab.Click();
            WaitForLoad();

            return new TrolleyAdjustingPage(_webDriver, _testContext);
        }



    }
}
