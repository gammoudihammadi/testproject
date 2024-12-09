using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServicePage : PageBase
    {

        public ServicePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________________________________ Constantes ____________________________________________________________

        // General

        private const string PRINT_LIST_MODAL = "/html/body/div[13]/h3";
        private const string NEW_SERVICE = "New service";

        private const string FOLD_BTN = "unfoldBtn";
        private const string UNFOLD_BTN = "unfoldBtn";
        private const string IMPORT = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/div/a[2]";

        private const string EXPORT = "btn-export-services";
        private const string PRINT = "Print cycles calendar";
        private const string PRINT_DATE_FROM = "datepicker-print-from";
        private const string PRINT_DATE_TO = "datepicker-print-to";
        private const string CONFIRM_PRINT = "btn-submit-print-calendar";
        private const string FIRST_ROW = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]";
        private const string FIRST_ROW_IN_CUTERMER_SERVICE = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr[1]";

        // Tableau services
        private const string DETAIL = "//*[starts-with(@id,\"content_\")]";
        private const string FIRST_SERVICE_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string SECOND_SERVICE_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string COUNTER = "//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span";

        // Filtres
        private const string RESET_FILTER = "ResetFilter";

        private const string SEARCH_FILTER = "SearchPattern";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string SHOW_ALL = "//*[@id=\"ServiceFilter\"][@value='All']";
        private const string SHOW_ACTIVE = "//*[@id=\"ServiceFilter\"][@value='Active']";
        private const string SHOW_INACTIVE = "//*[@id=\"ServiceFilter\"][@value='Inactive']";
        private const string SHOW_GENERIC_SERVICES = "//*[@id=\"ServiceFilter\"][@value='Generic']";
        private const string SHOW_PRICE_EXPIRED = "//*[@id=\"ServiceFilter\"][@value='WithExpiredPrices']";
        private const string SHOW_PRICE_VALID = "//*[@id=\"WithValidPriceOnRadio\"][@value='WithValidPriceOn']";
        private const string VALID_PRICES_DATE = "specific-date";
        private const string CUSTOMERS_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SELECT_ALL_CUSTOMERS = "/html/body/div[10]/div/ul/li[1]/a";
        private const string CUSTOMERS_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string SITES_FILTER = "SelectedSites_ms";
        private const string UNSELECT_ALL_SITES = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SELECT_ALL_SITES = "/html/body/div[11]/div/ul/li[1]/a";
        private const string SITES_SEARCH = "/html/body/div[11]/div/div/label/input";
        private const string CATEGORY_FILTER = "SelectedCategories_ms";
        private const string UNSELECT_ALL_CATEGORY = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SELECT_ALL_CATEGORY = "/html/body/div[12]/div/ul/li[1]/a";
        private const string CATEGORY_SEARCH = "/html/body/div[12]/div/div/label/input";

        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/div[{0}]/div[1]/div/div[2]/table/tbody/tr/td[1]/img[@alt='Inactive']";
        private const string SERVICE_NAME = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]";
        private const string CATEGORY = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[3]";
        private const string ACTIVE_PRICE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[6]";

        private const string CHOOSE_FILE = "fileSent";
        private const string VALIDATE_IMPORT = "//*[@id=\"form-import\"]/div[3]/button[2]";

        private const string MASSIVE_DELETE = "/html/body/div[2]/div/div[2]/div[1]/div/div[1]/div/a[4]";
        private const string SHOW_SELECTED = "//*[@id=\"ServiceFilter\"][@checked]";
        private const string CUSTOMERS = "SelectedCustomers_ms";
        private const string CUSTOMERS_SELECTED = "/html/body/div[10]/ul/li/label/input";
        private const string CATEGORIES = "SelectedCategories_ms";
        private const string CATEGORIES_SELECTED = "/html/body/div[12]/ul/li/label/input";
        private const string SITES = "SelectedSites_ms";
        private const string SITES_SELECTED = "/html/body/div[11]/ul/li/label/input";
        private const string ITEMS = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]";
        public const string DeliveriesID = "hrefTabContentDeliveries";
        public const string DELIVERYFORSERVICECHECK = "//*[@id=\"hrefTabContentLoading\"]";
        public const string INDEXSERVICE = "//*[@id=\"list-item-with-action\"]/div[{0}]";
        public const string PRICESERVICE = "//*[@id=\"hrefTabContentPrice\"]";
        public const string GeneralInfoSERVICE = "//*[@id=\"hrefTabContentService\"]"; 

        public const string SERVICEHEADER = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[2]";
        public const string CATEGORYHEADER = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[3]";
        public const string ACTIVEPRICEHEADER = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[6]";
        public const string MENUHEADER = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[7]";
        public const string DELIVERYHEADER = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[8]";
        public const string LOADINGPLANHEADER = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div[2]/table/thead/tr/th[9]";

        //_______________________________________________________ Variables _____________________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = MENUHEADER)]
        private IWebElement _menu_header;

        [FindsBy(How = How.XPath, Using = DELIVERYHEADER)]
        private IWebElement _delivery_header;

        [FindsBy(How = How.XPath, Using = LOADINGPLANHEADER)]
        private IWebElement _loadingplan_header;

        [FindsBy(How = How.XPath, Using = ACTIVEPRICEHEADER)]
        private IWebElement _activeprice_header;

        [FindsBy(How = How.XPath, Using = CATEGORYHEADER)]
        private IWebElement _category_header;

        [FindsBy(How = How.XPath, Using = SERVICEHEADER)]
        private IWebElement _service_header;

        [FindsBy(How = How.XPath, Using = NEW_SERVICE)]
        private IWebElement _createNewService;

        [FindsBy(How = How.Id, Using = DeliveriesID)]
        private IWebElement _deliveries;
        [FindsBy(How = How.XPath, Using = CUSTOMERS_SELECTED)]
        private IWebElement _customerselected;
        [FindsBy(How = How.Id, Using = CUSTOMERS)]
        private IWebElement _customers;
        [FindsBy(How = How.Id, Using = DELIVERYFORSERVICECHECK)]
        private IWebElement _deliveryservice;


        [FindsBy(How = How.XPath, Using = CATEGORIES_SELECTED)]
        private IWebElement _categoriesselected;
        [FindsBy(How = How.Id, Using = CATEGORIES)]
        private IWebElement _categories;

        [FindsBy(How = How.XPath, Using = SITES_SELECTED)]
        private IWebElement _siteselected;
        [FindsBy(How = How.Id, Using = SITES)]
        private IWebElement _sites;

        [FindsBy(How = How.Id, Using = UNFOLD_BTN)]
        private IWebElement _unfoldBtn;

        [FindsBy(How = How.Id, Using = FOLD_BTN)]
        private IWebElement _foldBtn;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.LinkText, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = PRINT_DATE_FROM)]
        private IWebElement _printDateFrom;

        [FindsBy(How = How.Id, Using = PRINT_DATE_TO)]
        private IWebElement _printDateTo;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;


        // Tableau services
        [FindsBy(How = How.XPath, Using = FIRST_SERVICE_NAME)]
        private IWebElement _firstServiceName;

        [FindsBy(How = How.XPath, Using = DETAIL)]
        private IWebElement _detail;


        //_______________________________________________________ Filters _____________________________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByFilter;

        [FindsBy(How = How.XPath, Using = SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE)]
        private IWebElement _showActive;

        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE)]
        private IWebElement _showInactive;

        [FindsBy(How = How.XPath, Using = SHOW_GENERIC_SERVICES)]
        private IWebElement _showGenericServices;

        [FindsBy(How = How.XPath, Using = SHOW_PRICE_EXPIRED)]
        private IWebElement _showServicesWithExpiredPrices;

        [FindsBy(How = How.XPath, Using = SHOW_PRICE_VALID)]
        private IWebElement _showServicesWithValidPrice;

        [FindsBy(How = How.Id, Using = VALID_PRICES_DATE)]
        private IWebElement _validPriceDate;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CUSTOMERS)]
        private IWebElement _unselectAllCustomers;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_SEARCH)]
        private IWebElement _searchCustomers;

        [FindsBy(How = How.Id, Using = SITES_FILTER)]
        private IWebElement _sitesFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_SITES)]
        private IWebElement _unselectAllSites;

        [FindsBy(How = How.XPath, Using = SITES_SEARCH)]
        private IWebElement _searchSites;

        [FindsBy(How = How.Id, Using = CATEGORY_FILTER)]
        private IWebElement _categoryFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_CATEGORY)]
        private IWebElement _unselectAllCategory;

        [FindsBy(How = How.XPath, Using = CATEGORY_SEARCH)]
        private IWebElement _searchCategory;

        //import

        [FindsBy(How = How.Id, Using = CHOOSE_FILE)]
        private IWebElement _chooseFile;

        [FindsBy(How = How.XPath, Using = VALIDATE_IMPORT)]
        private IWebElement _validateImport;

        [FindsBy(How = How.XPath, Using = IMPORT)]
        private IWebElement _import;

        [FindsBy(How = How.XPath, Using = FIRST_ROW_IN_CUTERMER_SERVICE)]
        private IWebElement _firstRow;

        public enum FilterType
        {
            Search,
            SortBy,
            ShowAll,
            ShowOnlyActive,
            ShowOnlyInactive,
            ShowGenericServices,
            ShowServicesWithExpiredPrice,
            ShowServicesWithValidPrice,
            Customers,
            Sites,
            Categories,
            CustomersUncheck,
            SitesUncheck,
            CategoriesUncheck,
            CategoriesCheckAll,
            SitesCheckAll,
            CustomersCheckAll
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisibleNew(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    WaitLoading();
                    break;
                case FilterType.SortBy:
                    _sortByFilter = WaitForElementIsVisibleNew(By.Id(SORTBY_FILTER));
                    _sortByFilter.Click();
                    var element = WaitForElementIsVisibleNew(By.XPath("//*[@id=\"cbSortBy\"]/option[@value='" + value + "']"));
                    _sortByFilter.SetValue(ControlType.DropDownList, element.Text);
                    _sortByFilter.Click();
                    break;
                case FilterType.ShowAll:
                    _showAll = WaitForElementExists(By.XPath(SHOW_ALL));
                    _showAll.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowOnlyActive:
                    _showActive = WaitForElementExists(By.XPath(SHOW_ACTIVE));
                    _showActive.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowOnlyInactive:
                    _showInactive = WaitForElementExists(By.XPath(SHOW_INACTIVE));
                    _showInactive.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowGenericServices:
                    _showGenericServices = WaitForElementExists(By.XPath(SHOW_GENERIC_SERVICES));
                    _showGenericServices.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowServicesWithExpiredPrice:
                    _showServicesWithExpiredPrices = WaitForElementExists(By.XPath(SHOW_PRICE_EXPIRED));
                    _showServicesWithExpiredPrices.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowServicesWithValidPrice:
                    _showServicesWithValidPrice = WaitForElementExists(By.XPath(SHOW_PRICE_VALID));
                    _showServicesWithValidPrice.SetValue(ControlType.RadioButton, value);

                    _validPriceDate = WaitForElementIsVisibleNew(By.Id(VALID_PRICES_DATE));
                    _validPriceDate.SetValue(ControlType.DateTime, DateUtils.Now);
                    _validPriceDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.CustomersUncheck:
                    _customersFilter = WaitForElementIsVisibleNew(By.Id(CUSTOMERS_FILTER));
                    _customersFilter.Click();

                    _unselectAllCustomers = WaitForElementIsVisibleNew(By.XPath(UNSELECT_ALL_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _customersFilter.Click();
                    break;

                case FilterType.CustomersCheckAll:
                    _customersFilter = WaitForElementIsVisibleNew(By.Id(CUSTOMERS_FILTER));
                    _customersFilter.Click();

                    var selectAllCustomers = WaitForElementIsVisibleNew(By.XPath(SELECT_ALL_CUSTOMERS));
                    selectAllCustomers.Click();

                    _customersFilter.Click();
                    break;

                case FilterType.Customers:
                    _customersFilter = WaitForElementIsVisibleNew(By.Id(CUSTOMERS_FILTER));
                    _customersFilter.Click();

                    _unselectAllCustomers = WaitForElementIsVisibleNew(By.XPath(UNSELECT_ALL_CUSTOMERS));
                    _unselectAllCustomers.Click();

                    _searchCustomers = WaitForElementIsVisibleNew(By.XPath(CUSTOMERS_SEARCH));
                    _searchCustomers.SetValue(ControlType.TextBox, value);
                    WaitForLoad();

                    var customerToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + "']"));
                    customerToCheck.SetValue(ControlType.CheckBox, true);

                    _customersFilter.Click();
                    break;
                case FilterType.Sites:
                    _sitesFilter = WaitForElementIsVisibleNew(By.Id(SITES_FILTER));
                    _sitesFilter.Click();

                    _unselectAllSites = WaitForElementIsVisibleNew(By.XPath(UNSELECT_ALL_SITES));
                    _unselectAllSites.Click();

                    _searchSites = WaitForElementIsVisibleNew(By.XPath(SITES_SEARCH));
                    _searchSites.SetValue(ControlType.TextBox, value);

                    var siteToCheck = WaitForElementIsVisibleNew(By.XPath("//span[text()='" + value + " - " + value + "']"));
                    siteToCheck.SetValue(ControlType.CheckBox, true);

                    _sitesFilter.Click();
                    break;
                case FilterType.SitesUncheck:
                    _sitesFilter = WaitForElementIsVisibleNew(By.Id(SITES_FILTER));
                    _sitesFilter.Click();

                    _unselectAllSites = WaitForElementIsVisibleNew(By.XPath(UNSELECT_ALL_SITES));
                    _unselectAllSites.Click();

                    _sitesFilter.Click();
                    break;

                case FilterType.SitesCheckAll:
                    _sitesFilter = WaitForElementIsVisibleNew(By.Id(SITES_FILTER));
                    _sitesFilter.Click();

                    var selectAllSites = WaitForElementIsVisibleNew(By.XPath(SELECT_ALL_SITES));
                    selectAllSites.Click();

                    _sitesFilter.Click();
                    break;

                case FilterType.Categories:
                    ScrollUntilElementIsInView(By.Id(CATEGORY_FILTER));
                    ComboBoxSelectById(new ComboBoxOptions(CATEGORY_FILTER, (string)value));
                    break;
                case FilterType.CategoriesUncheck:
                    _categoryFilter = WaitForElementIsVisibleNew(By.Id(CATEGORY_FILTER));
                    _categoryFilter.Click();

                    _unselectAllCategory = WaitForElementIsVisibleNew(By.XPath(UNSELECT_ALL_CATEGORY));
                    _unselectAllCategory.Click();

                    _categoryFilter.Click();
                    break;

                case FilterType.CategoriesCheckAll:
                    _categoryFilter = WaitForElementIsVisibleNew(By.Id(CATEGORY_FILTER));
                    _categoryFilter.Click();

                    var selectAllCategory = WaitForElementIsVisibleNew(By.XPath(SELECT_ALL_CATEGORY));
                    selectAllCategory.Click();

                    _categoryFilter.Click();
                    break;
                default:
                    break;
            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilters()
        {
            _resetFilterDev = WaitForElementIsVisibleNew(By.Id(RESET_FILTER));
            _resetFilterDev.Click();

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de date
            }
        }

        public void UncheckAllCustomers()
        {
            _customersFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
            _customersFilter.Click();

            _unselectAllCustomers = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_CUSTOMERS));
            _unselectAllCustomers.Click();

            _customersFilter.Click();

            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public void UncheckAllSites()
        {
            _sitesFilter = WaitForElementIsVisible(By.Id(SITES_FILTER));
            _sitesFilter.Click();

            _unselectAllSites = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_SITES));
            _unselectAllSites.Click();

            _sitesFilter.Click();

            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public bool IsSortedByName()
        {
            var listeNames = _webDriver.FindElements(By.XPath(SERVICE_NAME));

            if (listeNames.Count == 0)
                return false;

            var ancienName = listeNames[0].Text;

            foreach (var elm in listeNames)
            {
                try
                {
                    if (elm.Text.CompareTo(ancienName) < 0)
                        return false;

                    ancienName = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByCategory()
        {
            var listeCategories = _webDriver.FindElements(By.XPath(CATEGORY));

            if (listeCategories.Count == 0)
                return false;

            var ancienCategory = listeCategories[0].Text;

            foreach (var elm in listeCategories)
            {
                try
                {
                    if (elm.Text.CompareTo(ancienCategory) < 0)
                        return false;

                    ancienCategory = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                if (isElementVisible(By.XPath(String.Format(INACTIVE, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(String.Format(INACTIVE, i + 1)));

                    if (active)
                        return false;
                }
                else
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public bool IsPriceActive()
        {
            var listePrices = _webDriver.FindElements(By.XPath(ACTIVE_PRICE));

            if (listePrices.Count == 0)
                return false;

            foreach (var elm in listePrices)
            {
                if (elm.Text.Equals("0"))
                    return false;
            }

            return true;
        }

        public bool VerifyCategory(string category)
        {
            var listeCategories = _webDriver.FindElements(By.XPath(CATEGORY));

            if (listeCategories.Count == 0)
                return false;

            foreach (var elm in listeCategories)
            {
                if (!elm.GetAttribute("innerText").Contains(category))
                    return false;
            }

            return true;
        }

        public bool VerifyCustomer(string customer)
        {
            var listeCustomers = _webDriver.FindElements(By.XPath(SERVICE_NAME));

            if (listeCustomers.Count == 0)
                return false;

            foreach (var elm in listeCustomers)
            {
                if (!elm.Text.Contains(customer))
                    return false;
            }

            return true;
        }

        //import
        public bool ImportCustomersExcelFile(string filePath)
        {
            ShowExtendedMenu();

            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();
            WaitForLoad();

            _chooseFile = WaitForElementIsVisible(By.Id(CHOOSE_FILE));
            _chooseFile.SendKeys(filePath);

            _validateImport = WaitForElementIsVisible(By.XPath(VALIDATE_IMPORT));
            _validateImport.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath("//*[@id=\"form-import\"]/div[2]/div/div/p[1]/b")))
            {
                WaitForElementIsVisible(By.XPath("//*[@id=\"form-import\"]/div[2]/div/div/p[1]/b"));
                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"form-import\"]/div[3]/button"));
                close.Click();
                WaitForLoad();
                return false;
            }
            else
            {
                var close = WaitForElementIsVisible(By.XPath("//*[@id=\"form-import\"]/div[3]/button"));
                close.Click();
                WaitForLoad();
                return true;
            }
        }

        //_______________________________________________________ Méthodes ____________________________________________________________

        // General
        public ServiceCreateModalPage ServiceCreatePage()
        {
            WaitPageLoading();
            WaitForLoad();
            ShowPlusMenu();
            WaitForLoad();
            _createNewService = WaitForElementIsVisibleNew(By.LinkText(NEW_SERVICE));
            _createNewService.Click();
            WaitPageLoading();
            WaitForLoad();

            return new ServiceCreateModalPage(_webDriver, _testContext);
        }
        public ServiceCreateModalPage ServiceCreate()
        {
        
            ShowPlusMenu();
            WaitLoading();

            _createNewService = WaitForElementIsVisible(By.LinkText(NEW_SERVICE));

            // Scroller jusqu'à l'élément
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].scrollIntoView(true);", _createNewService);
            _createNewService.Click();

           

            return new ServiceCreateModalPage(_webDriver, _testContext);
        }

        public void FoldAll()
        {
            ShowExtendedMenu();

            _foldBtn = WaitForElementIsVisible(By.Id(FOLD_BTN));
            _foldBtn.Click();
            WaitForLoad();
        }

        public Boolean IsFoldAll()
        {
            _detail = WaitForElementExists(By.XPath(DETAIL));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();

            return _detail.GetAttribute("class") == "panel-collapse collapse";
        }

        public void UnfoldAll()
        {
            ShowExtendedMenu();

            _unfoldBtn = WaitForElementIsVisible(By.Id(UNFOLD_BTN));
            _unfoldBtn.Click();
            WaitForLoad();
        }


        public ServiceImportPage Import()
        {
            WaitForLoad();

            ShowExtendedMenu();

            _import = WaitForElementIsVisible(By.XPath(IMPORT), "IMPORT");
            _import.Click();
            WaitForLoad();


            return new ServiceImportPage(_webDriver, _testContext);
        }

        public ServiceImportPage ImportFile()
        {
            ShowExtendedMenu();

            _import = WaitForElementIsVisible(By.XPath(IMPORT));
            _import.Click();
            WaitForLoad();

            return new ServiceImportPage(_webDriver, _testContext);
        }


        public Boolean IsUnfoldAll()
        {
            _detail = WaitForElementIsVisible(By.XPath(DETAIL));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();
            return _detail.GetAttribute("class") == "panel-collapse collapse show";
        }

        public void Export(bool versionPrint)
        {
            ShowExtendedMenu();
            WaitForLoad();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClosePrintButton();
                WaitForLoad();
            }

            WaitForDownload();
            Close();
        }

        public void PrintCyclesForCalendar(bool versionPrint, DateTime dateFrom, DateTime dateTo)
        {
            ShowExtendedMenu();
            _print = WaitForElementIsVisible(By.LinkText(PRINT));
            _print.Click();
            WaitForLoad();

            _printDateFrom = WaitForElementIsVisible(By.Id(PRINT_DATE_FROM));
            _printDateFrom.SetValue(ControlType.DateTime, dateFrom);
            _printDateFrom.SendKeys(Keys.Tab);

            _printDateTo = WaitForElementIsVisible(By.Id(PRINT_DATE_TO));
            _printDateTo.SetValue(ControlType.DateTime, dateTo);
            _printDateTo.SendKeys(Keys.Tab);

            _confirmPrint = WaitForElementIsVisible(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExcelFile(FileInfo[] taskFiles, bool export)
        {
            WaitForLoad();

            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (IsExcelFileCorrect(file.Name, export))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
            {
                return null;
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsExcelFileCorrect(string filePath, bool export)
        {
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes

            Regex reg;

            if (export)
            {
                reg = new Regex("^Prices" + "_" + annee + mois + jour + "_" + heure + minutes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            }
            else
            {
                reg = new Regex("^CyclesCalendar" + "_" + annee + mois + jour + "_" + heure + minutes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            }

            Match m = reg.Match(filePath);

            return m.Success;
        }


        // Tableau services
        public ServicePricePage ClickOnFirstService()
        {
            WaitPageLoading();
            _firstServiceName = WaitForElementIsVisible(By.XPath(FIRST_SERVICE_NAME));
            _firstServiceName.Click();
            WaitForLoad();
            return new ServicePricePage(_webDriver, _testContext);
        }

        public string GetFirstServiceName()
        {
            _firstServiceName = WaitForElementIsVisibleNew(By.XPath(FIRST_SERVICE_NAME));
            WaitForLoad();
            return _firstServiceName.Text;
        }

        public string GetSecondServiceName()
        {
            _firstServiceName = WaitForElementIsVisible(By.XPath(SECOND_SERVICE_NAME));
            return _firstServiceName.Text;
        }

        internal int TableCount()
        {
            var counter = WaitForElementIsVisible(By.XPath(COUNTER), "COUNTER");
            return Convert.ToInt32(counter.Text);
        }

        public void CheckExport(FileInfo trouveXLSX)
        {
            List<string> servicesXLSX = OpenXmlExcel.GetValuesInList("Service", "Services Prices", trouveXLSX.FullName);
            int resultNumber = servicesXLSX.Distinct().Count();
            WaitPageLoading();
            Assert.AreEqual(CheckTotalNumber(), resultNumber);
            var services = _webDriver.FindElements(By.XPath("//*/td[2]/a[starts-with(@href,'/Customers/Service/Detail?id=')]"));
            Assert.AreEqual(Math.Min(resultNumber, 100), services.Count);
            Dictionary<string, int> serviceLigneXLSX = new Dictionary<string, int>();
            Dictionary<string, string> serviceLien = new Dictionary<string, string>();
            List<string> serviceNames = new List<string>();
            int i = 0;
            foreach (var service in services)
            {
                //if (service.Text == "") continue;
                string serviceName = service.Text.Substring(0, service.Text.IndexOf(" - "));
                serviceNames.Add(serviceName);
                Assert.IsTrue(servicesXLSX.Contains(serviceName), "Service " + service.Text + " non présent dans le XLSX");
                //index
                // services<string,noLigne>
                if (!serviceLigneXLSX.ContainsKey(serviceName))
                {
                    // les premiers prix sont affichés en premier
                    //serviceLigneXLSX.Add(serviceName, i);
                    serviceLigneXLSX.Add(serviceName, servicesXLSX.IndexOf(serviceName));
                }
                string lien = service.GetAttribute("href");
                serviceLien.Add(serviceName, lien);
                i++;
            }

            //Category    Is Produced Is Invoiced Is Generic Is SPML Check List Site Code
            //BEBIDAS     VRAI        VRAI        FAUX       FAUX    VRAI       MAD

            List<string> categoryXLSX = OpenXmlExcel.GetValuesInList("Category", "Services Prices", trouveXLSX.FullName);
            List<string> isProducedXLSX = OpenXmlExcel.GetValuesInList("Is Produced", "Services Prices", trouveXLSX.FullName);
            List<string> isInvoicedXLSX = OpenXmlExcel.GetValuesInList("Is Invoiced", "Services Prices", trouveXLSX.FullName);
            List<string> isGenericXLSX = OpenXmlExcel.GetValuesInList("Is Generic", "Services Prices", trouveXLSX.FullName);
            List<string> isSPMLXLSX = OpenXmlExcel.GetValuesInList("Is SPML", "Services Prices", trouveXLSX.FullName);
            List<string> isCheckListXLSX = OpenXmlExcel.GetValuesInList("Check List", "Services Prices", trouveXLSX.FullName);
            //List<string> siteCodeXLSX = OpenXmlExcel.GetValuesInList("Site Code", "Services Prices", trouveXLSX.FullName);
            List<string> priceXLSX = OpenXmlExcel.GetValuesInList("Price", "Services Prices", trouveXLSX.FullName);

            // on scan i : les lignes du tableau
            for (i = 0; i < Math.Min(resultNumber, 10); i++)
            {
                //clée
                string serviceName = serviceNames[i];
                // c'est un lien
                _webDriver.Navigate().GoToUrl(serviceLien[serviceName]);
                WaitPageLoading();
                var servPricePage = new ServicePricePage(_webDriver, _testContext);
                servPricePage.UnfoldAll();
                //servPricePage.EditFirstPrice();
                // pour chaque item
                Actions action = new Actions(_webDriver);
                var _content = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,\"content_\")]"));
                action.MoveToElement(_content).Perform();
                WaitForLoad();
                var edit = WaitForElementIsVisible(By.Id("btn-edit-price"));
                edit.Click();
                Thread.Sleep(2000);
                WaitForLoad();
                //TODO
                var editPrice = new ServiceCreatePriceModalPage(_webDriver, _testContext);
                Assert.AreEqual(String.IsNullOrEmpty(editPrice.GetPrice()) ? "0" : editPrice.GetPrice(), priceXLSX[serviceLigneXLSX[serviceName]]);
                editPrice.Close();
                var generalInfo = servPricePage.ClickOnGeneralInformationTab();
                Assert.AreEqual(generalInfo.GetCategory(), categoryXLSX[serviceLigneXLSX[serviceName]], i + " " + serviceName + " ligne " + serviceLigneXLSX[serviceName] + " Catégorie différente entre Excel et Winrest");
                Assert.AreEqual(generalInfo.IsProduced() ? "TRUE" : "FALSE", isProducedXLSX[serviceLigneXLSX[serviceName]], "IsProduced différent entre Excel et Winrest");
                Assert.AreEqual(generalInfo.IsInvoiced() ? "TRUE" : "FALSE", isInvoicedXLSX[serviceLigneXLSX[serviceName]], "IsProduced différent entre Excel et Winrest");
                Assert.AreEqual(generalInfo.IsGeneric() ? "TRUE" : "FALSE", isGenericXLSX[serviceLigneXLSX[serviceName]], "IsProduced différent entre Excel et Winrest");
                Assert.AreEqual(generalInfo.IsSPML() ? "TRUE" : "FALSE", isSPMLXLSX[serviceLigneXLSX[serviceName]], "IsProduced différent entre Excel et Winrest");
                Assert.AreEqual(generalInfo.IsCheckList() ? "TRUE" : "FALSE", isCheckListXLSX[serviceLigneXLSX[serviceName]], "IsProduced différent entre Excel et Winrest");
                //ServicePage retour = generalInfo.BackToList();
            }
        }

        //public void PageUp()
        //{
        //IJavaScriptExecutor js = (IJavaScriptExecutor)_webDriver;
        //js.ExecuteScript("window.scrollBy(0, -350)", "");
        //}

        public void CreateNewService(string code, string servicName = null)
        {
            if (servicName != null)
            {
                var serviceName = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/div/form/div/div[1]/div[1]/div[1]/div[1]/div/input"));
                serviceName.SetValue(ControlType.TextBox, servicName);
            }
            else
            {
                var serviceName = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/div/form/div/div[1]/div[1]/div[1]/div[1]/div/input"));
                serviceName.SetValue(ControlType.TextBox, "test service" + code);
            }

            var serviceSaveButton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/div/form/div/div[2]/div/button[3]"));
            serviceSaveButton.Click();

            WaitForLoad();

            //go to price to add one
            var pricePageButton = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/ul/li[2]"));
            pricePageButton.Click();
            WaitForLoad();
            var createNewPriceButton = WaitForElementIsVisible(By.Id("btn-add-item-detail"));
            createNewPriceButton.Click();
            WaitForLoad();

            //customersListdiv.Cl;
            var customersListDiv = WaitForElementExists(By.XPath("/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div[1]/div[2]/div")); ;
            customersListDiv.Click();
            var customersListInput = WaitForElementExists(By.Id("dropdown-customer-selectized"));
            customersListInput.SendKeys((string)code);
            customersListInput.SendKeys(Keys.Enter);

            //var customerSearch = WaitForElementIsVisible(By.Id("dropdown-customer-selectized"));
            //customerSearch.SetValue(ControlType.DropDownList, code);
            //WaitForLoad();
            //customersList.SendKeys(Keys.Enter);

            var customerServiceCode = WaitForElementIsVisible(By.Id("item_CustomerSrvId"));
            customerServiceCode.SetValue(ControlType.TextBox, code);

            var dateTo = WaitForElementIsVisible(By.Id("end-date-picker"));
            dateTo.SetValue(ControlType.DateTime, DateTime.Now.AddYears(10));
            dateTo.SendKeys(Keys.Tab);

            var invoicePrice = WaitForElementIsVisible(By.Id("item_InvoicePrice"));
            invoicePrice.SetValue(ControlType.TextBox, "100");

            var priceSaveButton = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div/div/form/div[3]/button[@id=\"last\"]"));
            priceSaveButton.Click();



        }

        public ServiceMassiveDeleteModalPage ClickMassiveDelete()
        {
            ShowExtendedMenu();
            WaitForElementIsVisible(By.XPath(MASSIVE_DELETE)).Click();
            WaitForLoad();
            WaitPageLoading();
            return new ServiceMassiveDeleteModalPage(_webDriver, _testContext);
        }

        public bool VerifyServiceExist(string serviceCode)
        {
            Filter(FilterType.Search, serviceCode);
            var counter = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span"));
            if (counter.Text == "0")
                return true;
            return false;
        }

        private const string c_serviceUrlPath = "Customers/Service";
        public static ServicePage NavigateTo(IWebDriver webDriver, TestContext testContext)
        {
            var winrestUrl = testContext.Properties["Winrest_URL"].ToString();
            var serviceUrl = new Uri(new Uri(winrestUrl), c_serviceUrlPath);
            webDriver.Navigate().GoToUrl(serviceUrl);
            return new ServicePage(webDriver, testContext);
        }
        public ServiceGeneralInformationPage SelectFirstRow()
        {
            WaitForLoad();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            WaitPageLoading();
            return new ServiceGeneralInformationPage(_webDriver, _testContext);
        }

        public string GetShowFilterSelected()
        {

            var itemSelected = _webDriver.FindElements(By.XPath(SHOW_SELECTED));
            var showName = string.Empty;
            foreach (var element in itemSelected)
            {
                if (element.Selected)
                {
                    var text = element.GetAttribute("value");
                    switch (element.GetAttribute("value"))
                    {
                        case "All":
                            showName = "Show all";
                            break;
                        case "Active":
                            showName = "Show active only";
                            break;
                        case "Inactive":
                            showName = "Show inactive only";
                            break;
                        default:
                            showName = "";
                            break;
                        case "Generic":
                            showName = "Show only generic service";
                            break;

                        case "WithExpiredPrices":
                            showName = "Show only active service with expired prices";
                            break;

                    }
                }
            }

            return showName;
        }

        public int GetSelectedCustomersToFilter()
        {

            var customer = WaitForElementIsVisible(By.XPath("//*[@id=\"SelectedCustomers_ms\"]/span[2]"));
            string nbre1 = customer.Text.Split(' ')[0];
            //string nbre2 = customer.Text.Split(' ')[2];
            //int.Parse(nbre1).CompareTo(nbre2);
            return int.Parse(nbre1);
        }
        public List<string> GetSelectedCategoriesToFilter()
        {
            var categoriesSelected = new List<string>();
            _categories = WaitForElementIsVisible(By.Id(CATEGORIES));
            _categories.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(CATEGORIES_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementExists(By.XPath(string.Format("/html/body/div[10]/ul/li[{0}]/label/input/../span", i + 1)));
                    categoriesSelected.Add(element.Text);
                }
            }

            _categories.Click();
            return categoriesSelected;
        }

        public List<string> GetSelectedSiteToFilter()
        {
            var siteSelected = new List<string>();
            _sites = WaitForElementIsVisible(By.Id(SITES));
            _sites.Click();
            WaitForLoad();
            var selectElement = _webDriver.FindElements(By.XPath(SITES_SELECTED));
            for (var i = 0; i < selectElement.Count; i++)
            {
                if (selectElement[i].Selected)
                {
                    var element = WaitForElementExists(By.XPath(string.Format("/html/body/div[10]/ul/li[{0}]/label/input/../span", i + 1)));
                    siteSelected.Add(element.Text);
                }
            }

            _categories.Click();
            return siteSelected;
        }
        public ServicePricePage SelectFirstRowInCustmerService()
        {
            WaitForLoad();
            _firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW_IN_CUTERMER_SERVICE));
            _firstRow.Click();
            WaitForLoad();
            return new ServicePricePage(_webDriver, _testContext);
        }
        public int GetPricesItem()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEMS));
            return elements.Count;
        }
        public int GetActivePricesFirstService()
        {
            int result;
            var activePrices = _webDriver.FindElement(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[6]"));
            int.TryParse(activePrices.Text, out result);
            return result;
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
        public ServicePricePage SelectIndexService(int indicerow)
        {

         if (isElementVisible(By.XPath(String.Format(INDEXSERVICE, indicerow + 1))))
         {
            var serviceindex = _webDriver.FindElement(By.XPath(String.Format(INDEXSERVICE, indicerow + 1)));
            serviceindex.Click();
            WaitForLoad();
         }     
         return new ServicePricePage(_webDriver, _testContext);

        }
        public void CLearDownloadsFolder(string path)
        {
            DirectoryInfo taskDirectory = new DirectoryInfo(path);
            FileInfo[] taskFiles = taskDirectory.GetFiles();
            foreach (var file in taskFiles)
            {
                file.Delete();
            }
        }
        public bool IsDetailPageServiceLoaded()
        {
            WaitPageLoading();
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));

            try
            {
                IWebElement priceTab = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(PRICESERVICE))); // Example verification of the 'Price' tab

                return priceTab.Displayed;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        public bool IsGeneralInformationServiceLoaded()
        {
            WaitPageLoading();
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(10));

            try
            {
                IWebElement GeneralInfoTab = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(GeneralInfoSERVICE))); // Example verification of the 'General Information' tab

                return GeneralInfoTab.Displayed;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }
        public bool CheckServiceHeadTableIHM()
        {
            _service_header = WaitForElementIsVisible(By.XPath(SERVICEHEADER));
            var xServiceHeader = _service_header.Location.X;
            var listServices = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[2]/a"));
            var resultService = listServices.Any(e => Math.Abs(xServiceHeader - e.Location.X) > 1);
            return resultService;
        }
        public bool CheckCategoryHeadTableIHM()
        {
            _category_header = WaitForElementIsVisible(By.XPath(CATEGORYHEADER));
            var xCategorieHeader = _category_header.Location.X;
            var listCategory = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]/a"));
            var resultCategorie = listCategory.Any(e=> Math.Abs(xCategorieHeader - e.Location.X) > 1);
            return resultCategorie;
        }
        public bool CheckActivePriceHeadTableIHM()
        {
            _activeprice_header = WaitForElementIsVisible(By.XPath(ACTIVEPRICEHEADER));
            var xActivepriceHeader = _activeprice_header.Location.X;
            var listActiveprice = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[6]"));
            var resultActiveprice = listActiveprice.Any(e => Math.Abs(xActivepriceHeader - e.Location.X) > 1);
            return resultActiveprice;
        }
        public bool CheckMenuHeadTableIHM()
        {
            _menu_header = WaitForElementIsVisible(By.XPath(MENUHEADER));
            var xMenuHeader = _menu_header.Location.X;
            var listMenu = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[7]"));
            var resultMenu = listMenu.Any(e => Math.Abs(xMenuHeader - e.Location.X) > 1);
            return resultMenu;
        }
        public bool CheckDeliveryHeadTableIHM()
        {
            _delivery_header = WaitForElementIsVisible(By.XPath(DELIVERYHEADER));
            var xDeliveryHeader = _delivery_header.Location.X;
            var listDelivery = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[8]"));
            var resultDelivery = listDelivery.Any(e => Math.Abs(xDeliveryHeader - e.Location.X) > 1);
            return resultDelivery;
        }
        public bool CheckLoadingPlanHeadTableIHM()
        {
            _loadingplan_header = WaitForElementIsVisible(By.XPath(LOADINGPLANHEADER));
            var xLoadingplanHeader = _loadingplan_header.Location.X;
            var listLoadingplan = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[9]"));
            var resultLoadingplan = listLoadingplan.Any(e => Math.Abs(xLoadingplanHeader - e.Location.X) > 1);
            return resultLoadingplan;
        }
    }
}